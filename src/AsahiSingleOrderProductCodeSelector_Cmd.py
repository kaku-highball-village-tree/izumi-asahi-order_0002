# -- coding: utf-8 --
###############################################################
#
# AsahiSingleOrderProductCodeSelector_Cmd.py
#
# pip install openpyxl
#
###############################################################

from __future__ import annotations

import csv
import os
import re
import sys
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


STEP_HEADERS: tuple[str, ...] = (
    "納品日",
    "曜日",
    "配送パターン",
    "便",
    "Ｐ品番",
    "APEX品番",
    "商品名",
    "産地",
    "仕様",
    "伝票原価",
    "伝票売価",
    "値入",
    "売価",
    "単位",
)
WEEKDAYS: tuple[str, ...] = ("月", "火", "水", "木", "金", "土", "日")
SUPPORTED_EXTENSIONS: set[str] = {".xlsx", ".tsv"}
STEP0007_MARKER: str = "_step0007_"
OUTPUT_PREFIX: str = "ProductCodeSelector_step0001_"


def normalize_cell(objValue: object, iColumn: int) -> str:
    """XLSXとTSVを同じ論理値として扱える文字列へ変換します。"""
    if objValue is None:
        return ""
    if iColumn == 0 and isinstance(objValue, datetime):
        return objValue.date().strftime("%Y/%m/%d")
    if iColumn == 0 and isinstance(objValue, date):
        return objValue.strftime("%Y/%m/%d")
    if isinstance(objValue, bool):
        return "TRUE" if objValue else "FALSE"
    if isinstance(objValue, float) and objValue.is_integer():
        return str(int(objValue))
    return str(objValue)


def validate_input_path(pszInputFileFullPath: str) -> Path:
    """商品別step0007のXLSXまたはTSV入力パスを検証します。"""
    objInputPath: Path = Path(pszInputFileFullPath).expanduser().resolve()
    if not objInputPath.is_file():
        raise ValueError("入力ファイルが見つかりません。Path = " + str(objInputPath))
    if objInputPath.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            "入力形式はXLSXまたはTSVではありません。Path = " + str(objInputPath)
        )
    if objInputPath.stem.count(STEP0007_MARKER) != 1:
        raise ValueError(
            "入力ファイル名には_step0007_が1つ必要です。Path = "
            + str(objInputPath)
        )
    return objInputPath


def read_excel_table(objInputPath: Path) -> tuple[list[list[str]], str]:
    """商品別step0007 XLSXの全セルとシート名を読み込みます。"""
    objWorkbook: Workbook = load_workbook(objInputPath, data_only=True)
    if len(objWorkbook.worksheets) != 1:
        raise ValueError("入力XLSXのワークシート数が1ではありません。")
    objWorksheet: Worksheet = objWorkbook.active
    listRows: list[list[str]] = [
        [
            normalize_cell(objWorksheet.cell(iRow, iColumn).value, iColumn - 1)
            for iColumn in range(1, objWorksheet.max_column + 1)
        ]
        for iRow in range(1, objWorksheet.max_row + 1)
    ]
    return listRows, objWorksheet.title


def read_tsv_table(objInputPath: Path) -> tuple[list[list[str]], str]:
    """商品別step0007 TSVをUTF-8として読み込みます。"""
    with objInputPath.open(mode="r", encoding="utf-8-sig", newline="") as objFile:
        listRawRows: list[list[str]] = list(
            csv.reader(objFile, delimiter="\t", strict=True)
        )
    listRows: list[list[str]] = [
        [normalize_cell(pszValue, iColumn) for iColumn, pszValue in enumerate(listRow)]
        for listRow in listRawRows
    ]
    return listRows, "ProductCodeSelector_step0001"


def validate_step0007_table(listRows: list[list[str]]) -> None:
    """商品別step0007の2行ヘッダーと月～日7行を検証します。"""
    if len(listRows) != 2 + len(WEEKDAYS):
        raise ValueError(
            "商品別step0007はヘッダー2行とデータ7行の合計9行ではありません。"
            + " 行数 = "
            + str(len(listRows))
        )
    iColumnCount: int = len(listRows[0])
    if iColumnCount < len(STEP_HEADERS):
        raise ValueError("商品別step0007の列数が14列未満です。")
    for iRow, listRow in enumerate(listRows, start=1):
        if len(listRow) != iColumnCount:
            raise ValueError(f"商品別step0007の{iRow}行目の列数が一致しません。")
    if any(pszValue.strip() for pszValue in listRows[0][: len(STEP_HEADERS)]):
        raise ValueError("商品別step0007の1行目A～N列が空欄ではありません。")
    if tuple(pszValue.strip() for pszValue in listRows[1][: len(STEP_HEADERS)]) != STEP_HEADERS:
        raise ValueError("商品別step0007の2行目A～N列が仕様どおりではありません。")
    listStoreCodes: list[str] = [
        pszValue.strip() for pszValue in listRows[0][len(STEP_HEADERS) :]
    ]
    if any(not pszCode for pszCode in listStoreCodes):
        raise ValueError("商品別step0007の店舗コードに空欄があります。")
    if len(set(listStoreCodes)) != len(listStoreCodes):
        raise ValueError("商品別step0007の店舗コードが重複しています。")

    listDataRows: list[list[str]] = listRows[2:]
    tupleProductValues: tuple[str, str, str] = tuple(
        listDataRows[0][iColumn].strip() for iColumn in (4, 5, 6)
    )
    if not tupleProductValues[2]:
        raise ValueError("商品別step0007の商品名が空欄です。")
    for iDay, (listRow, pszExpectedWeekday) in enumerate(
        zip(listDataRows, WEEKDAYS)
    ):
        iFileRow: int = iDay + 3
        if listRow[1].strip() != pszExpectedWeekday:
            raise ValueError(
                f"商品別step0007の{iFileRow}行目の曜日が{pszExpectedWeekday}ではありません。"
            )
        if tuple(listRow[iColumn].strip() for iColumn in (4, 5, 6)) != tupleProductValues:
            raise ValueError(
                f"商品別step0007の{iFileRow}行目の商品情報が月曜日行と一致しません。"
            )
    try:
        objMonday: date = datetime.strptime(listDataRows[0][0], "%Y/%m/%d").date()
    except ValueError as objException:
        raise ValueError("商品別step0007の月曜日日付が不正です。") from objException
    if objMonday.weekday() != 0:
        raise ValueError("商品別step0007の開始日が月曜日ではありません。")
    for iDay, listRow in enumerate(listDataRows):
        try:
            objActualDate: date = datetime.strptime(listRow[0], "%Y/%m/%d").date()
        except ValueError as objException:
            raise ValueError(
                f"商品別step0007の{iDay + 3}行目の日付が不正です。"
            ) from objException
        if objActualDate != objMonday + timedelta(days=iDay):
            raise ValueError("商品別step0007の日付が月～日の連続日付ではありません。")
    if not any(
        pszValue.strip()
        for listRow in listDataRows
        for pszValue in listRow[len(STEP_HEADERS) :]
    ):
        raise ValueError("商品別step0007に発注数量がありません。")


def clear_product_codes(listRows: list[list[str]]) -> list[list[str]]:
    """データ7行のＰ品番とAPEX品番を空欄にしたコピーを返します。"""
    listOutputRows: list[list[str]] = [listRow.copy() for listRow in listRows]
    for listRow in listOutputRows[2:]:
        listRow[4] = ""
        listRow[5] = ""
    return listOutputRows


def sanitize_filename_part(pszValue: str) -> str:
    """Windowsで安全なファイル名部分へ変換します。"""
    pszSafeValue: str = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", pszValue.strip())
    pszSafeValue = re.sub(r"_+", "_", pszSafeValue).rstrip(" .")
    if not pszSafeValue:
        raise ValueError("出力ファイル名を作成できません。")
    return pszSafeValue


def get_output_paths(objInputPath: Path, pszProductName: str) -> tuple[Path, Path]:
    """入力名からProductCodeSelector step0001の出力パスを作ります。"""
    _, pszIdentity = objInputPath.stem.split(STEP0007_MARKER, 1)
    if "_" not in pszIdentity:
        raise ValueError("step0007のファイル名に商品コード以降の情報がありません。")
    _, pszAfterCode = pszIdentity.split("_", 1)
    pszSafeProductName: str = sanitize_filename_part(pszProductName)
    pszExpectedProductPrefix: str = pszSafeProductName + "_"
    if not pszAfterCode.startswith(pszExpectedProductPrefix):
        raise ValueError(
            "step0007のファイル名の商品名が入力表の商品名と一致しません。"
        )
    pszSuffix: str = pszAfterCode[len(pszExpectedProductPrefix) :]
    if not pszSuffix:
        raise ValueError("step0007のファイル名に配送センター情報がありません。")
    pszOutputStem: str = OUTPUT_PREFIX + sanitize_filename_part(
        pszSafeProductName + "_" + pszSuffix
    )
    return (
        objInputPath.with_name(pszOutputStem + ".xlsx"),
        objInputPath.with_name(pszOutputStem + ".tsv"),
    )


def save_excel_table(
    objOutputPath: Path, listRows: list[list[str]], pszWorksheetTitle: str
) -> None:
    """ProductCodeSelector step0001 XLSXを保存します。"""
    objWorkbook: Workbook = Workbook()
    objWorksheet: Worksheet = objWorkbook.active
    objWorksheet.title = pszWorksheetTitle
    for iRow, listRow in enumerate(listRows, start=1):
        listValues: list[object] = listRow.copy()
        if iRow >= 3 and listValues[0]:
            listValues[0] = datetime.strptime(str(listValues[0]), "%Y/%m/%d").date()
        for iColumn in range(len(STEP_HEADERS), len(listValues)):
            if iRow >= 3 and str(listValues[iColumn]).isdigit():
                listValues[iColumn] = int(str(listValues[iColumn]))
        objWorksheet.append(listValues)
        if iRow == 1:
            for iColumn in range(len(STEP_HEADERS) + 1, len(listValues) + 1):
                objWorksheet.cell(iRow, iColumn).number_format = "@"
        elif iRow >= 3:
            objWorksheet.cell(iRow, 1).number_format = "yyyy/mm/dd"
            objWorksheet.cell(iRow, 5).number_format = "@"
            objWorksheet.cell(iRow, 6).number_format = "@"
    objWorkbook.save(objOutputPath)


def save_tsv_table(objOutputPath: Path, listRows: list[list[str]]) -> None:
    """ProductCodeSelector step0001 TSVを保存します。"""
    with objOutputPath.open(mode="w", encoding="utf-8", newline="") as objFile:
        csv.writer(objFile, delimiter="\t", lineterminator="\r\n").writerows(listRows)


def validate_outputs_match(objExcelPath: Path, objTsvPath: Path) -> None:
    """保存したXLSXとTSVの全セルが一致し、商品コードが空欄か確認します。"""
    listExcelRows, _ = read_excel_table(objExcelPath)
    listTsvRows, _ = read_tsv_table(objTsvPath)
    if listExcelRows != listTsvRows:
        raise ValueError("step0001のXLSXとTSVの内容が一致しません。")
    if any(listRow[4].strip() or listRow[5].strip() for listRow in listExcelRows[2:]):
        raise ValueError("step0001のＰ品番またはAPEX品番が空欄ではありません。")


def create_temporary_path(objOutputPath: Path) -> Path:
    """出力と同じフォルダーに一意な一時パスを作ります。"""
    iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
        prefix="." + objOutputPath.stem + ".", suffix=objOutputPath.suffix,
        dir=objOutputPath.parent,
    )
    os.close(iFileDescriptor)
    os.unlink(pszTemporaryPath)
    return Path(pszTemporaryPath)


def replace_output_pair(
    objTemporaryExcelPath: Path,
    objTemporaryTsvPath: Path,
    objExcelOutputPath: Path,
    objTsvOutputPath: Path,
) -> None:
    """XLSXとTSVをまとめて置換し、失敗時には以前の出力へ戻します。"""
    dictBackups: dict[Path, Path] = {}
    listReplaced: list[Path] = []
    try:
        for objOutputPath in (objExcelOutputPath, objTsvOutputPath):
            if objOutputPath.exists():
                objBackupPath = create_temporary_path(objOutputPath)
                os.replace(objOutputPath, objBackupPath)
                dictBackups[objOutputPath] = objBackupPath
        for objTemporaryPath, objOutputPath in (
            (objTemporaryExcelPath, objExcelOutputPath),
            (objTemporaryTsvPath, objTsvOutputPath),
        ):
            os.replace(objTemporaryPath, objOutputPath)
            listReplaced.append(objOutputPath)
    except Exception:
        for objOutputPath in reversed(listReplaced):
            if objOutputPath.exists():
                objOutputPath.unlink()
        for objOutputPath, objBackupPath in dictBackups.items():
            if objBackupPath.exists():
                os.replace(objBackupPath, objOutputPath)
        raise
    finally:
        for objBackupPath in dictBackups.values():
            if objBackupPath.exists():
                objBackupPath.unlink()


def get_error_path(objInputPath: Path) -> Path:
    """入力ファイルを基準にエラーテキストのパスを返します。"""
    return objInputPath.with_name(objInputPath.stem + "_error.txt")


def write_error_text(objErrorPath: Path, pszErrorMessage: str) -> None:
    """処理エラーをUTF-8テキストで保存します。"""
    pszText: str = (
        "処理名:\nProductCodeSelector step0001\n\n"
        + "エラー:\n"
        + pszErrorMessage
        + "\n"
    )
    objErrorPath.write_text(pszText, encoding="utf-8")


def process_input_file(pszInputFileFullPath: str) -> tuple[Path, Path, str]:
    """step0007からＰ品番とAPEX品番が空欄のstep0001を作成します。"""
    objInputPath: Path = validate_input_path(pszInputFileFullPath)
    if objInputPath.suffix.lower() == ".xlsx":
        listRows, pszWorksheetTitle = read_excel_table(objInputPath)
    else:
        listRows, pszWorksheetTitle = read_tsv_table(objInputPath)
    validate_step0007_table(listRows)
    pszProductName: str = listRows[2][6].strip()
    objExcelOutputPath, objTsvOutputPath = get_output_paths(
        objInputPath, pszProductName
    )
    listOutputRows: list[list[str]] = clear_product_codes(listRows)
    objTemporaryExcelPath: Path = create_temporary_path(objExcelOutputPath)
    objTemporaryTsvPath: Path = create_temporary_path(objTsvOutputPath)
    try:
        save_excel_table(objTemporaryExcelPath, listOutputRows, pszWorksheetTitle)
        save_tsv_table(objTemporaryTsvPath, listOutputRows)
        validate_outputs_match(objTemporaryExcelPath, objTemporaryTsvPath)
        replace_output_pair(
            objTemporaryExcelPath,
            objTemporaryTsvPath,
            objExcelOutputPath,
            objTsvOutputPath,
        )
    finally:
        for objTemporaryPath in (objTemporaryExcelPath, objTemporaryTsvPath):
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()
    return objExcelOutputPath, objTsvOutputPath, pszProductName


def parse_command_line_arguments() -> str:
    """1つの入力ファイルパスを解析します。"""
    if len(sys.argv) != 2 or sys.argv[1].startswith("--"):
        raise ValueError("入力ファイルパスは1つ指定してください。")
    return sys.argv[1]


def main() -> int:
    """引数を確認して処理し、成功0・失敗1の終了コードを返します。"""
    try:
        pszInputFileFullPath: str = parse_command_line_arguments()
    except ValueError as objException:
        pszMessage: str = (
            "Error: "
            + str(objException)
            + "\nUsage: python "
            + os.path.basename(__file__)
            + " <input_file_path>\n"
        )
        print(pszMessage, file=sys.stderr, end="")
        Path(os.path.splitext(os.path.basename(__file__))[0] + "_error_argument.txt").write_text(
            pszMessage, encoding="utf-8"
        )
        return 1
    try:
        objExcelPath, objTsvPath, pszProductName = process_input_file(
            pszInputFileFullPath
        )
    except Exception as objException:
        pszMessage = "Error: " + str(objException)
        print(pszMessage, file=sys.stderr)
        try:
            write_error_text(
                get_error_path(Path(pszInputFileFullPath).expanduser().resolve()),
                str(objException),
            )
        except OSError as objWriteException:
            print(
                "Error: エラーファイルを保存できません。Detail = "
                + str(objWriteException),
                file=sys.stderr,
            )
        return 1
    print("ProductCodeSelector step0001の作成が完了しました。")
    print("商品名: " + pszProductName)
    print("XLSX: " + str(objExcelPath))
    print("TSV: " + str(objTsvPath))
    return 0


if __name__ == "__main__":
    sys.exit(main())

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
import difflib
import os
import posixpath
import re
import shutil
import sys
import tempfile
import tkinter as tk
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
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
SOURCE_PRODUCTS_FILE_NAME: str = "products_all_109_readable.tsv"
PRODUCTS_FILE_NAME: str = "products_all_109_readable_ABC.tsv"
WEEKLY_TEMPLATE_FILE_NAME: str = "templete_イズミ週間予定表.xlsx"
WEEKLY_SHEET_NAME: str = "センター週間"
WEEKLY_TSV_MAX_ROW: int = 42
WEEKLY_TSV_MAX_COLUMN: int = 28
WEEKLY_TSV_RANGE: str = (
    "A1:" + get_column_letter(WEEKLY_TSV_MAX_COLUMN) + str(WEEKLY_TSV_MAX_ROW)
)
AREA_STORE_MAPPING_FILE_NAME: str = "AsahiOrderAreaStoreMapping_対応表.txt"
AREA_NAMES: tuple[str, str, str] = ("広島", "岡山", "四国／岡山")
SHIPMENT_DATE_ROW_RANGES: tuple[tuple[int, int], ...] = ((8, 4), (8, 13), (8, 22))
DELIVERY_DATE_ROW_RANGES: tuple[tuple[int, int], ...] = ((10, 4), (10, 13), (10, 22))
SHIPMENT_WEEKDAY_ROW_RANGES: tuple[tuple[int, int], ...] = ((9, 4), (9, 13), (9, 22))
DELIVERY_WEEKDAY_ROW_RANGES: tuple[tuple[int, int], ...] = ((11, 4), (11, 13), (11, 22))
STEP0004_AREA_RANGES: tuple[tuple[str, int, int, int, int], ...] = (
    ("広島", 12, 42, 2, 10),
    ("岡山", 12, 42, 11, 19),
    ("四国", 12, 42, 20, 28),
)
STEP0004_MAX_STORES_PER_AREA: int = 31
PRODUCT_HEADERS: tuple[str, str, str] = ("productCode", "productName", "spec")
COLUMN_WIDTH_LIMITS: tuple[tuple[int, int], ...] = (
    (12, 14),
    (6, 8),
    (14, 24),
    (8, 12),
    (14, 22),
    (14, 22),
    (24, 50),
    (14, 30),
    (14, 30),
    (14, 18),
    (14, 18),
    (10, 16),
    (12, 18),
    (10, 18),
)
STORE_COLUMN_WIDTH_LIMITS: tuple[int, int] = (10, 24)
PRODUCT_CATEGORY_TERMS: dict[str, tuple[str, ...]] = {
    "まぐろ": ("まぐろ", "鮪", "本まぐろ", "本鮪", "クロマグロ", "黒まぐろ", "キハダ", "メバチ", "ビンナガ", "ビンチョウ", "トンボ", "ミナミマグロ", "南まぐろ", "インドまぐろ"),
    "かき": ("かき", "牡蠣", "真牡蠣", "マガキ"),
    "いか": ("いか", "烏賊", "紋甲いか", "モンゴウイカ", "あおりいか", "するめいか", "やりいか"),
    "ぶり": ("ぶり", "鰤", "はまち", "ワラサ", "イナダ"),
    "たい": ("たい", "鯛", "真鯛", "マダイ"),
    "えび": ("えび", "海老", "蝦"),
    "かに": ("かに", "蟹", "ずわいがに", "ずわい蟹"),
    "たら": ("たら", "鱈", "真たら", "真鱈", "マダラ"),
}
PRODUCT_ATTRIBUTE_TERMS: dict[str, tuple[str, ...]] = {
    "冷凍": ("冷凍",), "生": ("生",), "ボイル": ("ボイル", "ゆで", "茹で"),
    "蒸し": ("蒸し",), "塩": ("塩",), "熟成": ("熟成",),
    "ロイン": ("ロイン",), "フィレ": ("フィレ", "フィーレ", "フィレット"),
    "切り落とし": ("切り落とし", "切落し"), "カマ": ("カマ",),
    "ほほ肉": ("ほほ肉", "頬肉"), "頭肉": ("頭肉",), "あら": ("あら",),
    "ラウンド": ("ラウンド",), "スキンレス": ("スキンレス", "皮なし", "皮無し"),
    "骨取り": ("骨取り", "骨とり"), "刺身用": ("刺身用", "生食用"),
    "加熱用": ("加熱用",), "加工品用": ("加工品用",), "原料": ("原料",),
    "養殖": ("養殖",), "天然": ("天然",), "MEL認証": ("MEL認証",),
    "代用品": ("代用品",),
}
PRODUCT_ORIGIN_TERMS: dict[str, tuple[str, ...]] = {
    "北海道": ("北海道",), "千葉": ("千葉県", "千葉県産", "千葉産"),
    "鳥取": ("鳥取県", "鳥取県産", "鳥取産"),
    "岡山": ("岡山県", "岡山県産", "岡山産"),
    "広島": ("広島県", "広島県産", "広島産"),
    "徳島": ("徳島県", "徳島県産", "徳島産"),
    "愛媛": ("愛媛県", "愛媛県産", "愛媛産"),
    "高知": ("高知県", "高知県産", "高知産"),
    "長崎": ("長崎県", "長崎県産", "長崎産"),
    "鹿児島": ("鹿児島県", "鹿児島県産", "鹿児島産"),
    "和歌山": ("和歌山県", "和歌山県産", "和歌山産"),
    "瀬戸内": ("瀬戸内", "瀬戸内産"),
    "国産": ("国産", "国内産"), "カナダ": ("カナダ", "カナダ産"),
}
PRODUCT_BRAND_TERMS: dict[str, tuple[str, ...]] = {
    "日の出": ("日の出", "日の出まぐろ"),
    "地元めし": ("地元めし", "瀬戸内地元めし"),
}
WEAK_PRODUCT_ATTRIBUTE_GROUPS: frozenset[str] = frozenset(
    {"冷凍", "生", "原料", "養殖", "天然"}
)
UNKNOWN_CATEGORY_SIMILARITY_THRESHOLD: float = 0.60


class SelectionCancelledError(Exception):
    """商品選択がユーザー操作によってキャンセルされたことを表します。"""


class NoMatchingProductError(Exception):
    """担当者が商品マスターに該当商品なしと判断したことを表します。"""


class ProductCandidate:
    """商品マスターの1商品を保持します。"""

    def __init__(self, pszCode: str, pszName: str, pszSpec: str) -> None:
        self.code: str = pszCode
        self.name: str = pszName
        self.spec: str = pszSpec

    @property
    def display_text(self) -> str:
        """プルダウンに表示する文字列を返します。"""
        pszText: str = self.code + " - " + self.name
        if self.spec:
            pszText += " - " + self.spec
        return pszText


def configure_standard_streams() -> None:
    """標準出力と標準エラーをUTF-8へ統一します。"""
    for objStream in (sys.stdout, sys.stderr):
        objReconfigure = getattr(objStream, "reconfigure", None)
        if callable(objReconfigure):
            objReconfigure(encoding="utf-8", errors="replace")


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
    """商品別のXLSXまたはTSV入力パスを検証します。"""
    objInputPath: Path = Path(pszInputFileFullPath).expanduser().resolve()
    if not objInputPath.is_file():
        raise ValueError("入力ファイルが見つかりません。Path = " + str(objInputPath))
    if objInputPath.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            "入力形式はXLSXまたはTSVではありません。Path = " + str(objInputPath)
        )
    if objInputPath.stem.count(STEP0007_MARKER) > 1:
        raise ValueError(
            "入力ファイル名に_step0007_を複数含めることはできません。Path = "
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
    # 商品基本情報は月曜日行だけに入り、火～日のC～N列は空欄またはメモです。
    if not listDataRows[0][6].strip():
        raise ValueError("商品別step0007の商品名が空欄です。")
    for iDay, (listRow, pszExpectedWeekday) in enumerate(
        zip(listDataRows, WEEKDAYS)
    ):
        iFileRow: int = iDay + 3
        if listRow[1].strip() != pszExpectedWeekday:
            raise ValueError(
                f"商品別step0007の{iFileRow}行目の曜日が{pszExpectedWeekday}ではありません。"
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
    iMarkerCount: int = objInputPath.stem.count(STEP0007_MARKER)
    if iMarkerCount == 0:
        pszOutputStem: str = OUTPUT_PREFIX + sanitize_filename_part(objInputPath.stem)
        return (
            objInputPath.with_name(pszOutputStem + ".xlsx"),
            objInputPath.with_name(pszOutputStem + ".tsv"),
        )
    if iMarkerCount > 1:
        raise ValueError(
            "入力ファイル名に_step0007_を複数含めることはできません。Path = "
            + str(objInputPath)
        )
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
    pszOutputStem = OUTPUT_PREFIX + sanitize_filename_part(
        pszSafeProductName + "_" + pszSuffix
    )
    return (
        objInputPath.with_name(pszOutputStem + ".xlsx"),
        objInputPath.with_name(pszOutputStem + ".tsv"),
    )


def calculate_display_width(objValue: object, iColumn: int) -> int:
    """全角文字を2、半角文字を1としてセルの最大行表示幅を返します。"""
    if objValue is None:
        return 0
    if iColumn == 1 and isinstance(objValue, datetime):
        pszValue: str = objValue.date().strftime("%Y/%m/%d")
    elif iColumn == 1 and isinstance(objValue, date):
        pszValue = objValue.strftime("%Y/%m/%d")
    else:
        pszValue = str(objValue)
    pszValue = pszValue.replace("\t", "    ")
    return max(
        (
            sum(
                2 if unicodedata.east_asian_width(pszCharacter) in ("W", "F") else 1
                for pszCharacter in pszLine
            )
            for pszLine in pszValue.splitlines() or [""]
        ),
        default=0,
    )


def adjust_excel_column_widths(objWorksheet: Worksheet) -> None:
    """セル内容と列別の最小・最大幅に基づいて全列幅を設定します。"""
    for iColumn in range(1, objWorksheet.max_column + 1):
        if iColumn <= len(COLUMN_WIDTH_LIMITS):
            iMinimumWidth, iMaximumWidth = COLUMN_WIDTH_LIMITS[iColumn - 1]
        else:
            iMinimumWidth, iMaximumWidth = STORE_COLUMN_WIDTH_LIMITS
        iContentWidth: int = max(
            calculate_display_width(
                objWorksheet.cell(iRow, iColumn).value, iColumn
            )
            for iRow in range(1, objWorksheet.max_row + 1)
        )
        iColumnWidth: int = min(iMaximumWidth, max(iMinimumWidth, iContentWidth + 2))
        objWorksheet.column_dimensions[get_column_letter(iColumn)].width = iColumnWidth


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
    adjust_excel_column_widths(objWorksheet)
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


def get_source_products_file_path() -> Path:
    """Cmdプログラムと同じフォルダーの元商品マスターパスを返します。"""
    return Path(__file__).resolve().parent / SOURCE_PRODUCTS_FILE_NAME


def get_products_file_path() -> Path:
    """Cmdプログラムと同じフォルダーの3列商品マスターパスを返します。"""
    return Path(__file__).resolve().parent / PRODUCTS_FILE_NAME


def get_weekly_template_file_path() -> Path:
    """Cmdプログラムと同じフォルダーの週間予定表パスを返します。"""
    return Path(__file__).resolve().parent / WEEKLY_TEMPLATE_FILE_NAME


def create_abc_product_master(
    objSourcePath: Path, objOutputPath: Path
) -> None:
    """元商品マスターのA～C列だけを抽出したTSVを安全に作成します。"""
    if not objSourcePath.is_file():
        raise ValueError("元商品マスターが見つかりません。Path = " + str(objSourcePath))
    with objSourcePath.open(mode="r", encoding="utf-8-sig", newline="") as objFile:
        listSourceRows: list[list[str]] = list(
            csv.reader(objFile, delimiter="\t", strict=True)
        )
    listSourceRows = [
        listRow
        for listRow in listSourceRows
        if any(pszValue.strip() for pszValue in listRow)
    ]
    if not listSourceRows:
        raise ValueError("元商品マスターが空です。")
    if len(listSourceRows[0]) < len(PRODUCT_HEADERS):
        raise ValueError("元商品マスターの列数が3列未満です。")
    if tuple(pszValue.strip() for pszValue in listSourceRows[0][:3]) != PRODUCT_HEADERS:
        raise ValueError(
            "元商品マスターの先頭3列はproductCode、productName、specではありません。"
        )
    listOutputRows: list[list[str]] = [list(PRODUCT_HEADERS)]
    for iRow, listRow in enumerate(listSourceRows[1:], start=2):
        if len(listRow) < len(PRODUCT_HEADERS):
            raise ValueError(f"元商品マスターの{iRow}行目が3列未満です。")
        pszCode, pszName, pszSpec = (pszValue.strip() for pszValue in listRow[:3])
        if not pszCode or not pszName:
            raise ValueError(
                f"元商品マスターの{iRow}行目の商品コードまたは商品名が空欄です。"
            )
        listOutputRows.append([pszCode, pszName, pszSpec])

    objTemporaryPath: Path = create_temporary_path(objOutputPath)
    try:
        save_tsv_table(objTemporaryPath, listOutputRows)
        with objTemporaryPath.open(
            mode="r", encoding="utf-8-sig", newline=""
        ) as objFile:
            listSavedRows: list[list[str]] = list(
                csv.reader(objFile, delimiter="\t", strict=True)
            )
        if listSavedRows != listOutputRows:
            raise ValueError(
                "products_all_109_readable_ABC.tsvの保存内容が元商品マスターのA～C列と一致しません。"
            )
        read_product_candidates(objTemporaryPath)
        os.replace(objTemporaryPath, objOutputPath)
    finally:
        if objTemporaryPath.exists():
            objTemporaryPath.unlink()


def read_product_candidates(objProductsPath: Path) -> list[ProductCandidate]:
    """商品マスターを読み込み、商品コードの矛盾を検証します。"""
    if not objProductsPath.is_file():
        raise ValueError("商品マスターが見つかりません。Path = " + str(objProductsPath))
    with objProductsPath.open(mode="r", encoding="utf-8-sig", newline="") as objFile:
        listRows: list[list[str]] = list(csv.reader(objFile, delimiter="\t", strict=True))
    listRows = [listRow for listRow in listRows if any(pszValue.strip() for pszValue in listRow)]
    if not listRows:
        raise ValueError("商品マスターが空です。")
    if tuple(pszValue.strip() for pszValue in listRows[0]) != PRODUCT_HEADERS:
        raise ValueError("商品マスターのヘッダーがproductCode、productName、specではありません。")
    listCandidates: list[ProductCandidate] = []
    dictCodes: dict[str, tuple[str, str]] = {}
    for iRow, listRow in enumerate(listRows[1:], start=2):
        if len(listRow) != 3:
            raise ValueError(f"商品マスターの{iRow}行目が3列ではありません。")
        pszCode, pszName, pszSpec = (pszValue.strip() for pszValue in listRow)
        if not pszCode or not pszName:
            raise ValueError(f"商品マスターの{iRow}行目の商品コードまたは商品名が空欄です。")
        tupleDefinition: tuple[str, str] = (pszName, pszSpec)
        if pszCode in dictCodes:
            if dictCodes[pszCode] != tupleDefinition:
                raise ValueError("商品マスターの商品コード定義が矛盾しています。Code = " + pszCode)
            continue
        dictCodes[pszCode] = tupleDefinition
        listCandidates.append(ProductCandidate(pszCode, pszName, pszSpec))
    if not listCandidates:
        raise ValueError("商品マスターに商品がありません。")
    return listCandidates


def normalize_product_name(pszValue: str) -> str:
    """商品名を候補検索用に正規化します。"""
    pszNormalized: str = unicodedata.normalize("NFKC", pszValue).strip().casefold()
    listCharacters: list[str] = []
    for pszCharacter in pszNormalized:
        iCodePoint: int = ord(pszCharacter)
        if 0x30A1 <= iCodePoint <= 0x30F6:
            listCharacters.append(chr(iCodePoint - 0x60))
        else:
            listCharacters.append(pszCharacter)
    return " ".join("".join(listCharacters).replace("\u3000", " ").split())


def compact_product_text(pszValue: str) -> str:
    """空白と記号を除いた候補比較用文字列を返します。"""
    return "".join(
        pszCharacter
        for pszCharacter in normalize_product_name(pszValue)
        if pszCharacter.isalnum()
    )


def extract_term_groups(
    pszValue: str, dictTermGroups: dict[str, tuple[str, ...]]
) -> set[str]:
    """文字列に含まれるカテゴリまたは属性のグループ名を返します。"""
    pszCompactValue: str = compact_product_text(pszValue)
    return {
        pszGroup
        for pszGroup, tupleTerms in dictTermGroups.items()
        if any(compact_product_text(pszTerm) in pszCompactValue for pszTerm in tupleTerms)
    }


def evaluate_product_candidate(
    objCandidate: ProductCandidate, pszProductName: str, pszInputSpec: str
) -> tuple[int, list[str], list[str]]:
    """候補の関連度、一致理由、主な相違点を返します。"""
    pszInputNormalized: str = normalize_product_name(pszProductName)
    pszCandidateNormalized: str = normalize_product_name(objCandidate.name)
    pszInputCompact: str = compact_product_text(pszProductName)
    pszCandidateCompact: str = compact_product_text(objCandidate.name)
    iScore: int = 0
    listReasons: list[str] = []
    listDifferences: list[str] = []
    if objCandidate.name.strip() == pszProductName.strip():
        iScore += 1000
        listReasons.append("完全一致")
    elif pszCandidateNormalized == pszInputNormalized:
        iScore += 900
        listReasons.append("正規化一致")
    elif pszCandidateCompact == pszInputCompact:
        iScore += 800
        listReasons.append("空白・記号除去一致")
    elif pszInputCompact in pszCandidateCompact or pszCandidateCompact in pszInputCompact:
        iScore += 400
        listReasons.append("名称包含")

    setInputCategories: set[str] = extract_term_groups(
        pszProductName, PRODUCT_CATEGORY_TERMS
    )
    setCandidateCategories: set[str] = extract_term_groups(
        objCandidate.name, PRODUCT_CATEGORY_TERMS
    )
    setSharedCategories: set[str] = setInputCategories & setCandidateCategories
    if setSharedCategories:
        iScore += 300 * len(setSharedCategories)
        listReasons.append("魚介カテゴリ一致:" + ",".join(sorted(setSharedCategories)))

    setInputAttributes: set[str] = extract_term_groups(
        pszProductName, PRODUCT_ATTRIBUTE_TERMS
    )
    setCandidateAttributes: set[str] = extract_term_groups(
        objCandidate.name, PRODUCT_ATTRIBUTE_TERMS
    )
    setSharedAttributes: set[str] = setInputAttributes & setCandidateAttributes
    if setSharedAttributes:
        iScore += 100 * len(setSharedAttributes)
        listReasons.append("属性一致:" + ",".join(sorted(setSharedAttributes)))
    for pszDifference in sorted(setInputAttributes ^ setCandidateAttributes):
        listDifferences.append("属性差:" + pszDifference)

    for pszLabel, dictTerms, iPoints in (
        ("産地", PRODUCT_ORIGIN_TERMS, 100),
        ("ブランド", PRODUCT_BRAND_TERMS, 80),
    ):
        setInputTerms: set[str] = extract_term_groups(pszProductName, dictTerms)
        setCandidateTerms: set[str] = extract_term_groups(objCandidate.name, dictTerms)
        setSharedTerms: set[str] = setInputTerms & setCandidateTerms
        if setSharedTerms:
            iScore += iPoints * len(setSharedTerms)
            listReasons.append(pszLabel + "一致:" + ",".join(sorted(setSharedTerms)))
        if setInputTerms and setCandidateTerms and not setSharedTerms:
            listDifferences.append(pszLabel + "差")

    pszInputSpecNormalized: str = compact_product_text(pszInputSpec)
    pszCandidateSpecNormalized: str = compact_product_text(objCandidate.spec)
    if pszInputSpecNormalized and pszCandidateSpecNormalized:
        if pszInputSpecNormalized == pszCandidateSpecNormalized:
            iScore += 150
            listReasons.append("仕様一致")
        else:
            listDifferences.append("仕様差")
    fSimilarity: float = difflib.SequenceMatcher(
        None, pszInputCompact, pszCandidateCompact
    ).ratio()
    iScore += round(fSimilarity * 200)
    if fSimilarity >= 0.6 and not any(
        pszReason in listReasons
        for pszReason in ("完全一致", "正規化一致", "空白・記号除去一致")
    ):
        listReasons.append("名称類似")
    return iScore, listReasons, listDifferences


def find_product_candidates(
    listCandidates: list[ProductCandidate], pszProductName: str, pszInputSpec: str = ""
) -> list[ProductCandidate]:
    """完全一致で終了せず、関連する可能性がある商品を関連度順に返します。"""
    listEvaluated: list[tuple[int, ProductCandidate]] = []
    setInputCategories: set[str] = extract_term_groups(
        pszProductName, PRODUCT_CATEGORY_TERMS
    )
    pszInputCompact: str = compact_product_text(pszProductName)
    for objCandidate in listCandidates:
        iScore, listReasons, _ = evaluate_product_candidate(
            objCandidate, pszProductName, pszInputSpec
        )
        setCandidateCategories: set[str] = extract_term_groups(
            objCandidate.name, PRODUCT_CATEGORY_TERMS
        )
        fSimilarity: float = difflib.SequenceMatcher(
            None, pszInputCompact, compact_product_text(objCandidate.name)
        ).ratio()
        if setInputCategories:
            bIsRelated: bool = bool(setInputCategories & setCandidateCategories)
        else:
            pszCandidateNormalized: str = normalize_product_name(objCandidate.name)
            pszCandidateCompact: str = compact_product_text(objCandidate.name)
            bStrongNameMatch: bool = (
                objCandidate.name.strip() == pszProductName.strip()
                or pszCandidateNormalized == normalize_product_name(pszProductName)
                or pszCandidateCompact == pszInputCompact
                or pszInputCompact in pszCandidateCompact
                or pszCandidateCompact in pszInputCompact
            )
            setSharedStrongAttributes: set[str] = (
                extract_term_groups(pszProductName, PRODUCT_ATTRIBUTE_TERMS)
                & extract_term_groups(objCandidate.name, PRODUCT_ATTRIBUTE_TERMS)
            ) - WEAK_PRODUCT_ATTRIBUTE_GROUPS
            bSharedOriginOrBrand: bool = any(
                extract_term_groups(pszProductName, dictTerms)
                & extract_term_groups(objCandidate.name, dictTerms)
                for dictTerms in (PRODUCT_ORIGIN_TERMS, PRODUCT_BRAND_TERMS)
            )
            pszInputSpecNormalized: str = compact_product_text(pszInputSpec)
            pszCandidateSpecNormalized: str = compact_product_text(objCandidate.spec)
            bMatchingSpec: bool = bool(
                pszInputSpecNormalized
                and pszCandidateSpecNormalized
                and pszInputSpecNormalized == pszCandidateSpecNormalized
            )
            bIsRelated = bool(
                bStrongNameMatch
                or fSimilarity >= UNKNOWN_CATEGORY_SIMILARITY_THRESHOLD
                or setSharedStrongAttributes
                or bSharedOriginOrBrand
                or bMatchingSpec
            )
        if bIsRelated:
            listEvaluated.append((iScore, objCandidate))
    listEvaluated.sort(key=lambda tupleItem: (-tupleItem[0], tupleItem[1].code))
    return [objCandidate for _, objCandidate in listEvaluated]


def select_product_candidate(
    pszProductName: str,
    pszInputSpec: str,
    listRelatedCandidates: list[ProductCandidate],
    listAllCandidates: list[ProductCandidate],
    pszSundayValue: str,
) -> ProductCandidate:
    """関連候補・全商品を検索できる画面から担当者が選択します。"""
    objRoot = tk.Tk()
    objRoot.title("Asahi Single Order Product Code Selector step0002")
    objRoot.resizable(True, False)
    objSelectedCandidate: ProductCandidate | None = None
    bNoMatchingProduct: bool = False
    listVisibleCandidates: list[ProductCandidate] = []
    objFrame = ttk.Frame(objRoot, padding=12)
    objFrame.grid(row=0, column=0, sticky="nsew")
    ttk.Label(objFrame, text="入力商品名:").grid(row=0, column=0, sticky="w")
    ttk.Label(objFrame, text=pszProductName).grid(row=1, column=0, columnspan=3, sticky="w")
    ttk.Label(objFrame, text="入力仕様:").grid(row=2, column=0, sticky="w")
    ttk.Label(objFrame, text=pszInputSpec or "(空欄)").grid(
        row=3, column=0, columnspan=3, sticky="w", pady=(0, 8)
    )
    ttk.Label(objFrame, text="検索:").grid(row=4, column=0, sticky="w")
    objSearch = tk.StringVar()
    objSearchEntry = ttk.Entry(objFrame, textvariable=objSearch, width=45)
    objSearchEntry.grid(row=4, column=1, columnspan=2, sticky="ew")
    ttk.Label(objFrame, text="表示対象:").grid(row=5, column=0, sticky="w")
    objMode = tk.StringVar(value="推奨・関連候補")
    objModeComboBox = ttk.Combobox(
        objFrame,
        textvariable=objMode,
        values=("推奨・関連候補", "同じ魚介カテゴリ", "全商品"),
        state="readonly",
        width=20,
    )
    objModeComboBox.grid(row=5, column=1, sticky="w")
    objCount = tk.StringVar()
    ttk.Label(objFrame, textvariable=objCount).grid(row=5, column=2, sticky="e")
    ttk.Label(objFrame, text="商品候補:").grid(row=6, column=0, sticky="w")
    objSelection = tk.StringVar()
    objComboBox = ttk.Combobox(
        objFrame, textvariable=objSelection, state="readonly", width=90,
    )
    objComboBox.grid(row=7, column=0, columnspan=3, sticky="ew")
    objDetails = tk.StringVar()
    ttk.Label(objFrame, textvariable=objDetails, wraplength=760).grid(
        row=8, column=0, columnspan=3, sticky="w", pady=(6, 12)
    )

    def update_candidate_details(*_objArguments: object) -> None:
        iSelectedIndex: int = objComboBox.current()
        if iSelectedIndex < 0 or iSelectedIndex >= len(listVisibleCandidates):
            objDetails.set("")
            return
        _, listReasons, listDifferences = evaluate_product_candidate(
            listVisibleCandidates[iSelectedIndex], pszProductName, pszInputSpec
        )
        pszDetails: str = "一致理由: " + (", ".join(listReasons) or "全商品から表示")
        if listDifferences:
            pszDetails += "\n相違点: " + ", ".join(listDifferences)
        objDetails.set(pszDetails)

    def refresh_candidates(*_objArguments: object) -> None:
        nonlocal listVisibleCandidates
        listBaseCandidates: list[ProductCandidate]
        if objMode.get() == "全商品":
            listBaseCandidates = listAllCandidates
        elif objMode.get() == "同じ魚介カテゴリ":
            setInputCategories: set[str] = extract_term_groups(
                pszProductName, PRODUCT_CATEGORY_TERMS
            )
            listBaseCandidates = [
                objCandidate
                for objCandidate in listAllCandidates
                if setInputCategories
                & extract_term_groups(objCandidate.name, PRODUCT_CATEGORY_TERMS)
            ]
        else:
            listBaseCandidates = listRelatedCandidates
        listSearchTerms: list[str] = [
            compact_product_text(pszTerm)
            for pszTerm in objSearch.get().split()
            if compact_product_text(pszTerm)
        ]
        listVisibleCandidates = [
            objCandidate
            for objCandidate in listBaseCandidates
            if all(
                pszTerm
                in compact_product_text(
                    objCandidate.code + " " + objCandidate.name + " " + objCandidate.spec
                )
                for pszTerm in listSearchTerms
            )
        ]
        listDisplayValues: list[str] = [
            objCandidate.display_text for objCandidate in listVisibleCandidates
        ]
        objComboBox.configure(values=listDisplayValues)
        objCount.set("候補: " + str(len(listVisibleCandidates)) + "件")
        if listDisplayValues:
            objComboBox.current(0)
        else:
            objSelection.set("")
        update_candidate_details()

    def confirm_selection() -> None:
        nonlocal objSelectedCandidate
        iSelectedIndex: int = objComboBox.current()
        if iSelectedIndex < 0:
            messagebox.showerror("商品選択", "商品を選択してください。", parent=objRoot)
            return
        if pszSundayValue and not messagebox.askyesno(
            "商品名列の移動確認",
            "日曜日行の商品名セルに値があります。\n\n値: "
            + pszSundayValue
            + "\n\n商品名列を移動すると、この値はstep0002に残りません。\n続行しますか？",
            parent=objRoot,
        ):
            return
        objSelectedCandidate = listVisibleCandidates[iSelectedIndex]
        objRoot.destroy()

    def cancel_selection() -> None:
        objRoot.destroy()

    def select_no_matching_product() -> None:
        nonlocal bNoMatchingProduct
        if messagebox.askyesno(
            "該当商品なし",
            "商品マスターに該当商品がないものとしてstep0002を作成せず終了しますか？",
            parent=objRoot,
        ):
            bNoMatchingProduct = True
            objRoot.destroy()

    ttk.Button(objFrame, text="確定", command=confirm_selection).grid(row=9, column=0, sticky="e")
    ttk.Button(objFrame, text="該当商品なし", command=select_no_matching_product).grid(
        row=9, column=1
    )
    ttk.Button(objFrame, text="キャンセル", command=cancel_selection).grid(row=9, column=2, sticky="w")
    objSearch.trace_add("write", refresh_candidates)
    objModeComboBox.bind("<<ComboboxSelected>>", refresh_candidates)
    objComboBox.bind("<<ComboboxSelected>>", update_candidate_details)
    objRoot.protocol("WM_DELETE_WINDOW", cancel_selection)
    refresh_candidates()
    objSearchEntry.focus_set()
    objRoot.mainloop()
    if bNoMatchingProduct:
        raise NoMatchingProductError("商品マスターに該当商品なしが選択されました。")
    if objSelectedCandidate is None:
        raise SelectionCancelledError("商品選択がキャンセルされました。")
    return objSelectedCandidate


def get_step0002_output_paths(
    objStep0001ExcelPath: Path, objStep0001TsvPath: Path
) -> tuple[Path, Path]:
    """step0001の出力名からstep0002のXLSX・TSVパスを作ります。"""
    pszMarker: str = "ProductCodeSelector_step0001_"
    if not objStep0001ExcelPath.stem.startswith(pszMarker):
        raise ValueError("step0001の出力ファイル名ではありません。")
    pszStep0002Stem: str = objStep0001ExcelPath.stem.replace(
        pszMarker, "ProductCodeSelector_step0002_", 1
    )
    return (
        objStep0001ExcelPath.with_name(pszStep0002Stem + ".xlsx"),
        objStep0001TsvPath.with_name(pszStep0002Stem + ".tsv"),
    )


def get_step0003_output_paths(
    objStep0002ExcelPath: Path, objStep0002TsvPath: Path
) -> tuple[Path, Path]:
    """step0002の出力名からstep0003のXLSX・TSVパスを作ります。"""
    pszMarker: str = "ProductCodeSelector_step0002_"
    if not objStep0002ExcelPath.stem.startswith(pszMarker):
        raise ValueError("step0002の出力ファイル名ではありません。")
    if objStep0002ExcelPath.stem != objStep0002TsvPath.stem:
        raise ValueError("step0002のXLSXとTSVのファイル名が一致しません。")
    pszStep0003Stem: str = objStep0002ExcelPath.stem.replace(
        pszMarker, "ProductCodeSelector_step0003_", 1
    )
    return (
        objStep0002ExcelPath.with_name(pszStep0003Stem + ".xlsx"),
        objStep0002TsvPath.with_name(pszStep0003Stem + ".tsv"),
    )


def get_step0004_output_paths(
    objStep0003ExcelPath: Path, objStep0003TsvPath: Path
) -> tuple[Path, Path]:
    """step0003の出力名からstep0004のXLSX・TSVパスを作ります。"""
    pszMarker: str = "ProductCodeSelector_step0003_"
    if not objStep0003ExcelPath.stem.startswith(pszMarker):
        raise ValueError("step0003の出力ファイル名ではありません。")
    if objStep0003ExcelPath.suffix.lower() != ".xlsx":
        raise ValueError("step0003のXLSXファイルではありません。")
    if objStep0003TsvPath.suffix.lower() != ".tsv":
        raise ValueError("step0003のTSVファイルではありません。")
    if objStep0003ExcelPath.stem != objStep0003TsvPath.stem:
        raise ValueError("step0003のXLSXとTSVのファイル名が一致しません。")
    pszStep0004Stem: str = objStep0003ExcelPath.stem.replace(
        pszMarker, "ProductCodeSelector_step0004_", 1
    )
    return (
        objStep0003ExcelPath.with_name(pszStep0004Stem + ".xlsx"),
        objStep0003TsvPath.with_name(pszStep0004Stem + ".tsv"),
    )


def get_area_store_mapping_file_path() -> Path:
    """プログラムと同じフォルダーのエリア・店舗対応表を返します。"""
    return Path(__file__).resolve().parent / AREA_STORE_MAPPING_FILE_NAME


def read_area_store_mapping(objMappingPath: Path) -> dict[str, str]:
    """4列の対応表を検証し、店舗コードごとのエリアを返します。"""
    if not objMappingPath.is_file():
        raise ValueError(
            AREA_STORE_MAPPING_FILE_NAME
            + " が見つかりません。Path = "
            + str(objMappingPath)
        )
    with objMappingPath.open(mode="r", encoding="utf-8-sig", newline="") as objFile:
        listRows: list[list[str]] = list(
            csv.reader(objFile, delimiter="\t", strict=True)
        )
    tupleExpectedHeaders: tuple[str, str, str, str] = (
        "配送センター名",
        "エリア名",
        "店舗コード",
        "店舗略称",
    )
    if not listRows or tuple(listRows[0]) != tupleExpectedHeaders:
        raise ValueError(
            AREA_STORE_MAPPING_FILE_NAME + "のヘッダーが4列仕様と一致しません。"
        )
    dictStoreAreas: dict[str, str] = {}
    for iRow, listRow in enumerate(listRows[1:], start=2):
        if len(listRow) != len(tupleExpectedHeaders):
            raise ValueError(
                f"{AREA_STORE_MAPPING_FILE_NAME}の{iRow}行目が4列ではありません。"
            )
        pszCenterName, pszAreaName, pszStoreCode, pszStoreName = (
            pszValue.strip() for pszValue in listRow
        )
        if not all((pszCenterName, pszAreaName, pszStoreCode, pszStoreName)):
            raise ValueError(
                f"{AREA_STORE_MAPPING_FILE_NAME}の{iRow}行目に空欄があります。"
            )
        if pszAreaName not in AREA_NAMES:
            raise ValueError(
                f"{AREA_STORE_MAPPING_FILE_NAME}の{iRow}行目のエリア名が不正です。"
                + " Value = "
                + pszAreaName
            )
        try:
            pszNormalizedCode: str = str(int(pszStoreCode))
        except ValueError as objException:
            raise ValueError(
                f"{AREA_STORE_MAPPING_FILE_NAME}の{iRow}行目の店舗コードが不正です。"
            ) from objException
        if pszNormalizedCode in dictStoreAreas:
            raise ValueError(
                AREA_STORE_MAPPING_FILE_NAME
                + "の店舗コードが重複しています。店舗コード = "
                + pszNormalizedCode
            )
        dictStoreAreas[pszNormalizedCode] = pszAreaName
    return dictStoreAreas


def get_step0003_store_order_output_paths(
    objStep0002TsvPath: Path,
) -> tuple[Path, Path, Path, Path]:
    """step0002名から4つの店舗別step0003 TSVパスを返します。"""
    pszMarker: str = "ProductCodeSelector_step0002_"
    if not objStep0002TsvPath.stem.startswith(pszMarker):
        raise ValueError("step0002の出力ファイル名ではありません。")
    pszStem: str = objStep0002TsvPath.stem.replace(
        pszMarker, "ProductCodeSelector_step0003_", 1
    )
    return tuple(
        objStep0002TsvPath.with_name(pszStem + pszSuffix + ".tsv")
        for pszSuffix in ("_store_order", "_広島", "_岡山", "_四国")
    )


def build_store_order_rows(
    listStep0002Rows: list[list[str]], dictStoreAreas: dict[str, str]
) -> tuple[list[list[str]], list[list[str]], list[list[str]], list[list[str]]]:
    """O1から最終店舗列の9行を転置し、全店舗と3エリアに分けます。"""
    if len(listStep0002Rows) != 9:
        raise ValueError("step0002は9行ではありません。")
    iColumnCount: int = len(listStep0002Rows[0])
    if iColumnCount <= len(STEP_HEADERS):
        raise ValueError("step0002のO列以降に店舗がありません。")
    if any(len(listRow) != iColumnCount for listRow in listStep0002Rows):
        raise ValueError("step0002の行ごとの列数が一致しません。")
    listAllRows: list[list[str]] = []
    dictAreaRows: dict[str, list[list[str]]] = {
        pszAreaName: [] for pszAreaName in AREA_NAMES
    }
    listMissingStores: list[tuple[str, str]] = []
    setStoreCodes: set[str] = set()
    for iColumn in range(len(STEP_HEADERS), iColumnCount):
        pszRawCode: str = listStep0002Rows[0][iColumn].strip()
        pszStoreName: str = listStep0002Rows[1][iColumn].strip()
        if not pszRawCode or not pszStoreName:
            raise ValueError(
                f"step0002の{iColumn + 1}列目の店舗コードまたは店舗略称が空欄です。"
            )
        objCodeMatch: re.Match[str] | None = re.fullmatch(r"(\d+)(?:\.0+)?", pszRawCode)
        if objCodeMatch is None:
            raise ValueError(
                f"step0002の{iColumn + 1}列目の店舗コードが不正です。"
            )
        pszStoreCode: str = str(int(objCodeMatch.group(1)))
        if pszStoreCode in setStoreCodes:
            raise ValueError("step0002の店舗コードが重複しています。")
        setStoreCodes.add(pszStoreCode)
        listStoreRow: list[str] = [
            pszStoreCode,
            pszStoreName,
            *(listStep0002Rows[iRow][iColumn] for iRow in range(2, 9)),
        ]
        listAllRows.append(listStoreRow)
        pszAreaName: str | None = dictStoreAreas.get(pszStoreCode)
        if pszAreaName is None:
            listMissingStores.append((pszStoreCode, pszStoreName))
        else:
            dictAreaRows[pszAreaName].append(listStoreRow.copy())
    if listMissingStores:
        pszDetails: str = "\n".join(
            "店舗コード = " + pszCode + "、店舗略称 = " + pszName
            for pszCode, pszName in listMissingStores
        )
        raise ValueError(
            "step0002の次の店舗コードが"
            + AREA_STORE_MAPPING_FILE_NAME
            + "に存在しません。\n"
            + pszDetails
        )
    return (
        listAllRows,
        dictAreaRows["広島"],
        dictAreaRows["岡山"],
        dictAreaRows["四国／岡山"],
    )


def get_unique_path(objDesiredPath: Path) -> Path:
    """既存ファイルまたはフォルダーを上書きしないパスを返します。"""
    if not objDesiredPath.exists():
        return objDesiredPath
    iSequence: int = 2
    while True:
        objCandidate: Path = objDesiredPath.with_name(
            objDesiredPath.stem + "_" + str(iSequence) + objDesiredPath.suffix
        )
        if not objCandidate.exists():
            return objCandidate
        iSequence += 1


def archive_existing_store_order_files(
    tupleOutputPaths: tuple[Path, Path, Path, Path],
) -> tuple[Path | None, list[Path]]:
    """過去の通常名TSVを%TEMP%へコピー後、更新日時付きへ変更します。"""
    listExistingPaths: list[Path] = [
        objPath for objPath in tupleOutputPaths if objPath.is_file()
    ]
    if not listExistingPaths:
        return None, []
    iRequiredBytes: int = sum(objPath.stat().st_size for objPath in listExistingPaths)
    iFreeBytes: int = shutil.disk_usage(tempfile.gettempdir()).free
    if iRequiredBytes > iFreeBytes:
        raise ValueError(
            "%TEMP%の空き容量が不足しています。必要容量 = "
            + str(iRequiredBytes)
            + "、空き容量 = "
            + str(iFreeBytes)
        )
    objTempRoot: Path = (
        Path(tempfile.gettempdir()) / "AsahiSingleOrderProductCodeSelector"
    )
    objTempRoot.mkdir(parents=True, exist_ok=True)
    objBackupDirectory: Path = get_unique_path(
        objTempRoot / datetime.now().strftime("%Y%m%d%H%M%S")
    )
    objBackupDirectory.mkdir()
    listCopiedPaths: list[Path] = []
    try:
        for objPath in listExistingPaths:
            objCopyPath: Path = objBackupDirectory / objPath.name
            shutil.copy2(objPath, objCopyPath)
            listCopiedPaths.append(objCopyPath)
            if not objCopyPath.is_file() or objCopyPath.stat().st_size != objPath.stat().st_size:
                raise ValueError(
                    "%TEMP%へのバックアップ検証に失敗しました。Path = "
                    + str(objPath)
                )
    except Exception:
        for objCopiedPath in listCopiedPaths:
            if objCopiedPath.exists():
                objCopiedPath.unlink()
        try:
            objBackupDirectory.rmdir()
        except OSError:
            pass
        raise
    listRenames: list[tuple[Path, Path]] = []
    try:
        for objPath in listExistingPaths:
            pszModifiedTimestamp: str = datetime.fromtimestamp(
                objPath.stat().st_mtime
            ).strftime("%Y%m%d%H%M%S")
            objArchivePath: Path = get_unique_path(
                objPath.with_name(
                    objPath.stem + "_" + pszModifiedTimestamp + objPath.suffix
                )
            )
            objPath.rename(objArchivePath)
            listRenames.append((objPath, objArchivePath))
    except Exception:
        listRestoreFailures: list[str] = []
        for objOriginalPath, objArchivePath in reversed(listRenames):
            try:
                objArchivePath.rename(objOriginalPath)
            except OSError:
                listRestoreFailures.append(str(objArchivePath))
        if listRestoreFailures:
            raise ValueError(
                "過去の店舗別TSVの復元に失敗しました。"
                + " %TEMP%バックアップ = "
                + str(objBackupDirectory)
                + "、復元失敗パス = "
                + ", ".join(listRestoreFailures)
            )
        raise
    return objBackupDirectory, [objArchivePath for _, objArchivePath in listRenames]


def replace_new_output_set(dictTemporaryOutputs: dict[Path, Path]) -> None:
    """4つの新規TSVを一括確定し、失敗時は今回確定分を削除します。"""
    listReplacedPaths: list[Path] = []
    try:
        for objOutputPath, objTemporaryPath in dictTemporaryOutputs.items():
            os.replace(objTemporaryPath, objOutputPath)
            listReplacedPaths.append(objOutputPath)
    except Exception:
        for objOutputPath in reversed(listReplacedPaths):
            if objOutputPath.exists():
                objOutputPath.unlink()
        raise


def create_step0003_store_order_outputs(
    objStep0002ExcelPath: Path, objStep0002TsvPath: Path
) -> tuple[tuple[Path, Path, Path, Path], Path | None, list[Path]]:
    """step0002を転置・エリア分割し、4つの店舗別TSVを作成します。"""
    if not objStep0002ExcelPath.is_file() or not objStep0002TsvPath.is_file():
        raise ValueError("step0002のXLSXとTSVの両方が必要です。")
    listExcelRows, _ = read_excel_table(objStep0002ExcelPath)
    listTsvRows, _ = read_tsv_table(objStep0002TsvPath)
    if listExcelRows != listTsvRows:
        raise ValueError("step0002のXLSXとTSVの内容が一致しません。")
    dictStoreAreas: dict[str, str] = read_area_store_mapping(
        get_area_store_mapping_file_path()
    )
    tupleOutputRows = build_store_order_rows(listExcelRows, dictStoreAreas)
    tupleOutputPaths = get_step0003_store_order_output_paths(objStep0002TsvPath)
    objBackupDirectory, listArchivePaths = archive_existing_store_order_files(
        tupleOutputPaths
    )
    dictTemporaryOutputs: dict[Path, Path] = {}
    try:
        for objOutputPath, listRows in zip(tupleOutputPaths, tupleOutputRows):
            objTemporaryPath: Path = create_temporary_path(objOutputPath)
            dictTemporaryOutputs[objOutputPath] = objTemporaryPath
            save_tsv_table(objTemporaryPath, listRows)
            listSavedRows, _ = read_tsv_table(objTemporaryPath)
            if listSavedRows != listRows:
                raise ValueError(
                    "店舗別step0003 TSVの保存内容が仕様と一致しません。Path = "
                    + str(objOutputPath)
                )
            if any(len(listRow) != 9 for listRow in listSavedRows):
                raise ValueError(
                    "店舗別step0003 TSVに9列ではない行があります。Path = "
                    + str(objOutputPath)
                )
        if sorted(
            pszRow[0] for listRows in tupleOutputRows[1:] for pszRow in listRows
        ) != sorted(pszRow[0] for pszRow in tupleOutputRows[0]):
            raise ValueError("エリア別TSVの店舗集合が全店舗TSVと一致しません。")
        replace_new_output_set(dictTemporaryOutputs)
    except Exception as objException:
        pszArchiveDetail: str = ""
        if objBackupDirectory is not None:
            pszArchiveDetail = (
                "\n過去店舗別TSVのアーカイブ: 完了"
                + "\n%TEMP%バックアップ: "
                + str(objBackupDirectory)
                + "\n日時付きファイル: "
                + (", ".join(str(objPath) for objPath in listArchivePaths) or "なし")
                + "\n通常名の新規出力: 未作成"
            )
        raise ValueError(str(objException) + pszArchiveDetail) from objException
    finally:
        for objTemporaryPath in dictTemporaryOutputs.values():
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()
    return tupleOutputPaths, objBackupDirectory, listArchivePaths


def normalize_weekly_tsv_value(objValue: object) -> str:
    """A1:AB42の保存済みセル値をTSV用文字列へ変換します。"""
    if objValue is None:
        return ""
    if isinstance(objValue, bool):
        return "TRUE" if objValue else "FALSE"
    if isinstance(objValue, datetime):
        return objValue.isoformat(sep=" ")
    if isinstance(objValue, date):
        return objValue.isoformat()
    if isinstance(objValue, float) and objValue.is_integer():
        return str(int(objValue))
    return str(objValue)


def validate_weekly_template(objWorkbook: Workbook) -> Worksheet:
    """週間予定表の対象シートとA1:AB42が非結合であることを確認します。"""
    if WEEKLY_SHEET_NAME not in objWorkbook.sheetnames:
        raise ValueError("テンプレートに「センター週間」シートがありません。")
    objWorksheet: Worksheet = objWorkbook[WEEKLY_SHEET_NAME]
    for objMergedRange in objWorksheet.merged_cells.ranges:
        if not (
            objMergedRange.max_row < 1
            or objMergedRange.min_row > WEEKLY_TSV_MAX_ROW
            or objMergedRange.max_col < 1
            or objMergedRange.min_col > WEEKLY_TSV_MAX_COLUMN
        ):
            raise ValueError(
                "「センター週間」!"
                + WEEKLY_TSV_RANGE
                + "に結合セルがあります。範囲 = "
                + str(objMergedRange)
            )
    return objWorksheet


def build_weekly_tsv_rows(
    objCachedWorksheet: Worksheet,
    pszCreationDate: str,
    listDeliveryDates: list[date],
) -> list[list[str]]:
    """A1:AB42の保存済み計算結果を42行×28列で返します。"""
    listRows: list[list[str]] = [
        [
            normalize_weekly_tsv_value(objCachedWorksheet.cell(iRow, iColumn).value)
            for iColumn in range(1, WEEKLY_TSV_MAX_COLUMN + 1)
        ]
        for iRow in range(1, WEEKLY_TSV_MAX_ROW + 1)
    ]
    listRows[0][24] = pszCreationDate
    listShipmentDates: list[date] = [
        objDeliveryDate - timedelta(days=1) for objDeliveryDate in listDeliveryDates
    ]
    for iRow, iStartColumn in SHIPMENT_DATE_ROW_RANGES:
        for iOffset, objShipmentDate in enumerate(listShipmentDates):
            listRows[iRow - 1][iStartColumn - 1 + iOffset] = objShipmentDate.isoformat()
    for iRow, iStartColumn in DELIVERY_DATE_ROW_RANGES:
        for iOffset, objDeliveryDate in enumerate(listDeliveryDates):
            listRows[iRow - 1][iStartColumn - 1 + iOffset] = objDeliveryDate.isoformat()
    return listRows


def get_step0002_delivery_dates(listRows: list[list[str]]) -> list[date]:
    """step0002のA3:A9を月～日の連続した納品日として返します。"""
    if len(listRows) < 9 or any(len(listRow) < 2 for listRow in listRows[:9]):
        raise ValueError("step0002にA3:B9の1週間データがありません。")
    listDeliveryDates: list[date] = []
    for iOffset, pszExpectedWeekday in enumerate(WEEKDAYS):
        iRowIndex: int = iOffset + 2
        try:
            objDeliveryDate: date = datetime.strptime(
                listRows[iRowIndex][0], "%Y/%m/%d"
            ).date()
        except ValueError as objException:
            raise ValueError(
                f"step0002のA{iRowIndex + 1}の納品日が不正です。"
            ) from objException
        if listRows[iRowIndex][1].strip() != pszExpectedWeekday:
            raise ValueError(
                f"step0002のB{iRowIndex + 1}が{pszExpectedWeekday}曜日ではありません。"
            )
        if objDeliveryDate.weekday() != iOffset:
            raise ValueError(
                f"step0002のA{iRowIndex + 1}とB{iRowIndex + 1}の日付・曜日が一致しません。"
            )
        listDeliveryDates.append(objDeliveryDate)
    objMonday: date = listDeliveryDates[0]
    if any(
        objDeliveryDate != objMonday + timedelta(days=iOffset)
        for iOffset, objDeliveryDate in enumerate(listDeliveryDates)
    ):
        raise ValueError("step0002のA3:A9が月～日の連続日付ではありません。")
    return listDeliveryDates


def validate_fixed_weekdays(objWorksheet: Worksheet) -> None:
    """週間予定表の固定曜日が仕様どおりか確認します。"""
    tupleShipmentWeekdays: tuple[str, ...] = ("日", "月", "火", "水", "木", "金", "土")
    for iRow, iStartColumn in SHIPMENT_WEEKDAY_ROW_RANGES:
        tupleActual: tuple[str, ...] = tuple(
            normalize_weekly_tsv_value(
                objWorksheet.cell(iRow, iStartColumn + iOffset).value
            ).strip()
            for iOffset in range(7)
        )
        if tupleActual != tupleShipmentWeekdays:
            raise ValueError("週間予定表の出荷曜日が日～土の固定順ではありません。")
    for iRow, iStartColumn in DELIVERY_WEEKDAY_ROW_RANGES:
        tupleActual = tuple(
            normalize_weekly_tsv_value(
                objWorksheet.cell(iRow, iStartColumn + iOffset).value
            ).strip()
            for iOffset in range(7)
        )
        if tupleActual != WEEKDAYS:
            raise ValueError("週間予定表の納品曜日が月～日の固定順ではありません。")


def get_xml_local_name(pszQualifiedName: str) -> str:
    """XMLの名前空間付き名前からローカル名を返します。"""
    return pszQualifiedName.rsplit("}", 1)[-1]


def get_zip_member_bytes(objArchive: zipfile.ZipFile, pszMemberName: str) -> bytes:
    """XLSX内の重複していない必須パーツを読み込みます。"""
    if sum(objInfo.filename == pszMemberName for objInfo in objArchive.infolist()) != 1:
        raise ValueError(
            "XLSX内の必須パーツを1つに特定できません。Path = "
            + pszMemberName
        )
    return objArchive.read(pszMemberName)


def get_weekly_worksheet_part_name(objArchive: zipfile.ZipFile) -> str:
    """workbookとrelationshipからセンター週間のXMLパスを解決します。"""
    pszWorkbookPart: str = "xl/workbook.xml"
    bytesWorkbook: bytes = get_zip_member_bytes(objArchive, pszWorkbookPart)
    objWorkbookRoot: ET.Element = ET.fromstring(bytesWorkbook)
    listMatchedSheets: list[ET.Element] = [
        objElement
        for objElement in objWorkbookRoot.iter()
        if get_xml_local_name(objElement.tag) == "sheet"
        and objElement.attrib.get("name") == WEEKLY_SHEET_NAME
    ]
    if len(listMatchedSheets) != 1:
        raise ValueError(
            "XLSX内の「センター週間」シートを1つに特定できません。"
        )
    pszRelationshipId: str = next(
        (
            pszValue
            for pszName, pszValue in listMatchedSheets[0].attrib.items()
            if get_xml_local_name(pszName) == "id"
        ),
        "",
    )
    if not pszRelationshipId:
        raise ValueError("「センター週間」シートのrelationship IDがありません。")

    pszRelationshipsPart: str = "xl/_rels/workbook.xml.rels"
    bytesRelationships: bytes = get_zip_member_bytes(
        objArchive, pszRelationshipsPart
    )
    objRelationshipsRoot: ET.Element = ET.fromstring(bytesRelationships)
    listMatchedRelationships: list[ET.Element] = [
        objElement
        for objElement in objRelationshipsRoot.iter()
        if get_xml_local_name(objElement.tag) == "Relationship"
        and objElement.attrib.get("Id") == pszRelationshipId
    ]
    if len(listMatchedRelationships) != 1:
        raise ValueError(
            "「センター週間」シートのrelationshipを1つに特定できません。"
        )
    objRelationship: ET.Element = listMatchedRelationships[0]
    if objRelationship.attrib.get("TargetMode", "Internal") != "Internal":
        raise ValueError("「センター週間」シートがXLSX内部にありません。")
    pszTarget: str = objRelationship.attrib.get("Target", "").replace("\\", "/")
    if not pszTarget:
        raise ValueError("「センター週間」シートのXMLパスがありません。")
    if pszTarget.startswith("/"):
        pszWorksheetPart: str = posixpath.normpath(pszTarget.lstrip("/"))
    else:
        pszWorksheetPart = posixpath.normpath(
            posixpath.join(posixpath.dirname(pszWorkbookPart), pszTarget)
        )
    if pszWorksheetPart.startswith("../") or pszWorksheetPart not in objArchive.namelist():
        raise ValueError(
            "「センター週間」シートのXMLが見つかりません。Path = "
            + pszWorksheetPart
        )
    return pszWorksheetPart


def get_cell_xml_span(bytesWorksheet: bytes, pszCellReference: str) -> tuple[int, int, bytes]:
    """worksheetから指定セル要素のバイト範囲と接頭辞を返します。"""
    bytesReference: bytes = re.escape(pszCellReference.encode("ascii"))
    objCellPattern: re.Pattern[bytes] = re.compile(
        rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?c\b"
        rb"(?P<attributes>[^<>]*\br\s*=\s*(?P<quote>[\"'])"
        + bytesReference
        + rb"(?P=quote)[^<>]*)(?P<self_closing>/?)>"
    )
    listMatches: list[re.Match[bytes]] = list(objCellPattern.finditer(bytesWorksheet))
    if len(listMatches) != 1:
        raise ValueError(
            "「センター週間」シートの"
            + pszCellReference
            + "セルを1つに特定できません。"
        )
    objMatch: re.Match[bytes] = listMatches[0]
    bytesPrefix: bytes = objMatch.group("prefix") or b""
    if objMatch.group(0).rstrip().endswith(b"/>"):
        return objMatch.start(), objMatch.end(), bytesPrefix
    objClosingMatch: re.Match[bytes] | None = re.search(
        rb"</" + re.escape(bytesPrefix) + rb"c\s*>",
        bytesWorksheet[objMatch.end() :],
    )
    if objClosingMatch is None:
        raise ValueError(
            "「センター週間」シートの"
            + pszCellReference
            + "セルのXML終了要素がありません。"
        )
    iEnd: int = objMatch.end() + objClosingMatch.end()
    return objMatch.start(), iEnd, bytesPrefix


def update_creation_date_in_worksheet_xml(
    bytesWorksheet: bytes, pszCreationDate: str
) -> bytes:
    """worksheet XMLのY1だけを文字列の作成日へ変更します。"""
    iStart, iEnd, bytesPrefix = get_cell_xml_span(bytesWorksheet, "Y1")
    bytesOriginalCell: bytes = bytesWorksheet[iStart:iEnd]
    iStartTagEnd: int = bytesOriginalCell.find(b">")
    if iStartTagEnd < 0:
        raise ValueError("「センター週間」シートのY1セル形式が不正です。")
    bytesStartTag: bytes = bytesOriginalCell[: iStartTagEnd + 1]
    bytesStartTag = re.sub(
        rb"\s+t\s*=\s*([\"'])[^\"']*\1", b"", bytesStartTag, count=1
    )
    if bytesStartTag.endswith(b"/>"):
        bytesStartTag = bytesStartTag[:-2] + b' t="inlineStr">'
    else:
        bytesStartTag = bytesStartTag[:-1] + b' t="inlineStr">'
    pszEscapedDate: str = (
        pszCreationDate.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    bytesInlineString: bytes = (
        b"<"
        + bytesPrefix
        + b"is><"
        + bytesPrefix
        + b"t>"
        + pszEscapedDate.encode("utf-8")
        + b"</"
        + bytesPrefix
        + b"t></"
        + bytesPrefix
        + b"is></"
        + bytesPrefix
        + b"c>"
    )
    return bytesWorksheet[:iStart] + bytesStartTag + bytesInlineString + bytesWorksheet[iEnd:]


def update_excel_date_in_worksheet_xml(
    bytesWorksheet: bytes, pszCellReference: str, iExcelSerial: int
) -> bytes:
    """worksheet XMLの指定セルだけをExcel日付シリアル値へ変更します。"""
    iStart, iEnd, bytesPrefix = get_cell_xml_span(
        bytesWorksheet, pszCellReference
    )
    bytesOriginalCell: bytes = bytesWorksheet[iStart:iEnd]
    iStartTagEnd: int = bytesOriginalCell.find(b">")
    if iStartTagEnd < 0:
        raise ValueError(
            "「センター週間」シートの"
            + pszCellReference
            + "セル形式が不正です。"
        )
    bytesStartTag: bytes = bytesOriginalCell[: iStartTagEnd + 1]
    bytesStartTag = re.sub(
        rb"\s+t\s*=\s*([\"'])[^\"']*\1", b"", bytesStartTag, count=1
    )
    if bytesStartTag.endswith(b"/>"):
        bytesStartTag = bytesStartTag[:-2] + b">"
    bytesNumericValue: bytes = (
        b"<"
        + bytesPrefix
        + b"v>"
        + str(iExcelSerial).encode("ascii")
        + b"</"
        + bytesPrefix
        + b"v></"
        + bytesPrefix
        + b"c>"
    )
    return bytesWorksheet[:iStart] + bytesStartTag + bytesNumericValue + bytesWorksheet[iEnd:]


def get_excel_date_epoch(objArchive: zipfile.ZipFile) -> date:
    """テンプレートの1900または1904日付システムの基準日を返します。"""
    bytesWorkbook: bytes = get_zip_member_bytes(objArchive, "xl/workbook.xml")
    objWorkbookRoot: ET.Element = ET.fromstring(bytesWorkbook)
    objWorkbookProperties: ET.Element | None = next(
        (
            objElement
            for objElement in objWorkbookRoot.iter()
            if get_xml_local_name(objElement.tag) == "workbookPr"
        ),
        None,
    )
    pszDate1904: str = (
        objWorkbookProperties.attrib.get("date1904", "0")
        if objWorkbookProperties is not None
        else "0"
    )
    if pszDate1904 not in ("0", "1", "false", "true"):
        raise ValueError("テンプレートのExcel日付システム設定が不正です。")
    return date(1904, 1, 1) if pszDate1904 in ("1", "true") else date(1899, 12, 30)


def update_weekly_dates_in_worksheet_xml(
    bytesWorksheet: bytes,
    listDeliveryDates: list[date],
    objExcelEpoch: date,
) -> bytes:
    """出荷日・納品日の6範囲に同じ1週間をExcel日付で設定します。"""
    listShipmentDates: list[date] = [
        objDeliveryDate - timedelta(days=1) for objDeliveryDate in listDeliveryDates
    ]
    for tupleRanges, listDates in (
        (SHIPMENT_DATE_ROW_RANGES, listShipmentDates),
        (DELIVERY_DATE_ROW_RANGES, listDeliveryDates),
    ):
        for iRow, iStartColumn in tupleRanges:
            for iOffset, objTargetDate in enumerate(listDates):
                pszCellReference: str = (
                    get_column_letter(iStartColumn + iOffset) + str(iRow)
                )
                iExcelSerial: int = (objTargetDate - objExcelEpoch).days
                bytesWorksheet = update_excel_date_in_worksheet_xml(
                    bytesWorksheet, pszCellReference, iExcelSerial
                )
    return bytesWorksheet


def save_weekly_template_with_creation_date(
    objTemplatePath: Path,
    objOutputPath: Path,
    pszCreationDate: str,
    listDeliveryDates: list[date],
) -> str:
    """XLSXの描画パーツを保ったままY1と週間日付を更新します。"""
    with zipfile.ZipFile(objTemplatePath, mode="r") as objSourceArchive:
        pszWorksheetPart: str = get_weekly_worksheet_part_name(objSourceArchive)
        bytesWorksheet: bytes = get_zip_member_bytes(
            objSourceArchive, pszWorksheetPart
        )
        bytesUpdatedWorksheet: bytes = update_creation_date_in_worksheet_xml(
            bytesWorksheet, pszCreationDate
        )
        bytesUpdatedWorksheet = update_weekly_dates_in_worksheet_xml(
            bytesUpdatedWorksheet,
            listDeliveryDates,
            get_excel_date_epoch(objSourceArchive),
        )
        with zipfile.ZipFile(objOutputPath, mode="w") as objOutputArchive:
            objOutputArchive.comment = objSourceArchive.comment
            for objInfo in objSourceArchive.infolist():
                bytesContent: bytes = (
                    bytesUpdatedWorksheet
                    if objInfo.filename == pszWorksheetPart
                    else objSourceArchive.read(objInfo.filename)
                )
                objOutputArchive.writestr(objInfo, bytesContent)
    return pszWorksheetPart


def validate_unmodified_xlsx_parts(
    objTemplatePath: Path, objOutputPath: Path, pszWorksheetPart: str
) -> None:
    """Y1を含むworksheet以外のXLSX内部パーツが不変か確認します。"""
    with zipfile.ZipFile(objTemplatePath, mode="r") as objTemplateArchive:
        with zipfile.ZipFile(objOutputPath, mode="r") as objOutputArchive:
            listTemplateNames: list[str] = [
                objInfo.filename for objInfo in objTemplateArchive.infolist()
            ]
            listOutputNames: list[str] = [
                objInfo.filename for objInfo in objOutputArchive.infolist()
            ]
            if listTemplateNames != listOutputNames:
                raise ValueError("step0003 XLSXの内部パーツ構成が変更されました。")
            for pszMemberName in listTemplateNames:
                if pszMemberName == pszWorksheetPart:
                    continue
                if objTemplateArchive.read(pszMemberName) != objOutputArchive.read(
                    pszMemberName
                ):
                    raise ValueError(
                        "step0003 XLSXの変更対象外パーツが変更されました。Path = "
                        + pszMemberName
                    )


def validate_step0003_outputs(
    objExcelPath: Path,
    objTsvPath: Path,
    listExpectedTsvRows: list[list[str]],
    pszCreationDate: str,
    listDeliveryDates: list[date],
) -> None:
    """step0003の作成日とA1:AB42 TSVを保存後に確認します。"""
    objWorkbook: Workbook = load_workbook(objExcelPath, data_only=False)
    try:
        objWorksheet: Worksheet = validate_weekly_template(objWorkbook)
        objCreationDate: object = objWorksheet["Y1"].value
        if not isinstance(objCreationDate, str):
            raise ValueError("「センター週間」!Y1が文字列ではありません。")
        if objCreationDate != pszCreationDate:
            raise ValueError("「センター週間」!Y1の作成日が一致しません。")
        if objWorksheet["Y1"].data_type == "f" or objCreationDate.startswith(("=", "'")):
            raise ValueError("「センター週間」!Y1が正しい文字列セルではありません。")
        validate_fixed_weekdays(objWorksheet)
        listShipmentDates: list[date] = [
            objDeliveryDate - timedelta(days=1)
            for objDeliveryDate in listDeliveryDates
        ]
        for tupleRanges, listExpectedDates in (
            (SHIPMENT_DATE_ROW_RANGES, listShipmentDates),
            (DELIVERY_DATE_ROW_RANGES, listDeliveryDates),
        ):
            for iRow, iStartColumn in tupleRanges:
                for iOffset, objExpectedDate in enumerate(listExpectedDates):
                    objValue: object = objWorksheet.cell(
                        iRow, iStartColumn + iOffset
                    ).value
                    if isinstance(objValue, datetime):
                        objActualDate: date | None = objValue.date()
                    elif isinstance(objValue, date):
                        objActualDate = objValue
                    else:
                        objActualDate = None
                    if objActualDate != objExpectedDate:
                        raise ValueError(
                            "step0003 XLSXの出荷日または納品日が"
                            "Excel日付型の期待値と一致しません。セル = "
                            + get_column_letter(iStartColumn + iOffset)
                            + str(iRow)
                        )
    finally:
        objWorkbook.close()
    listTsvRows, _ = read_tsv_table(objTsvPath)
    if len(listTsvRows) != WEEKLY_TSV_MAX_ROW or any(
        len(listRow) != WEEKLY_TSV_MAX_COLUMN for listRow in listTsvRows
    ):
        raise ValueError(
            "step0003 TSVが"
            + str(WEEKLY_TSV_MAX_ROW)
            + "行×"
            + str(WEEKLY_TSV_MAX_COLUMN)
            + "列ではありません。"
        )
    if listTsvRows != listExpectedTsvRows:
        raise ValueError(
            "step0003 TSVが「センター週間」!"
            + WEEKLY_TSV_RANGE
            + "と一致しません。"
        )


def create_step0003_outputs(
    objStep0002ExcelPath: Path, objStep0002TsvPath: Path
) -> tuple[Path, Path]:
    """step0002ペアを再読込し、作成日入りの週間予定表とTSVを作ります。"""
    if not objStep0002ExcelPath.is_file() or not objStep0002TsvPath.is_file():
        raise ValueError("step0002のXLSXとTSVの両方が必要です。")
    listExcelRows, _ = read_excel_table(objStep0002ExcelPath)
    listTsvRows, _ = read_tsv_table(objStep0002TsvPath)
    if listExcelRows != listTsvRows:
        raise ValueError("step0002のXLSXとTSVの内容が一致しません。")
    listDeliveryDates: list[date] = get_step0002_delivery_dates(listExcelRows)

    objTemplatePath: Path = get_weekly_template_file_path()
    if not objTemplatePath.is_file():
        raise ValueError(
            "週間予定表テンプレートが見つかりません。Path = "
            + str(objTemplatePath)
        )
    objStep0003ExcelPath, objStep0003TsvPath = get_step0003_output_paths(
        objStep0002ExcelPath, objStep0002TsvPath
    )
    pszCreationDate: str = date.today().strftime("%Y年%m月%d日")
    objCachedWorkbook: Workbook = load_workbook(objTemplatePath, data_only=True)
    try:
        objCachedWorksheet: Worksheet = validate_weekly_template(objCachedWorkbook)
        validate_fixed_weekdays(objCachedWorksheet)
        listWeeklyRows: list[list[str]] = build_weekly_tsv_rows(
            objCachedWorksheet, pszCreationDate, listDeliveryDates
        )
        objTemporaryExcelPath: Path = create_temporary_path(objStep0003ExcelPath)
        objTemporaryTsvPath: Path = create_temporary_path(objStep0003TsvPath)
        try:
            pszWorksheetPart: str = save_weekly_template_with_creation_date(
                objTemplatePath,
                objTemporaryExcelPath,
                pszCreationDate,
                listDeliveryDates,
            )
            save_tsv_table(objTemporaryTsvPath, listWeeklyRows)
            validate_step0003_outputs(
                objTemporaryExcelPath,
                objTemporaryTsvPath,
                listWeeklyRows,
                pszCreationDate,
                listDeliveryDates,
            )
            validate_unmodified_xlsx_parts(
                objTemplatePath, objTemporaryExcelPath, pszWorksheetPart
            )
            replace_output_pair(
                objTemporaryExcelPath,
                objTemporaryTsvPath,
                objStep0003ExcelPath,
                objStep0003TsvPath,
            )
        finally:
            for objTemporaryPath in (objTemporaryExcelPath, objTemporaryTsvPath):
                if objTemporaryPath.exists():
                    objTemporaryPath.unlink()
    finally:
        objCachedWorkbook.close()
    return objStep0003ExcelPath, objStep0003TsvPath


def is_weekly_date_cell(iRow: int, iColumn: int) -> bool:
    """セルが週間予定表の出荷日・納品日範囲内ならTrueを返します。"""
    return any(
        iRow == iDateRow
        and iStartColumn <= iColumn < iStartColumn + 7
        for tupleRanges in (SHIPMENT_DATE_ROW_RANGES, DELIVERY_DATE_ROW_RANGES)
        for iDateRow, iStartColumn in tupleRanges
    )


def normalize_weekly_xlsx_value(
    objValue: object, iRow: int, iColumn: int
) -> str:
    """step0003・step0004 XLSXの値をTSVと比較できる文字列にします。"""
    if is_weekly_date_cell(iRow, iColumn):
        if isinstance(objValue, datetime):
            return objValue.date().isoformat()
        if isinstance(objValue, date):
            return objValue.isoformat()
    return normalize_weekly_tsv_value(objValue)


def read_weekly_xlsx_rows(objExcelPath: Path) -> list[list[str]]:
    """週間予定表XLSXのA1:AB42を論理比較用文字列で返します。"""
    objWorkbook: Workbook = load_workbook(objExcelPath, data_only=True)
    try:
        objWorksheet: Worksheet = validate_weekly_template(objWorkbook)
        return [
            [
                normalize_weekly_xlsx_value(
                    objWorksheet.cell(iRow, iColumn).value, iRow, iColumn
                )
                for iColumn in range(1, WEEKLY_TSV_MAX_COLUMN + 1)
            ]
            for iRow in range(1, WEEKLY_TSV_MAX_ROW + 1)
        ]
    finally:
        objWorkbook.close()


def validate_weekly_tsv_size(listRows: list[list[str]], pszStepName: str) -> None:
    """週間予定表TSVが42行×28列であることを確認します。"""
    if len(listRows) != WEEKLY_TSV_MAX_ROW or any(
        len(listRow) != WEEKLY_TSV_MAX_COLUMN for listRow in listRows
    ):
        raise ValueError(
            pszStepName
            + " TSVが"
            + str(WEEKLY_TSV_MAX_ROW)
            + "行×"
            + str(WEEKLY_TSV_MAX_COLUMN)
            + "列ではありません。"
        )


def validate_weekly_xlsx_tsv_match(
    objExcelPath: Path, objTsvPath: Path, pszStepName: str
) -> list[list[str]]:
    """週間予定表XLSXとTSVのA1:AB42が論理的に一致することを確認します。"""
    listTsvRows, _ = read_tsv_table(objTsvPath)
    validate_weekly_tsv_size(listTsvRows, pszStepName)
    listExcelRows: list[list[str]] = read_weekly_xlsx_rows(objExcelPath)
    for iRowIndex, (listExcelRow, listTsvRow) in enumerate(
        zip(listExcelRows, listTsvRows), start=1
    ):
        for iColumnIndex, (pszExcelValue, pszTsvValue) in enumerate(
            zip(listExcelRow, listTsvRow), start=1
        ):
            if pszExcelValue != pszTsvValue:
                raise ValueError(
                    pszStepName
                    + " XLSXとTSVの内容が一致しません。セル = "
                    + get_column_letter(iColumnIndex)
                    + str(iRowIndex)
                    + "、XLSX = "
                    + repr(pszExcelValue)
                    + "、TSV = "
                    + repr(pszTsvValue)
                )
    return listTsvRows


def normalize_step0004_area_rows(
    objAreaTsvPath: Path, pszAreaName: str
) -> list[list[str]]:
    """エリア別TSVを検証し、XLSXへ設定する9列の値へ正規化します。"""
    if not objAreaTsvPath.is_file():
        raise ValueError(
            "step0003 " + pszAreaName + " TSVが見つかりません。Path = "
            + str(objAreaTsvPath)
        )
    listRows, _ = read_tsv_table(objAreaTsvPath)
    if len(listRows) > STEP0004_MAX_STORES_PER_AREA:
        raise ValueError(
            "step0003 "
            + pszAreaName
            + " TSVの店舗数が31店舗を超えています。店舗数 = "
            + str(len(listRows))
        )
    listNormalizedRows: list[list[str]] = []
    for iRow, listRow in enumerate(listRows, start=1):
        if len(listRow) != 9:
            raise ValueError(
                "step0003 "
                + pszAreaName
                + " TSVに9列ではない行があります。行 = "
                + str(iRow)
            )
        pszStoreCode: str = listRow[0].strip()
        objStoreCodeMatch: re.Match[str] | None = re.fullmatch(
            r"(\d+)(?:\.0+)?", pszStoreCode
        )
        if objStoreCodeMatch is None:
            raise ValueError(
                "step0003 "
                + pszAreaName
                + " TSVの店舗コードが整数ではありません。行 = "
                + str(iRow)
            )
        pszStoreName: str = listRow[1].strip()
        if not pszStoreName:
            raise ValueError(
                "step0003 "
                + pszAreaName
                + " TSVの店舗略称が空欄です。行 = "
                + str(iRow)
            )
        listNormalizedQuantities: list[str] = []
        for iColumn, pszRawQuantity in enumerate(listRow[2:], start=3):
            pszQuantity: str = pszRawQuantity.strip()
            if not pszQuantity:
                listNormalizedQuantities.append("")
                continue
            if re.fullmatch(r"[+-]?\d+", pszQuantity) is None:
                raise ValueError(
                    "step0003 "
                    + pszAreaName
                    + " TSVの発注数量が整数ではありません。行 = "
                    + str(iRow)
                    + "、列 = "
                    + str(iColumn)
                )
            iQuantity: int = int(pszQuantity)
            listNormalizedQuantities.append(
                "" if iQuantity == 0 else str(iQuantity)
            )
        listNormalizedRows.append(
            [
                str(int(objStoreCodeMatch.group(1))),
                pszStoreName,
                *listNormalizedQuantities,
            ]
        )
    return listNormalizedRows


def build_step0004_rows(
    listStep0003Rows: list[list[str]],
    tupleAreaRows: tuple[list[list[str]], list[list[str]], list[list[str]]],
) -> list[list[str]]:
    """3エリアの31行×9列を空欄化し、エリア別TSVを上から転記します。"""
    listOutputRows: list[list[str]] = [listRow.copy() for listRow in listStep0003Rows]
    for (
        (_, iStartRow, iEndRow, iStartColumn, iEndColumn),
        listAreaRows,
    ) in zip(STEP0004_AREA_RANGES, tupleAreaRows):
        if len(listAreaRows) > STEP0004_MAX_STORES_PER_AREA:
            raise ValueError("step0004のエリア別店舗数が31店舗を超えています。")
        for iRow in range(iStartRow, iEndRow + 1):
            for iColumn in range(iStartColumn, iEndColumn + 1):
                listOutputRows[iRow - 1][iColumn - 1] = ""
        for iRowOffset, listAreaRow in enumerate(listAreaRows):
            listOutputRows[iStartRow - 1 + iRowOffset][
                iStartColumn - 1 : iEndColumn
            ] = listAreaRow
    return listOutputRows


def set_cell_value_in_worksheet_xml(
    bytesWorksheet: bytes, pszCellReference: str, pszValue: str, bNumeric: bool
) -> bytes:
    """既存セルと書式属性を残し、値だけを数値・文字列・空欄へ更新します。"""
    iStart, iEnd, bytesPrefix = get_cell_xml_span(bytesWorksheet, pszCellReference)
    bytesOriginalCell: bytes = bytesWorksheet[iStart:iEnd]
    iStartTagEnd: int = bytesOriginalCell.find(b">")
    if iStartTagEnd < 0:
        raise ValueError(
            "「センター週間」シートの"
            + pszCellReference
            + "セル形式が不正です。"
        )
    bytesStartTag: bytes = bytesOriginalCell[: iStartTagEnd + 1]
    bytesStartTag = re.sub(
        rb"\s+t\s*=\s*([\"'])[^\"']*\1", b"", bytesStartTag, count=1
    )
    if bytesStartTag.endswith(b"/>"):
        bytesStartTag = bytesStartTag[:-2] + b">"
    if not pszValue:
        bytesNewContent: bytes = b""
    elif bNumeric:
        bytesNewContent = (
            b"<" + bytesPrefix + b"v>" + pszValue.encode("ascii")
            + b"</" + bytesPrefix + b"v>"
        )
    else:
        bytesStartTag = bytesStartTag[:-1] + b' t="inlineStr">'
        pszEscapedValue: str = (
            pszValue.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        bytesNewContent = (
            b"<" + bytesPrefix + b"is><" + bytesPrefix + b"t>"
            + pszEscapedValue.encode("utf-8")
            + b"</" + bytesPrefix + b"t></" + bytesPrefix + b"is>"
        )
    bytesNewCell: bytes = (
        bytesStartTag + bytesNewContent + b"</" + bytesPrefix + b"c>"
    )
    return bytesWorksheet[:iStart] + bytesNewCell + bytesWorksheet[iEnd:]


def update_step0004_cells_in_worksheet_xml(
    bytesWorksheet: bytes,
    tupleAreaRows: tuple[list[list[str]], list[list[str]], list[list[str]]],
) -> bytes:
    """3エリアの店舗コード・略称・月～日数量をセル値だけ更新します。"""
    for (
        (_, iStartRow, iEndRow, iStartColumn, iEndColumn),
        listAreaRows,
    ) in zip(STEP0004_AREA_RANGES, tupleAreaRows):
        for iRow in range(iStartRow, iEndRow + 1):
            iAreaRow: int = iRow - iStartRow
            for iColumn in range(iStartColumn, iEndColumn + 1):
                iAreaColumn: int = iColumn - iStartColumn
                pszValue: str = (
                    listAreaRows[iAreaRow][iAreaColumn]
                    if iAreaRow < len(listAreaRows)
                    else ""
                )
                bytesWorksheet = set_cell_value_in_worksheet_xml(
                    bytesWorksheet,
                    get_column_letter(iColumn) + str(iRow),
                    pszValue,
                    bNumeric=iAreaColumn != 1 and bool(pszValue),
                )
    return bytesWorksheet


def save_step0004_xlsx(
    objStep0003Path: Path,
    objStep0004Path: Path,
    tupleAreaRows: tuple[list[list[str]], list[list[str]], list[list[str]]],
) -> str:
    """step0003の描画・書式を保ち、3エリアのセル値だけ更新します。"""
    with zipfile.ZipFile(objStep0003Path, mode="r") as objSourceArchive:
        pszWorksheetPart: str = get_weekly_worksheet_part_name(objSourceArchive)
        bytesWorksheet: bytes = get_zip_member_bytes(
            objSourceArchive, pszWorksheetPart
        )
        bytesUpdatedWorksheet: bytes = update_step0004_cells_in_worksheet_xml(
            bytesWorksheet, tupleAreaRows
        )
        with zipfile.ZipFile(objStep0004Path, mode="w") as objOutputArchive:
            objOutputArchive.comment = objSourceArchive.comment
            for objInfo in objSourceArchive.infolist():
                bytesContent: bytes = (
                    bytesUpdatedWorksheet
                    if objInfo.filename == pszWorksheetPart
                    else objSourceArchive.read(objInfo.filename)
                )
                objOutputArchive.writestr(objInfo, bytesContent)
    return pszWorksheetPart


def validate_step0004_xlsx_parts(
    objStep0003Path: Path,
    objStep0004Path: Path,
    pszWorksheetPart: str,
    tupleAreaRows: tuple[list[list[str]], list[list[str]], list[list[str]]],
) -> None:
    """対象セル値以外のXLSX内部データが変わっていないことを確認します。"""
    with zipfile.ZipFile(objStep0003Path, mode="r") as objSourceArchive:
        with zipfile.ZipFile(objStep0004Path, mode="r") as objOutputArchive:
            listSourceNames: list[str] = [
                objInfo.filename for objInfo in objSourceArchive.infolist()
            ]
            listOutputNames: list[str] = [
                objInfo.filename for objInfo in objOutputArchive.infolist()
            ]
            if listSourceNames != listOutputNames:
                raise ValueError("step0004 XLSXの内部パーツ構成が変更されました。")
            for pszMemberName in listSourceNames:
                bytesSource: bytes = objSourceArchive.read(pszMemberName)
                bytesExpected: bytes = (
                    update_step0004_cells_in_worksheet_xml(bytesSource, tupleAreaRows)
                    if pszMemberName == pszWorksheetPart
                    else bytesSource
                )
                if objOutputArchive.read(pszMemberName) != bytesExpected:
                    raise ValueError(
                        "step0004 XLSXの指定セル値以外が変更されました。Path = "
                        + pszMemberName
                    )


def validate_step0004_outputs(
    objExcelPath: Path,
    objTsvPath: Path,
    listExpectedRows: list[list[str]],
) -> None:
    """step0004 XLSX・TSVと3エリアの転記内容を保存後に確認します。"""
    listTsvRows: list[list[str]] = validate_weekly_xlsx_tsv_match(
        objExcelPath, objTsvPath, "step0004"
    )
    if listTsvRows != listExpectedRows:
        raise ValueError("step0004 TSVが期待するA1:AB42と一致しません。")


def create_step0004_outputs(
    objStep0003ExcelPath: Path,
    objStep0003TsvPath: Path,
    tupleAreaTsvPaths: tuple[Path, Path, Path],
) -> tuple[Path, Path]:
    """step0003ペアと3つのエリア別TSVからstep0004を作ります。"""
    if not objStep0003ExcelPath.is_file() or not objStep0003TsvPath.is_file():
        raise ValueError("step0003のXLSXとTSVの両方が必要です。")
    objStep0004ExcelPath, objStep0004TsvPath = get_step0004_output_paths(
        objStep0003ExcelPath, objStep0003TsvPath
    )
    listStep0003Rows: list[list[str]] = validate_weekly_xlsx_tsv_match(
        objStep0003ExcelPath, objStep0003TsvPath, "step0003"
    )
    tupleAreaRows: tuple[list[list[str]], list[list[str]], list[list[str]]] = tuple(
        normalize_step0004_area_rows(objAreaPath, pszAreaName)
        for objAreaPath, pszAreaName in zip(
            tupleAreaTsvPaths, ("広島", "岡山", "四国")
        )
    )
    listStep0004Rows: list[list[str]] = build_step0004_rows(
        listStep0003Rows, tupleAreaRows
    )
    objTemporaryExcelPath: Path = create_temporary_path(objStep0004ExcelPath)
    objTemporaryTsvPath: Path = create_temporary_path(objStep0004TsvPath)
    try:
        pszWorksheetPart: str = save_step0004_xlsx(
            objStep0003ExcelPath, objTemporaryExcelPath, tupleAreaRows
        )
        save_tsv_table(objTemporaryTsvPath, listStep0004Rows)
        validate_step0004_outputs(
            objTemporaryExcelPath, objTemporaryTsvPath, listStep0004Rows
        )
        validate_step0004_xlsx_parts(
            objStep0003ExcelPath,
            objTemporaryExcelPath,
            pszWorksheetPart,
            tupleAreaRows,
        )
        replace_output_pair(
            objTemporaryExcelPath,
            objTemporaryTsvPath,
            objStep0004ExcelPath,
            objStep0004TsvPath,
        )
    finally:
        for objTemporaryPath in (objTemporaryExcelPath, objTemporaryTsvPath):
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()
    return objStep0004ExcelPath, objStep0004TsvPath


def build_step0002_rows(
    listStep0001Rows: list[list[str]], objCandidate: ProductCandidate
) -> list[list[str]]:
    """選択商品を月曜日へ設定し、元のG3～G8をG4～G9へ移動します。"""
    listRows: list[list[str]] = [listRow.copy() for listRow in listStep0001Rows]
    listOriginalProductNames: list[str] = [listRows[iRow][6] for iRow in range(2, 8)]
    for listRow in listRows[2:]:
        listRow[4] = ""
        listRow[5] = ""
    listRows[2][5] = objCandidate.code
    listRows[2][6] = objCandidate.name
    for iOffset, pszOriginalName in enumerate(listOriginalProductNames, start=3):
        listRows[iOffset][6] = pszOriginalName
    return listRows


def validate_step0002_outputs(
    objExcelPath: Path,
    objTsvPath: Path,
    listStep0001Rows: list[list[str]],
    objCandidate: ProductCandidate,
) -> None:
    """step0002のXLSX・TSV一致と、指定セル以外が不変であることを確認します。"""
    listExcelRows, _ = read_excel_table(objExcelPath)
    listTsvRows, _ = read_tsv_table(objTsvPath)
    if listExcelRows != listTsvRows:
        raise ValueError("step0002のXLSXとTSVの内容が一致しません。")
    listExpectedRows: list[list[str]] = build_step0002_rows(listStep0001Rows, objCandidate)
    if listExcelRows != listExpectedRows:
        raise ValueError("step0002の保存内容が仕様どおりではありません。")


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


def get_success_path(objInputPath: Path) -> Path:
    """入力ファイルを基準に成功テキストのパスを返します。"""
    return objInputPath.with_name(objInputPath.stem + "_success.txt")


def get_next_result_history_path(objResultPath: Path, pszResult: str) -> Path:
    """既存の最大連番の次の成功またはエラー履歴パスを返します。"""
    pszBaseStem: str = objResultPath.stem[: -(len(pszResult) + 1)]
    objPattern: re.Pattern[str] = re.compile(
        re.escape(pszBaseStem) + "_" + re.escape(pszResult) + r"_(\d{4,})\.txt$"
    )
    iMaximumSequence: int = 0
    for objPath in objResultPath.parent.glob(
        pszBaseStem + "_" + pszResult + "_*.txt"
    ):
        objMatch: re.Match[str] | None = objPattern.fullmatch(objPath.name)
        if objMatch is not None:
            iMaximumSequence = max(iMaximumSequence, int(objMatch.group(1)))
    return objResultPath.with_name(
        pszBaseStem + "_" + pszResult + "_" + f"{iMaximumSequence + 1:04d}.txt"
    )


def replace_result_text(
    objInputPath: Path, pszResult: str, pszText: str
) -> Path:
    """過去の成功・エラーを履歴化し、今回の結果を安全に保存します。"""
    if pszResult not in ("success", "error"):
        raise ValueError("結果テキストの種類が不正です。")
    objSuccessPath: Path = get_success_path(objInputPath)
    objErrorPath: Path = get_error_path(objInputPath)
    objCurrentPath: Path = objSuccessPath if pszResult == "success" else objErrorPath
    objTemporaryPath: Path = create_temporary_path(objCurrentPath)
    pszCrLfText: str = (
        pszText.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    )
    objTemporaryPath.write_bytes(pszCrLfText.encode("utf-8"))
    if objTemporaryPath.read_bytes() != pszCrLfText.encode("utf-8"):
        objTemporaryPath.unlink(missing_ok=True)
        raise ValueError("結果テキストの一時保存内容が一致しません。")

    listRenames: list[tuple[Path, Path]] = []
    try:
        for objOldPath, pszOldResult in (
            (objSuccessPath, "success"),
            (objErrorPath, "error"),
        ):
            if not objOldPath.is_file():
                continue
            objHistoryPath: Path = get_next_result_history_path(
                objOldPath, pszOldResult
            )
            objOldPath.rename(objHistoryPath)
            listRenames.append((objOldPath, objHistoryPath))
        os.replace(objTemporaryPath, objCurrentPath)
    except Exception:
        if objCurrentPath.exists() and not any(
            objOriginalPath == objCurrentPath
            for objOriginalPath, _ in listRenames
        ):
            objCurrentPath.unlink()
        for objOriginalPath, objHistoryPath in reversed(listRenames):
            if objHistoryPath.exists() and not objOriginalPath.exists():
                objHistoryPath.rename(objOriginalPath)
        raise
    finally:
        if objTemporaryPath.exists():
            objTemporaryPath.unlink()
    return objCurrentPath


def write_error_text(objErrorPath: Path, pszErrorMessage: str) -> None:
    """処理エラーをUTF-8テキストで保存します。"""
    pszText: str = (
        "処理名:\nProductCodeSelector step0001～step0004\n\n"
        + "エラー:\n"
        + pszErrorMessage
        + "\n"
    )
    pszInputStem: str = objErrorPath.stem.removesuffix("_error")
    objInputPath: Path = objErrorPath.with_name(pszInputStem)
    replace_result_text(objInputPath, "error", pszText)


def write_success_text(
    objInputPath: Path,
    objStep0001ExcelPath: Path,
    objStep0001TsvPath: Path,
    objStep0002ExcelPath: Path,
    objStep0002TsvPath: Path,
    tupleStoreOrderPaths: tuple[Path, Path, Path, Path],
    objStoreOrderBackupDirectory: Path | None,
    listStoreOrderArchivePaths: list[Path],
    objStep0003ExcelPath: Path,
    objStep0003TsvPath: Path,
    objStep0004ExcelPath: Path,
    objStep0004TsvPath: Path,
    pszProductName: str,
    objSelectedCandidate: ProductCandidate,
) -> Path:
    """全出力とバックアップ情報を今回の_success.txtへ保存します。"""
    listLines: list[str] = [
        "処理名:",
        "ProductCodeSelector step0001～step0004",
        "",
        "処理結果:",
        "成功",
        "",
        "入力ファイル:",
        str(objInputPath),
        "",
        "商品名:",
        pszProductName,
        "",
        "選択商品:",
        objSelectedCandidate.display_text,
        "",
        "出力ファイル:",
        "step0001 XLSX: " + str(objStep0001ExcelPath),
        "step0001 TSV: " + str(objStep0001TsvPath),
        "step0002 XLSX: " + str(objStep0002ExcelPath),
        "step0002 TSV: " + str(objStep0002TsvPath),
        "step0003 Store Order TSV: " + str(tupleStoreOrderPaths[0]),
        "step0003 広島 TSV: " + str(tupleStoreOrderPaths[1]),
        "step0003 岡山 TSV: " + str(tupleStoreOrderPaths[2]),
        "step0003 四国 TSV: " + str(tupleStoreOrderPaths[3]),
        "step0003 XLSX: " + str(objStep0003ExcelPath),
        "step0003 TSV: " + str(objStep0003TsvPath),
        "step0004 XLSX: " + str(objStep0004ExcelPath),
        "step0004 TSV: " + str(objStep0004TsvPath),
        "",
        "バックアップ:",
    ]
    if objStoreOrderBackupDirectory is None:
        listLines.append("なし")
    else:
        listLines.extend(
            [
                "%TEMP%バックアップ: "
                + str(objStoreOrderBackupDirectory),
                "",
                "日時付きアーカイブ:",
                *(str(objPath) for objPath in listStoreOrderArchivePaths),
            ]
        )
    return replace_result_text(objInputPath, "success", "\n".join(listLines) + "\n")


def process_input_file(
    pszInputFileFullPath: str,
) -> tuple[
    Path,
    Path,
    Path,
    Path,
    tuple[Path, Path, Path, Path],
    Path | None,
    list[Path],
    Path,
    Path,
    Path,
    Path,
    str,
    ProductCandidate,
]:
    """step0007からstep0001～step0004を順に作成します。"""
    objInputPath: Path = validate_input_path(pszInputFileFullPath)
    create_abc_product_master(
        get_source_products_file_path(), get_products_file_path()
    )
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
    listStep0001ExcelRows, pszStep0001WorksheetTitle = read_excel_table(objExcelOutputPath)
    listStep0001TsvRows, _ = read_tsv_table(objTsvOutputPath)
    if listStep0001ExcelRows != listStep0001TsvRows:
        raise ValueError("正式なstep0001のXLSXとTSVの内容が一致しません。")
    listAllCandidates: list[ProductCandidate] = read_product_candidates(
        get_products_file_path()
    )
    pszInputSpec: str = listStep0001ExcelRows[2][8].strip()
    listMatchedCandidates: list[ProductCandidate] = find_product_candidates(
        listAllCandidates, pszProductName, pszInputSpec
    )
    objSelectedCandidate: ProductCandidate = select_product_candidate(
        pszProductName,
        pszInputSpec,
        listMatchedCandidates,
        listAllCandidates,
        listStep0001ExcelRows[8][6].strip(),
    )
    listStep0002Rows: list[list[str]] = build_step0002_rows(
        listStep0001ExcelRows, objSelectedCandidate
    )
    objStep0002ExcelPath, objStep0002TsvPath = get_step0002_output_paths(
        objExcelOutputPath, objTsvOutputPath
    )
    objTemporaryStep0002ExcelPath: Path = create_temporary_path(objStep0002ExcelPath)
    objTemporaryStep0002TsvPath: Path = create_temporary_path(objStep0002TsvPath)
    try:
        save_excel_table(
            objTemporaryStep0002ExcelPath,
            listStep0002Rows,
            pszStep0001WorksheetTitle,
        )
        save_tsv_table(objTemporaryStep0002TsvPath, listStep0002Rows)
        validate_step0002_outputs(
            objTemporaryStep0002ExcelPath,
            objTemporaryStep0002TsvPath,
            listStep0001ExcelRows,
            objSelectedCandidate,
        )
        replace_output_pair(
            objTemporaryStep0002ExcelPath,
            objTemporaryStep0002TsvPath,
            objStep0002ExcelPath,
            objStep0002TsvPath,
        )
    finally:
        for objTemporaryPath in (
            objTemporaryStep0002ExcelPath,
            objTemporaryStep0002TsvPath,
        ):
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()
    (
        tupleStoreOrderPaths,
        objStoreOrderBackupDirectory,
        listStoreOrderArchivePaths,
    ) = create_step0003_store_order_outputs(
        objStep0002ExcelPath, objStep0002TsvPath
    )
    objStep0003ExcelPath, objStep0003TsvPath = create_step0003_outputs(
        objStep0002ExcelPath, objStep0002TsvPath
    )
    objStep0004ExcelPath, objStep0004TsvPath = create_step0004_outputs(
        objStep0003ExcelPath,
        objStep0003TsvPath,
        (tupleStoreOrderPaths[1], tupleStoreOrderPaths[2], tupleStoreOrderPaths[3]),
    )
    return (
        objExcelOutputPath,
        objTsvOutputPath,
        objStep0002ExcelPath,
        objStep0002TsvPath,
        tupleStoreOrderPaths,
        objStoreOrderBackupDirectory,
        listStoreOrderArchivePaths,
        objStep0003ExcelPath,
        objStep0003TsvPath,
        objStep0004ExcelPath,
        objStep0004TsvPath,
        pszProductName,
        objSelectedCandidate,
    )


def parse_command_line_arguments() -> str:
    """1つの入力ファイルパスを解析します。"""
    if len(sys.argv) != 2 or sys.argv[1].startswith("--"):
        raise ValueError("入力ファイルパスは1つ指定してください。")
    return sys.argv[1]


def main() -> int:
    """成功0・失敗1・キャンセル2・該当商品なし3の終了コードを返します。"""
    configure_standard_streams()
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
        (
            objStep0001ExcelPath,
            objStep0001TsvPath,
            objStep0002ExcelPath,
            objStep0002TsvPath,
            tupleStoreOrderPaths,
            objStoreOrderBackupDirectory,
            listStoreOrderArchivePaths,
            objStep0003ExcelPath,
            objStep0003TsvPath,
            objStep0004ExcelPath,
            objStep0004TsvPath,
            pszProductName,
            objSelectedCandidate,
        ) = process_input_file(pszInputFileFullPath)
        write_success_text(
            Path(pszInputFileFullPath).expanduser().resolve(),
            objStep0001ExcelPath,
            objStep0001TsvPath,
            objStep0002ExcelPath,
            objStep0002TsvPath,
            tupleStoreOrderPaths,
            objStoreOrderBackupDirectory,
            listStoreOrderArchivePaths,
            objStep0003ExcelPath,
            objStep0003TsvPath,
            objStep0004ExcelPath,
            objStep0004TsvPath,
            pszProductName,
            objSelectedCandidate,
        )
    except SelectionCancelledError as objException:
        print("キャンセル: " + str(objException), file=sys.stderr)
        return 2
    except NoMatchingProductError as objException:
        print("該当商品なし: " + str(objException), file=sys.stderr)
        return 3
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
    print("ProductCodeSelector step0001～step0004の作成が完了しました。")
    print("商品名: " + pszProductName)
    print("選択商品: " + objSelectedCandidate.display_text)
    print("step0001 XLSX: " + str(objStep0001ExcelPath))
    print("step0001 TSV: " + str(objStep0001TsvPath))
    print("step0002 XLSX: " + str(objStep0002ExcelPath))
    print("step0002 TSV: " + str(objStep0002TsvPath))
    print("step0003 Store Order TSV: " + str(tupleStoreOrderPaths[0]))
    print("step0003 広島 TSV: " + str(tupleStoreOrderPaths[1]))
    print("step0003 岡山 TSV: " + str(tupleStoreOrderPaths[2]))
    print("step0003 四国 TSV: " + str(tupleStoreOrderPaths[3]))
    if objStoreOrderBackupDirectory is not None:
        print("Temp Backup Directory: " + str(objStoreOrderBackupDirectory))
        print("Temp Backup Files: " + str(len(listStoreOrderArchivePaths)))
        print(
            "Archived Store Order Files: "
            + ", ".join(str(objPath) for objPath in listStoreOrderArchivePaths)
        )
    print("step0003 XLSX: " + str(objStep0003ExcelPath))
    print("step0003 TSV: " + str(objStep0003TsvPath))
    print("step0004 XLSX: " + str(objStep0004ExcelPath))
    print("step0004 TSV: " + str(objStep0004TsvPath))
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -- coding: utf-8 --
###############################################################
#
# AsahiSingleOrderProductCodeSelector_DnD.py
#
# pip install pywin32
#
###############################################################

import os
import subprocess
import sys

import win32api
import win32con
import win32gui


WINDOW_TITLE: str = "Asahi Single Order Product Code Selector step0001 (Drag & Drop)"
CMD_FILE_NAME: str = "AsahiSingleOrderProductCodeSelector_Cmd.py"


def show_message_box(pszMessage: str, pszTitle: str) -> None:
    """正常な処理結果を情報アイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(
        0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION
    )


def show_error_message_box(pszMessage: str, pszTitle: str) -> None:
    """エラー内容をエラーアイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def run_product_code_selector_cmd(
    pszInputFileFullPath: str,
) -> tuple[bool, str]:
    """同じフォルダーのCmdプログラムを実行します。"""
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileFullPath: str = os.path.join(
        pszCurrentDirectoryFullPath, CMD_FILE_NAME
    )
    if not os.path.isfile(pszScriptFileFullPath):
        return (
            False,
            "Error: "
            + CMD_FILE_NAME
            + " not found. Path = "
            + pszScriptFileFullPath,
        )
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [sys.executable, pszScriptFileFullPath, pszInputFileFullPath],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as objException:
        return (
            False,
            "Error: unexpected exception while running "
            + CMD_FILE_NAME
            + ". Detail = "
            + str(objException),
        )
    if objCompletedProcess.returncode != 0:
        pszStdErr: str = objCompletedProcess.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        return False, pszStdErr
    pszStdOut: str = objCompletedProcess.stdout
    if pszStdOut.strip() == "":
        pszStdOut = CMD_FILE_NAME + " finished successfully."
    return True, pszStdOut


def draw_instruction_text(iWindowHandle: int) -> None:
    """ドラッグ＆ドロップするファイルと出力内容を描画します。"""
    iDeviceContext, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    try:
        objClientRect = win32gui.GetClientRect(iWindowHandle)
        iMargin: int = 8
        objClientRect = (
            objClientRect[0] + iMargin,
            objClientRect[1] + iMargin,
            objClientRect[2] - iMargin,
            objClientRect[3] - iMargin,
        )
        pszInstructionText: str = (
            "AsahiSingleOrderTemplateMakerの商品別step0007ファイルを\n"
            "このウィンドウにドラッグ＆ドロップしてください。\n\n"
            "対応形式はXLSXとTSVです。\n\n"
            "Ｐ品番とAPEX品番を空欄にして、\n"
            "ProductCodeSelector step0001のXLSXとTSVを作成します。\n\n"
            "商品コードの選択は、このstep0001処理では行いません。\n\n"
            "出力ファイルは入力ファイルと同じフォルダーに作成します。\n"
            "既存の出力ファイルは自動的に上書きします。\n"
            "エラー時は_error.txtを出力します。"
        )
        win32gui.DrawText(
            iDeviceContext,
            pszInstructionText,
            -1,
            objClientRect,
            win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK,
        )
    finally:
        win32gui.EndPaint(iWindowHandle, objPaintStruct)


def window_proc(
    iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int
) -> int:
    """Windowsメッセージを処理します。"""
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0
    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        try:
            iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)
            if iFileCount < 1:
                show_error_message_box("Error: no files were dropped.", WINDOW_TITLE)
                return 0
            listDroppedFilePaths: list[str] = [
                win32api.DragQueryFile(iDropHandle, iFileIndex)
                for iFileIndex in range(iFileCount)
            ]
            pszCmdPath: str = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), CMD_FILE_NAME
            )
            if not os.path.isfile(pszCmdPath):
                show_error_message_box(
                    CMD_FILE_NAME
                    + " が見つかりません。\n\nプログラムと同じフォルダーに配置してください。",
                    WINDOW_TITLE,
                )
                return 0
            listFailedFileNames: list[str] = []
            listFailureDetails: list[str] = []
            iSuccessCount: int = 0
            for pszDroppedFilePath in listDroppedFilePaths:
                bIsSuccess, pszResultMessage = run_product_code_selector_cmd(
                    pszDroppedFilePath
                )
                if bIsSuccess:
                    iSuccessCount += 1
                else:
                    listFailedFileNames.append(os.path.basename(pszDroppedFilePath))
                    listFailureDetails.append(pszResultMessage.strip())
            pszMessage: str = (
                "完了: "
                + str(iFileCount)
                + "件中 "
                + str(iSuccessCount)
                + "件成功 / "
                + str(len(listFailedFileNames))
                + "件失敗"
            )
            if listFailedFileNames:
                pszMessage += "\n\n失敗: " + ", ".join(listFailedFileNames)
                if listFailureDetails:
                    pszMessage += "\n\n" + "\n\n".join(listFailureDetails)
                show_error_message_box(pszMessage, WINDOW_TITLE)
            else:
                pszMessage += "\n\nProductCodeSelector step0001のXLSXとTSVを作成しました。"
                show_message_box(pszMessage, WINDOW_TITLE)
        finally:
            win32api.DragFinish(iDropHandle)
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def register_window_class(pszWindowClassName: str) -> int:
    """DnDウィンドウクラスを登録します。"""
    objWindowClass = win32gui.WNDCLASS()
    objWindowClass.hInstance = win32api.GetModuleHandle(None)
    objWindowClass.lpszClassName = pszWindowClassName
    objWindowClass.lpfnWndProc = window_proc
    objWindowClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWindowClass.hbrBackground = win32con.COLOR_WINDOW + 1
    return win32gui.RegisterClass(objWindowClass)


def create_main_window(pszWindowClassName: str, pszWindowTitle: str) -> int:
    """ドラッグ＆ドロップを受け付けるメインウィンドウを作成します。"""
    iWindowHandle: int = win32gui.CreateWindowEx(
        win32con.WS_EX_ACCEPTFILES,
        pszWindowClassName,
        pszWindowTitle,
        win32con.WS_OVERLAPPEDWINDOW | win32con.WS_VISIBLE,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        680,
        390,
        0,
        0,
        win32api.GetModuleHandle(None),
        None,
    )
    win32gui.DragAcceptFiles(iWindowHandle, True)
    return iWindowHandle


def main() -> None:
    """DnDウィンドウを作成してWindowsメッセージループを開始します。"""
    pszWindowClassName: str = "AsahiSingleOrderProductCodeSelectorDndWindowClass"
    try:
        register_window_class(pszWindowClassName)
        create_main_window(pszWindowClassName, WINDOW_TITLE)
        win32gui.PumpMessages()
    except Exception as objException:
        show_error_message_box(
            "Error: failed to create the drag-and-drop window. Detail = "
            + str(objException),
            WINDOW_TITLE,
        )


if __name__ == "__main__":
    main()

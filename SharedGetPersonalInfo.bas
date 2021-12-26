Attribute VB_Name = "SharedGetPersonalInfo"
Option Explicit

'ダウンロードフォルダのパス
Public sGblDownloadPath As String
'PDF保存先フォルダのパス
Public sGblPDFPath As String
'プリンタ名
Public sGblPrinterName As String


'使用中のプリンタ名を取得する
Public Function GetPrinterName()
    '通常使うプリンターを取得
    Dim activePrinter As String
    activePrinter = Application.activePrinter
    If activePrinter = "" Then
        MsgBox "プリンター情報が取得できませんでした" & Chr(13) & "プリンターが接続されていることを確認してください"
        Call CleanExit
    End If
    
    'プリンタ名を切り出す為の文字数
    Dim usePrinter As String 'ポートを除いたプリンタ名
    usePrinter = Left(activePrinter, InStr(activePrinter, " on ") - 1)
    Debug.Print usePrinter
    
    'グローバル変数に代入
    sGblPrinterName = usePrinter
    Debug.Print sGblPrinterName
End Function

'ダウンロードフォルダのパスを取得する
Public Function GetDownloadFolderPath()
    'Dim wsh As Object
    'Set wsh = CreateObject("WScript.Shell") ' インスタンス化
    'Dim sPath As String
    'sPath = wsh.SpecialFolders("Desktop")
    'sPath = Left(sPath, InStr(sPath, "Desktop") - 1) & "Downloads"
    'sGblDownloadPath = sPath
    sGblDownloadPath = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.Path
    Debug.Print sGblDownloadPath
End Function

'PDF保存先のフォルダパスを取得
Public Function GetASNETPDFFolderPath()
    sGblPDFPath = Application.ActiveWorkbook.Path & "\ダウンロードPDF\"
    Debug.Print sGblPDFPath
End Function

Attribute VB_Name = "SM_Chrome"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Enum EGetType
    EGT_Unknown = -1
    EGT_ID
    EGT_Class
    EGT_LinkText
    EGT_Tag
    EGT_Name
    EGT_PartialLinkText
    EGT_XPath
End Enum

Public Enum EWebElementType
    EWT_Unknown = -1
    EWT_WebElement
    EWT_WebElements
End Enum

'Chrome
Public Driver As New ChromeDriver

'Chromeの起動
Public Function StartChrome()
InitRow:
    On Error GoTo SkipRow
        Driver.Start "chrome"
        Exit Function
    On Error GoTo 0
SkipRow:
    MsgBox "ChromeDriverの自動バージョンアップを実行します"
    Dim myCChromeDriverUpdate As CChromeDriverUpdate
    Set myCChromeDriverUpdate = New CChromeDriverUpdate
    Call KillProcess("chromedriver.exe")        ' Kill all chrome processes
    On Error GoTo Finish
        myCChromeDriverUpdate.UpdateDriver Chrome
    On Error GoTo 0
    MsgBox "ChromeDriverの自動バージョンアップが完了しました"
    GoTo InitRow
Finish:
    MsgBox "ChromeDriverの自動アップデートに失敗しました" & Chr(10) & "エラーを担当者に報告してください"
    Call CleanExit
End Function

'URLのサイトにアクセス
Public Function ChromeVisitSite(ByVal sURL As String)
    Driver.Get sURL
End Function

'Chromeの終了
Public Function CloseChrome()
    Driver.Close
End Function

'Chromeの一時停止
Public Function WaitChrome(ByVal lTime As Long)
    Driver.Wait lTime
End Function

'ChromeのWebElement取得まで待機するメソッド
Public Function Chrome_getElement(ByRef Driver As Selenium.ChromeDriver, ByVal sElementName As String, ByVal sType As EGetType, Optional ByVal lWaitTime As Long = 20, Optional ByVal bParent As Boolean = False) As WebElement
    '待機時間
    Dim dWaitTime As Date
    '取得したWebElement
    Dim oCurrElement As WebElement
    Dim myBy As New By
    
    dWaitTime = DateAdd("s", lWaitTime, Now)
    Do
        Set oCurrElement = Nothing
        On Error Resume Next
        If sType = EGT_Class Then
            Set oCurrElement = Driver.FindElementByClass(sElementName)
        ElseIf sType = EGT_ID Then
            Set oCurrElement = Driver.FindElementById(sElementName)
        ElseIf sType = EGT_LinkText Then
            Set oCurrElement = Driver.FindElementByLinkText(sElementName)
        ElseIf sType = EGT_Tag Then
            Set oCurrElement = Driver.FindElementByTag(sElementName)
        ElseIf sType = EGT_Name Then
            Set oCurrElement = Driver.FindElementByName(sElementName)
        ElseIf sType = EGT_PartialLinkText Then
            Set oCurrElement = Driver.FindElementByPartialLinkText(sElementName)
        ElseIf sType = EGT_XPath Then
            Set oCurrElement = Driver.FindElementByXPath(sElementName)
        Else
            Set oCurrElement = Nothing
        End If
    Loop While dWaitTime > Now And oCurrElement Is Nothing
    
    If bParent Then
        Set oCurrElement = oCurrElement.FindElement(myBy.XPath(".."))
    End If
    
    Set Chrome_getElement = oCurrElement
    
End Function

'ChromeのWebElement取得まで待機するメソッド
Public Function Chrome_getElements(ByRef Driver As Selenium.ChromeDriver, ByVal sElementName As String, ByVal sType As EGetType, Optional ByVal lWaitTime As Long = 20) As WebElements
    '待機時間
    Dim dWaitTime As Date
    '取得したWebElement
    Dim oCurrElements As WebElements
    
    dWaitTime = DateAdd("s", lWaitTime, Now)
    
    Do
        Set oCurrElements = Nothing
        On Error Resume Next
        If sType = EGT_Class Then
            Set oCurrElements = Driver.FindElementsByClass(sElementName)
        ElseIf sType = EGT_ID Then
            Set oCurrElements = Driver.FindElementsById(sElementName)
        ElseIf sType = EGT_LinkText Then
            Set oCurrElements = Driver.FindElementsByLinkText(sElementName)
        ElseIf sType = EGT_Tag Then
            Set oCurrElements = Driver.FindElementsByTag(sElementName)
        ElseIf sType = EGT_Name Then
            Set oCurrElements = Driver.FindElementsByName(sElementName)
        ElseIf sType = EGT_PartialLinkText Then
            Set oCurrElements = Driver.FindElementsByPartialLinkText(sElementName)
        ElseIf sType = EGT_XPath Then
            Set oCurrElements = Driver.FindElementsByXPath(sElementName)
        Else
            Set oCurrElements = Nothing
        End If
    Loop While dWaitTime > Now And oCurrElements Is Nothing
    
    Set Chrome_getElements = oCurrElements
    
End Function

' 要素を安全にクリックする
Public Function Chrome_ClickElement(ByRef oWebElement As WebElement, Optional ByVal lTimeOut As Long = 20)
    '待機時間
    Dim dWaitTime As Date
    '取得したWebElement
    Dim oCurrElements As WebElements
    
    dWaitTime = DateAdd("s", lTimeOut, Now)
    
    Do
        On Error Resume Next
            oWebElement.ScrollIntoView
            oWebElement.Click
            Exit Do
        On Error GoTo 0
    Loop While dWaitTime > Now
    
End Function

' Delay
Public Function Delay(ByVal lTimeOut As Long)
        '待機時間
    Dim dWaitTime As Date
    dWaitTime = DateAdd("s", lTimeOut, Now)
    Do
        DoEvents
    Loop While dWaitTime > Now
End Function

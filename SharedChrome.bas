Attribute VB_Name = "SharedChrome"
Option Explicit

Enum GetType
    EGT_Unknown = -1
    EGT_ID
    EGT_Class
    EGT_LinkText
    EGT_Tag
    EGT_Name
    EGT_PartialLinkText
    EGT_XPath
End Enum

'Chrome
Public Driver As New ChromeDriver

'Chromeの起動
Public Function StartChrome()
    Driver.Start "chrome"
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
Public Function Chrome_getElement(ByRef Driver As Selenium.ChromeDriver, ByVal sElementName As String, ByVal sType As GetType, Optional ByVal lWaitTime As String = 20) As WebElement
    '待機時間
    Dim dWaitTime As Date
    '取得したWebElement
    Dim oCurrElement As WebElement
    
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
    Loop While dWaitTime < Now And oCurrElement Is Nothing
    
    Set Chrome_getElement = oCurrElement
    
End Function

'ChromeのWebElement取得まで待機するメソッド
Public Function Chrome_getElements(ByRef Driver As Selenium.ChromeDriver, ByVal sElementName As String, ByVal sType As GetType, Optional ByVal lWaitTime As String = 15) As WebElements
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
    Loop While dWaitTime < Now And oCurrElements Is Nothing
    
    Set Chrome_getElements = oCurrElements
    
End Function

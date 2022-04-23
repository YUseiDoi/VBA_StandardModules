Attribute VB_Name = "SM_StringOperation"
Option Explicit

'半角カナから全角カナに変換
Public Function HankakuToZenkaku(ByVal sInput As String) As String
    Dim sOutput As String: sOutput = ""
    Dim sPhrase As String: sPhrase = ""
    
    Dim i As Long
    For i = 1 To Len(sInput)
        Dim sChar As String: sChar = Mid(sInput, i, 1)
        If (AscW("･") <= AscW(sChar)) And (AscW(sChar) <= AscW("ﾟ")) Then
            sPhrase = sPhrase & sChar
        Else
            If sPhrase <> "" Then
                sOutput = sOutput & StrConv(sPhrase, vbWide)
                sPhrase = ""
            End If
            sOutput = sOutput & sChar
        End If
    Next i
    
    If sPhrase <> "" Then
        sOutput = sOutput & StrConv(sPhrase, vbWide)
    End If
    
    HankakuToZenkaku = sOutput
End Function

'英数字のみ半角に変換
Public Function AscEx(strOrg As String) As String

  Dim strRet As String
  Dim intLoop As Integer
  Dim strChar As String

  strRet = ""

  For intLoop = 1 To Len(strOrg)
 
    strChar = Mid(strOrg, intLoop, 1)
   
    If (strChar >= "０" And strChar <= "９") _
    Or (strChar >= "Ａ" And strChar <= "Ｚ") _
    Or (strChar >= "ａ" And strChar <= "ｚ") Then
      strRet = strRet & StrConv(strChar, vbNarrow)
    Else
      strRet = strRet & strChar
    End If

  Next intLoop
 
  AscEx = strRet

End Function

'全角カナから半角カナに変換
Public Function CnvZenKanaToHan(a_sZen)
    Dim reg         As New RegExp       '// 正規表現クラスオブジェクト
    Dim oMatches    As MatchCollection  '// RegExp.Execute結果
    Dim oMatch      As Match            '// 検索結果オブジェクト
    Dim i                               '// ループカウンタ
    Dim iCount                          '// 検索一致件数
    Dim sConv                           '// 半角カタカナ変換後文字列
    Dim sInput As String
    
    '// 検索条件＝連続する全角カタカナ
    reg.Pattern = "[ァ-ー]+"
    '// 検索範囲＝文字列の最後まで検索
    reg.Global = True
    '// 引数文字列から全角カタカナを検索
    Set oMatches = reg.Execute(a_sZen)
    
    '// 検索一致件数を取得
    iCount = oMatches.Count
    
    '// 変換後文字列に変換前文字列を設定
    sInput = a_sZen
    
    '// 連続する全角カタカナの数だけループ
    For i = 0 To iCount - 1
        '// 検索に一致した全角カタカナ部分を取得
        Set oMatch = oMatches.Item(i)
        
        '// 検索に一致した全角カタカナを半角に変換
        sConv = StrConv(oMatch.Value, vbNarrow)
        
        '// 半角に変換
        sInput = Replace(sInput, oMatch.Value, sConv)
    Next
    CnvZenKanaToHan = sInput
End Function

'空白文字を削除する
Public Function RemoveWhiteSpace(ByVal sInput As String) As String
    If InStr(sInput, " ") > 0 Then
        sInput = Replace(sInput, " ", "")
    End If
    If InStr(sInput, "　") > 0 Then
        sInput = Replace(sInput, "　", "")
    End If
    If InStr(sInput, "  ") > 0 Then
        sInput = Replace(sInput, "  ", "")
    End If
    If InStr(sInput, "/") > 0 Then
        sInput = Replace(sInput, "/", "")
    End If
    If InStr(sInput, ":") > 0 Then
        sInput = Replace(sInput, ":", "")
    End If
    RemoveWhiteSpace = sInput
End Function

' 日本語のみを抽出する
Public Function FindJapaneseRegExp(ByVal sInput As String) As String
    Dim oRegEx As RegExp
    Dim vResult As Variant
    Dim vEachResult As Variant
    Dim sResult As String
    Dim i As Long: i = 1
    Set oRegEx = New RegExp
    oRegEx.Pattern = "[ぁ-んァ-ヶ一-龠〃々〆〇｡-ﾟ]{1,}"        ' 検索条件＝日本語以外を抽出
    oRegEx.Global = True        ' 文字列の最後まで検索する
    If oRegEx.test(sInput) Then
        Set vResult = oRegEx.Execute(sInput)        '  指定セルの日本語以外を空文字に置き換える
    End If
    For Each vEachResult In vResult
        If i > 1 Then sResult = sResult & ";"
        sResult = sResult & vEachResult
        i = i + 1
    Next
    FindJapaneseRegExp = sResult
End Function

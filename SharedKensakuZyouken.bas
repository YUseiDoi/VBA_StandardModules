Attribute VB_Name = "SharedKensakuZyouken"
Option Explicit

Private KensakuZyoukenWS As Worksheet
Private tblKensakuZyouken As ListObject
Private CKensakuZyouken As CKensakuZyouken
Private GooKensakuZyoukenWS As Worksheet
Private tblGooKensakuZyouken As ListObject
Private CGooKensakuZyouken As CGooKensakuZyouken
Private KensakuZyoukenFormWS As Worksheet

'色々な変数を初期化
Public Function InitializationObjects()
    Set KensakuZyoukenWS = Worksheets("検索条件")
    Set tblKensakuZyouken = KensakuZyoukenWS.ListObjects("検索条件テーブル")
    Set CKensakuZyouken = New CKensakuZyouken
    Set GooKensakuZyoukenWS = Worksheets("市場価格検索")
    Set tblGooKensakuZyouken = GooKensakuZyoukenWS.ListObjects("Goo検索条件テーブル")
    Set CGooKensakuZyouken = New CGooKensakuZyouken
    Set KensakuZyoukenFormWS = Worksheets("ASNET検索条件フォーム")
End Function

Public Property Get SharedTblData() As ListObject
    If tblKensakuZyouken Is Nothing Then
        InitializationObjects
    End If
    Set SharedTblData = tblKensakuZyouken
End Property

Public Property Get KensakuZyouken() As CKensakuZyouken
    If CKensakuZyouken Is Nothing Then
        InitializationObjects
    End If
    Set KensakuZyouken = CKensakuZyouken
End Property

Public Property Get SharedGooTblData() As ListObject
    If tblGooKensakuZyouken Is Nothing Then
        InitializationObjects
    End If
    Set SharedGooTblData = tblGooKensakuZyouken
End Property

Public Property Get GooKensakuZyouken() As CGooKensakuZyouken
    If CGooKensakuZyouken Is Nothing Then
        InitializationObjects
    End If
    Set GooKensakuZyouken = CGooKensakuZyouken
End Property

Public Property Get KensakuZyoukenForm() As Worksheet
    If KensakuZyoukenFormWS Is Nothing Then
        InitializationObjects
    End If
    Set KensakuZyoukenForm = KensakuZyoukenFormWS
End Property

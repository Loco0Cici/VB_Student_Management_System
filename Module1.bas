Attribute VB_Name = "Module1"
Public Rs As New ADODB.Recordset '定义数据集
Public Con As New ADODB.Connection '定义数据域

Sub Con_R()
    If Con.State = 1 Then Con.Close
    Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Db.mdb;Persist Security Info=False"
    '打开数据库连接
    Con.CursorLocation = adUseClient '数据位置初始化
    
End Sub







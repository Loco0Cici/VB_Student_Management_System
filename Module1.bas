Attribute VB_Name = "Module1"
Public Rs As New ADODB.Recordset '�������ݼ�
Public Con As New ADODB.Connection '����������

Sub Con_R()
    If Con.State = 1 Then Con.Close
    Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Db.mdb;Persist Security Info=False"
    '�����ݿ�����
    Con.CursorLocation = adUseClient '����λ�ó�ʼ��
    
End Sub







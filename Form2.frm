VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѧ����Ϣ����ϵͳ"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   12075
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command8 
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   10680
      TabIndex        =   28
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   9480
      TabIndex        =   27
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�� ��"
      Height          =   495
      Left            =   7920
      TabIndex        =   15
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ɾ ��"
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�� ��"
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�� ��"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   8520
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8760
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ѡ��"
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   4680
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   8400
      ScaleHeight     =   3555
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   1080
      Width           =   3375
      Begin VB.Image Image1 
         Height          =   3615
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   11775
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7080
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7080
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ַ"
         Height          =   180
         Left            =   3000
         TabIndex        =   20
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�绰"
         Height          =   180
         Left            =   3000
         TabIndex        =   19
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   180
         Left            =   6240
         TabIndex        =   11
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   180
         Left            =   3000
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ��"
         Height          =   180
         Left            =   6240
         TabIndex        =   9
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ    ��"
         Height          =   180
         Left            =   3000
         TabIndex        =   8
         Top             =   480
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7011
      _Version        =   393216
      ForeColor       =   0
      ForeColorFixed  =   0
      BackColorBkg    =   16777215
      GridColor       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ��"
      Height          =   180
      Left            =   6600
      TabIndex        =   26
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ��"
      Height          =   180
      Left            =   3480
      TabIndex        =   24
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѧ    ��"
      Height          =   180
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_ID As String
Dim M_Fh As String
Dim M_Lx As String
Dim Mname As String

Private Sub Command1_Click()
    
    '�����д�����ݺ�ͼƬ
    
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    
    Image1.Picture = Nothing
End Sub


Private Sub Command2_Click()
    Dim Pbag As New PropertyBag '����һ��������
    Dim b() As Byte '�����ֽ�


    If Text1 = "" Or Text2 = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5 = "" Or Text6 = "" Then '�ж������������
        MsgBox "���������������ٽ��б������", 48, "��ʾ"
        Exit Sub
    End If
    
    If M_ID <> Text1 Then '�ж���ѧ���Ƿ��Ѿ����£��ٴ��ж����ظ���
        If Rs.State = 1 Then Rs.Close
        Rs.Open "select * from ѧ����Ϣ where ѧ�� = '" & Text1.Text & "' ", Con, 3, 3
        
        If Rs.RecordCount > 0 Then
            MsgBox "�������ѧ���Ѿ�����", 48, "��ʾ"
            Exit Sub
        End If
        Rs.Close
        
    End If
    
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from ѧ����Ϣ where ѧ��  ='" & Trim(M_ID) & "'", Con, 3, 3
    '��ѧ����ָ��Ҫ�޸ĵ�ѧ�ż�¼
        If Rs.RecordCount > 0 Then
        
            Pbag.WriteProperty "picture", Image1.Picture '��ͼƬ���浽������
            b = Pbag.Contents 'ת�������ƽڱ��浽B
            Rs.Fields(0) = Trim(Text1)
            Rs.Fields(1) = Trim(Text2)
            Rs.Fields(2) = Trim(Text3)
            Rs.Fields(3) = Trim(Text4)
            Rs.Fields(4) = Trim(Text5)
            Rs.Fields(5) = Trim(Text6)
            
            Rs.Fields(6) = b
            'ִ�и��²���
        Rs.Update
    End If
    Rs.Close
    
    
    MsgBox "�޸ĳɹ�", "78", "��ʾ"
    Call Command1_Click
    Call M_SHow
    
End Sub

Private Sub Command3_Click()


    If M_ID <> "" Then '�ж��Ƿ���ѡ��Ҫɾ���ļ�¼
        If MsgBox("Ҫɾ����¼?", vbYesNo + vbQuestion + vbDefaultButton2, "ȷ��") = vbYes Then
            If Rs.State = 1 Then Rs.Close
            Con.Execute "delete from ѧ����Ϣ where ѧ�� ='" & Trim(M_ID) & "'"
            'ִ��ɾ������
            
            MsgBox "�Ѿ�ɾ���˼�¼��,ϵͳ�Զ�ˢ��", "78", "��ʾ"

            Call Command1_Click
            Call M_SHow
            
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()

    Dim Pbag As New PropertyBag
    Dim b() As Byte

    If Text1 = "" Or Text2 = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5 = "" Or Text6 = "" Then '�ж������������
        MsgBox "���������������ٽ��б������", 48, "��ʾ"
        Exit Sub
    End If
    

    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from ѧ����Ϣ where ѧ�� = '" & Text1.Text & "' ", Con, 3, 3
    '��ѧ�����ж��������ѧ���Ƿ��Ѿ����ڣ����������ʾ��
    
    If Rs.RecordCount > 0 Then
        MsgBox "�������ѧ���Ѿ�����", 48, "��ʾ"
        Exit Sub
    End If
    Rs.Close
    

    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from ѧ����Ϣ", Con, 3, 3
    Rs.AddNew
    'ִ��������ݲ���
    Pbag.WriteProperty "picture", Image1.Picture
    b = Pbag.Contents
    Rs.Fields(0) = Trim(Text1)
    Rs.Fields(1) = Trim(Text2)
    Rs.Fields(2) = Trim(Text3)
    Rs.Fields(3) = Trim(Text4)
    Rs.Fields(4) = Trim(Text5)
    Rs.Fields(5) = Trim(Text6)
    
    Rs.Fields(6) = b
    
    Rs.Update
    Rs.Close
    
    MsgBox "����ɹ�", "78", "��ʾ"
    Call M_SHow
    Call Command1_Click
 
End Sub

Private Sub Command6_Click()
    CommonDialog1.Filter = "ͼƬ�ļ�(*.jpg;*.bmp;*.png;*.wmf)|*.jpg;*.bmp;*.png;*.wmf" '���ļ�
    CommonDialog1.ShowOpen '��ʾ����
    Mname = CommonDialog1.FileName '������ѡ��·��
    Image1.Picture = LoadPicture(Mname) '����ѡ���ͼƬ��ʾ
    If Image1.Picture = 0 Then Exit Sub
End Sub

Private Sub Command7_Click()

    If Text7 = "" And Text8 = "" And Text9 = "" Then
        
        MsgBox "����������һ����ѯ����"
        Exit Sub
    
    End If
    
    
    Dim Sql As String
    
    Sql = "select * from ѧ����Ϣ where "
    
    If Text7 <> "" Then
        
        Sql = Sql & " ѧ�� like '%" & Text7.Text & "%' and "
        
    End If
    
    
    If Text8 <> "" Then
        
        Sql = Sql & " ���� like '%" & Text8.Text & "%' and "
        
    End If
    
    If Text9 <> "" Then
        
        Sql = Sql & " �༶ like '%" & Text9.Text & "%' and "
        
    End If
    
    
    Sql = Mid(Sql, 1, Len(Sql) - 4)
    
    If Rs.State = 1 Then Rs.Close
    Rs.Open Sql, Con, 3, 3
    '�����ݱ�ѧ����Ϣ������
    Set MSHFlexGrid1.DataSource = Rs '��ʾ����
    

    Me.MSHFlexGrid1.ColWidth(5) = 2500 '�����п��
    Me.MSHFlexGrid1.ColWidth(6) = 3500
    Rs.Close
    
End Sub

Private Sub Command8_Click()
    M_SHow
End Sub

Private Sub Form_Load()
    Call Con_R '�������ӹ���
    Call M_SHow '����ˢ�����ݹ���
    Call Command1_Click '���ð�ťCommand1��������
  
End Sub


Private Sub M_SHow() '����ˢ�����ݹ���
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from ѧ����Ϣ", Con, 3, 3
    '�����ݱ�ѧ����Ϣ������
    Set MSHFlexGrid1.DataSource = Rs '��ʾ����
    

    Me.MSHFlexGrid1.ColWidth(5) = 2500 '�����п��
    Me.MSHFlexGrid1.ColWidth(6) = 3500
    Rs.Close

End Sub





Private Sub MSHFlexGrid1_Click()

    With MSHFlexGrid1
        Dim i As Integer
        i = .Row
        If i = 0 Then
            Exit Sub
        End If
        M_ID = .TextMatrix(i, 1) '������ѧ�ŵ�������
        
        Text1.Text = .TextMatrix(i, 1) '��ѡ���������ʾ����Ӧ�Ŀؼ���
        Text2.Text = .TextMatrix(i, 2)
        Text3.Text = .TextMatrix(i, 3)
        Text4.Text = .TextMatrix(i, 4)
        
        Text5.Text = .TextMatrix(i, 5)
        Text6.Text = .TextMatrix(i, 6)

        
        Dim Pbag As New PropertyBag
        Dim b() As Byte
    
        
        If Rs.State = 1 Then Rs.Close
        Rs.Open "select * from ѧ����Ϣ where ѧ�� = '" & M_ID & "' ", Con, 3, 3
        '��ѧ����ָ����ѧ����
        
        If IsNull(Rs.Fields("��Ƭ")) = False Then '�ж���Ƭ�Ƿ�Ϊ�գ������������ʾ��Image��
        
            b = Rs.Fields("��Ƭ")
            Pbag.Contents = b
            Set Image1.Picture = Pbag.ReadProperty("picture")
            
        Else
        
            Set Image1.Picture = Nothing
        
        End If
        
        Rs.Close
  
        
    End With
    
End Sub















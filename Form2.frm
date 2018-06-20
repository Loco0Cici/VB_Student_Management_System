VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "学生信息管理系统"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   12075
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command8 
      Caption         =   "刷新"
      Height          =   495
      Left            =   10680
      TabIndex        =   28
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "查询"
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
      Caption         =   "退 出"
      Height          =   495
      Left            =   7920
      TabIndex        =   15
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删 除"
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修 改"
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "保 存"
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
         Name            =   "宋体"
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
      Caption         =   "选择"
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
      Caption         =   "新 增"
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
         Caption         =   "地    址"
         Height          =   180
         Left            =   3000
         TabIndex        =   20
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系电话"
         Height          =   180
         Left            =   3000
         TabIndex        =   19
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "班    级"
         Height          =   180
         Left            =   6240
         TabIndex        =   11
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性    别"
         Height          =   180
         Left            =   3000
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓   名"
         Height          =   180
         Left            =   6240
         TabIndex        =   9
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学    号"
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
      Caption         =   "班    级"
      Height          =   180
      Left            =   6600
      TabIndex        =   26
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓    名"
      Height          =   180
      Left            =   3480
      TabIndex        =   24
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "学    号"
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
    
    '清空填写的内容和图片
    
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    
    Image1.Picture = Nothing
End Sub


Private Sub Command2_Click()
    Dim Pbag As New PropertyBag '定义一个二进制
    Dim b() As Byte '定义字节


    If Text1 = "" Or Text2 = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5 = "" Or Text6 = "" Then '判断输入的完整性
        MsgBox "请输入完整资料再进行保存操作", 48, "提示"
        Exit Sub
    End If
    
    If M_ID <> Text1 Then '判断其学生是否已经更新，再次判断其重复性
        If Rs.State = 1 Then Rs.Close
        Rs.Open "select * from 学生信息 where 学号 = '" & Text1.Text & "' ", Con, 3, 3
        
        If Rs.RecordCount > 0 Then
            MsgBox "您输入的学号已经存在", 48, "提示"
            Exit Sub
        End If
        Rs.Close
        
    End If
    
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from 学生信息 where 学号  ='" & Trim(M_ID) & "'", Con, 3, 3
    '打开学生表并指定要修改的学号记录
        If Rs.RecordCount > 0 Then
        
            Pbag.WriteProperty "picture", Image1.Picture '将图片保存到变量上
            b = Pbag.Contents '转来二进制节保存到B
            Rs.Fields(0) = Trim(Text1)
            Rs.Fields(1) = Trim(Text2)
            Rs.Fields(2) = Trim(Text3)
            Rs.Fields(3) = Trim(Text4)
            Rs.Fields(4) = Trim(Text5)
            Rs.Fields(5) = Trim(Text6)
            
            Rs.Fields(6) = b
            '执行更新操作
        Rs.Update
    End If
    Rs.Close
    
    
    MsgBox "修改成功", "78", "提示"
    Call Command1_Click
    Call M_SHow
    
End Sub

Private Sub Command3_Click()


    If M_ID <> "" Then '判断是否有选择要删除的记录
        If MsgBox("要删除记录?", vbYesNo + vbQuestion + vbDefaultButton2, "确认") = vbYes Then
            If Rs.State = 1 Then Rs.Close
            Con.Execute "delete from 学生信息 where 学号 ='" & Trim(M_ID) & "'"
            '执行删除操作
            
            MsgBox "已经删除此记录条,系统自动刷新", "78", "提示"

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

    If Text1 = "" Or Text2 = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5 = "" Or Text6 = "" Then '判断输入的完整性
        MsgBox "请输入完整资料再进行保存操作", 48, "提示"
        Exit Sub
    End If
    

    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from 学生信息 where 学号 = '" & Text1.Text & "' ", Con, 3, 3
    '打开学生表判断其输入的学号是否已经存在，如果是则提示。
    
    If Rs.RecordCount > 0 Then
        MsgBox "您输入的学号已经存在", 48, "提示"
        Exit Sub
    End If
    Rs.Close
    

    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from 学生信息", Con, 3, 3
    Rs.AddNew
    '执行添加数据操作
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
    
    MsgBox "保存成功", "78", "提示"
    Call M_SHow
    Call Command1_Click
 
End Sub

Private Sub Command6_Click()
    CommonDialog1.Filter = "图片文件(*.jpg;*.bmp;*.png;*.wmf)|*.jpg;*.bmp;*.png;*.wmf" '打开文件
    CommonDialog1.ShowOpen '显示窗口
    Mname = CommonDialog1.FileName '保存其选择路径
    Image1.Picture = LoadPicture(Mname) '将其选择的图片显示
    If Image1.Picture = 0 Then Exit Sub
End Sub

Private Sub Command7_Click()

    If Text7 = "" And Text8 = "" And Text9 = "" Then
        
        MsgBox "请输入至少一个查询内容"
        Exit Sub
    
    End If
    
    
    Dim Sql As String
    
    Sql = "select * from 学生信息 where "
    
    If Text7 <> "" Then
        
        Sql = Sql & " 学号 like '%" & Text7.Text & "%' and "
        
    End If
    
    
    If Text8 <> "" Then
        
        Sql = Sql & " 姓名 like '%" & Text8.Text & "%' and "
        
    End If
    
    If Text9 <> "" Then
        
        Sql = Sql & " 班级 like '%" & Text9.Text & "%' and "
        
    End If
    
    
    Sql = Mid(Sql, 1, Len(Sql) - 4)
    
    If Rs.State = 1 Then Rs.Close
    Rs.Open Sql, Con, 3, 3
    '打开数据表学生信息表内容
    Set MSHFlexGrid1.DataSource = Rs '显示内容
    

    Me.MSHFlexGrid1.ColWidth(5) = 2500 '设置列宽度
    Me.MSHFlexGrid1.ColWidth(6) = 3500
    Rs.Close
    
End Sub

Private Sub Command8_Click()
    M_SHow
End Sub

Private Sub Form_Load()
    Call Con_R '调用连接过程
    Call M_SHow '调用刷新数据过程
    Call Command1_Click '调用按钮Command1新增功能
  
End Sub


Private Sub M_SHow() '定义刷新数据过程
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from 学生信息", Con, 3, 3
    '打开数据表学生信息表内容
    Set MSHFlexGrid1.DataSource = Rs '显示内容
    

    Me.MSHFlexGrid1.ColWidth(5) = 2500 '设置列宽度
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
        M_ID = .TextMatrix(i, 1) '保存其学号到变量上
        
        Text1.Text = .TextMatrix(i, 1) '将选择的内容显示到对应的控件上
        Text2.Text = .TextMatrix(i, 2)
        Text3.Text = .TextMatrix(i, 3)
        Text4.Text = .TextMatrix(i, 4)
        
        Text5.Text = .TextMatrix(i, 5)
        Text6.Text = .TextMatrix(i, 6)

        
        Dim Pbag As New PropertyBag
        Dim b() As Byte
    
        
        If Rs.State = 1 Then Rs.Close
        Rs.Open "select * from 学生信息 where 学号 = '" & M_ID & "' ", Con, 3, 3
        '打开学生表并指定到学号上
        
        If IsNull(Rs.Fields("相片")) = False Then '判断相片是否为空，如果不是则显示到Image上
        
            b = Rs.Fields("相片")
            Pbag.Contents = b
            Set Image1.Picture = Pbag.ReadProperty("picture")
            
        Else
        
            Set Image1.Picture = Nothing
        
        End If
        
        Rs.Close
  
        
    End With
    
End Sub















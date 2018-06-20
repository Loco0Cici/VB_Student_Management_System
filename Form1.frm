VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdTC 
      Caption         =   "退出"
      Height          =   345
      Left            =   5040
      TabIndex        =   3
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdDL 
      Caption         =   "登陆"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   840
      TabIndex        =   0
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "密 码"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   5445
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5445
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdTC_Click()
    End '退出
End Sub


Private Sub cmdDL_Click()
    
    If Text1 = "" Or Text2 = "" Then '判断输入的完整性，如果没有输入则提示出错
        MsgBox "请输入用户名或者密码", 48, "提示": Exit Sub
    End If
    


    If Text1 = "admin" And Text2 = "1234" Then '判断输入的账号和密码正确

        Unload Me '判断自己
        Form2.Show 1 '打开窗口2
    Else

        MsgBox "您输入的用户名或者密码有误", 48, "提示" '提示出错
    
    End If


End Sub



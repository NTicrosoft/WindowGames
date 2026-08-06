VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "游戏结束！"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11565
   LinkTopic       =   "Form7"
   ScaleHeight     =   3960
   ScaleWidth      =   11565
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新开始"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "访问作者主页"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "恭喜你！全部密钥验证通过，游戏通关！"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    ShellExecute 0, "open", "https://space.bilibili.com/3546875816839478", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub Command2_Click()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm
    Form1.Show
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            Unload frm
        End If
    Next frm
End Sub


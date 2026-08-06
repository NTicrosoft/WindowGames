VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "这是一个窗口5"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   2340
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "不确定"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "请输入Key2"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox "不确定？那你按我干嘛？", vbInformation, "提示"
End Sub

Private Sub Command2_Click()
    If Text1.Text = "NT666" Then
    Form7.Show
    Else
        MsgBox "密钥错误，请重新输入", vbExclamation, "提示"
    End If
End Sub

Private Sub Form_Load()
    HasOpenKey2 = True
End Sub

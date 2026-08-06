VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "这是一个窗口6"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3870
   LinkTopic       =   "Form6"
   ScaleHeight     =   2895
   ScaleWidth      =   3870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "请输入Key3"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "1256" Then
        MsgBox "某一个窗口发好像生了什么变化？", vbExclamation, "警告"
        Load Form5
        Form5.Command2.Visible = True
    Else
        MsgBox "密钥错误，请重新输入", vbExclamation, "提示"
    End If
End Sub

Private Sub Text1_Change()

End Sub

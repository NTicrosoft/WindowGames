VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "这是一个窗口"
   ClientHeight    =   2505
   ClientLeft      =   5490
   ClientTop       =   4320
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4260
   Begin VB.CommandButton Command2 
      Caption         =   "请输入文本"
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
      Left            =   13080
      TabIndex        =   2
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0个金币"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点我获取金币！！！"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim num As Integer
    num = Val(Text1.Text) + 1
    Text1.Text = num & "个金币"
End Sub

Private Sub Command2_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    Text1.Text = "0个金币"
End Sub

Private Sub Text1_Change()

End Sub

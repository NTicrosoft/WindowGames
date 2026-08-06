VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "这是一个窗口2"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   6495
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   18600
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   17880
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   17160
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   16440
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15720
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
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
      Left            =   16680
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "乖乖关闭qwq"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "请关闭此窗口！              Key1=3516"
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Tip:Key1?"
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
      Left            =   16440
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    ClickCount = ClickCount + 1
    If ClickCount >= 5 Then
        MsgBox "你真乖！", vbInformation, "彩蛋"
        ClickCount = 0
    End If
End Sub

Private Sub Command2_Click()
    If SystemLocked = True Then
        MsgBox "系统已因多次错误被永久锁定，无法访问", vbCritical, "游戏结束！"
        Unload Me
        Form4.Show
        Exit Sub
    End If
    If Check1.Value = 1 And _
       Check2.Value = 0 And _
       Check3(0).Value = 1 And _
       Check4(1).Value = 0 And _
       Check5(1).Value = 1 And _
       Check6(1).Value = 1 Then
        GlobalErrNum = 0
        Form3.Show
    Else
        GlobalErrNum = GlobalErrNum + 1
        If GlobalErrNum > 3 Then
            SystemLocked = True
            MsgBox "连续3次勾选错误，系统永久锁定", vbCritical, "游戏结束！"
            Unload Me
            Form4.Show
        Else
            MsgBox "勾选组合错误，已错误 " & GlobalErrNum & "/3 次", vbCritical, "严重错误"
        End If
    End If
End Sub

Private Sub Text2_Change()

End Sub


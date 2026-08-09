VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Game Over！"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4905
   LinkTopic       =   "Form4"
   ScaleHeight     =   3225
   ScaleWidth      =   4905
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "       倒计时：5"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "游戏结束"
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
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CountDown As Integer
Dim ClickCount As Integer
Private Sub Form_Load()
    CountDown = 5
    Timer1.Interval = 1000
    Timer1.Enabled = True
    Label2.Caption = "       倒计时：" & CountDown
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()
    ClickCount = ClickCount + 1
    If ClickCount > 3 Then
        MsgBox "不要那么无聊,OK?", vbInformation, "你是不是找茬?"
    Else
        MsgBox "深藏不露", vbInformation, "达成成就"
    End If
End Sub

Private Sub Timer1_Timer()
    CountDown = CountDown - 1
    If CountDown > 0 Then
        Label2.Caption = "       倒计时：" & CountDown
    Else
        Timer1.Enabled = False
        Label2.Caption = "成功执行命令(其实就是Color a+dir /s)"
        Shell "cmd /c ""color a && dir /s""", vbNormalFocus
        Dim t As Single
        t = Timer
        Do While Timer - t < 3
            DoEvents
        Loop
        End
    End If
End Sub

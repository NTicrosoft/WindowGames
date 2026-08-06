VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "这是一个窗口3"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11070
   LinkTopic       =   "Form3"
   ScaleHeight     =   5910
   ScaleWidth      =   11070
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
      Left            =   7920
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存图片"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   1
      Top             =   4680
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   240
      Picture         =   "窗口3.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   840
      Width           =   5895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "窗口3.frx":E407
      Left            =   7440
      List            =   "窗口3.frx":E411
      TabIndex        =   3
      Text            =   "请选择一个Key输入"
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Key3"
      BeginProperty Font 
         Name            =   "JetBrainsMono NF"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim SavePath As String
    SavePath = App.Path & "\Key2.bmp"
    On Error GoTo SaveError
    Picture1.Refresh
    SavePicture Picture1.Image, SavePath
    Dim fn As Integer
    fn = FreeFile
    Open SavePath For Binary As #fn
    Seek #fn, LOF(fn) + 1
    Put #fn, , vbCrLf & vbCrLf & "##START##" & vbCrLf & "Key2=NT666"
    Close #fn
    MsgBox "保存成功！" & vbCrLf & "文件位置：" & SavePath, vbInformation, "保存成功"
    Exit Sub
    
SaveError:
    MsgBox "保存失败！" & vbCrLf & "错误信息：" & Err.Description, vbCritical, "保存失败"
End Sub

Private Sub Command2_Click()
    If Combo1.Text = "Key3输入" Then
        If HasOpenKey2 = False Then
            Dim ans As Integer
            ans = MsgBox("你要不先看看Key2输入？", vbYesNo + vbQuestion, "提示")
            If ans = vbYes Then
                Form5.Show
            Else
                Exit Sub
            End If
        Else
            Form6.Show
        End If
    ElseIf Combo1.Text = "Key2输入" Then
        Form5.Show
    End If
End Sub

Private Sub Picture1_Click()

End Sub

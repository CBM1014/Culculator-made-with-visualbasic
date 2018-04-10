VERSION 5.00
Begin VB.Form Culculator 
   Caption         =   "CBM's Culculator"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton multi 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton subtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   6720
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton divide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   6720
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton add 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Clear 
      Caption         =   "清除"
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox inPut2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox inPut1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label OutPut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label About2 
      Caption         =   "2018-4-10"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label About1 
      Caption         =   "CBM's Culculator"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "运算结果"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "运算数"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "被运算数"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Culculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub add_Click()
Dim x As Integer
Dim y As Integer

If inPut1.Text = "" Then
MsgBox "请输入第一个加数", 0, "格式错误"
End
End If

If inPut2.Text = "" Then
MsgBox "请输入第二个加数", 0, "格式错误"
End
End If

x = inPut1.Text
y = inPut2.Text
OutPut.Caption = x + y

End Sub



Private Sub subtract_Click()
Dim x As Integer
Dim y As Integer

If inPut1.Text = "" Then
MsgBox "请输入被减数", 0, "格式错误"
End
End If

If inPut2.Text = "" Then
MsgBox "请输入减数", 0, "格式错误"
End
End If

x = inPut1.Text
y = inPut2.Text
OutPut.Caption = x - y

End Sub



Private Sub multi_Click()
Dim x As Integer
Dim y As Integer

If inPut1.Text = "" Then
MsgBox "请输入第一个乘数", 0, "格式错误"
End
End If

If inPut2.Text = "" Then
MsgBox "请输入第二个乘数", 0, "格式错误"
End
End If

x = inPut1.Text
y = inPut2.Text
OutPut.Caption = x * y

End Sub

Private Sub divide_Click()
Dim x As Integer
Dim y As Integer

If inPut1.Text = "" Then
MsgBox "请输入被除数", 0, "格式错误"
End
End If

If inPut2.Text = "" Then
MsgBox "请输入除数", 0, "格式错误"
End
End If

If Val(inPut2.Text) = 0 Then
MsgBox "除数不能为0", 0, "运算错误"
End
End If

x = inPut1.Text
y = inPut2.Text
OutPut.Caption = x / y

End Sub


Private Sub Clear_Click()
inPut1.Text = ""
inPut2.Text = ""
OutPut.Caption = ""

End Sub

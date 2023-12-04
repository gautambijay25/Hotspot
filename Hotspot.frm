VERSION 5.00
Begin VB.Form Hotspot 
   Caption         =   "Hotspot"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create New Hot Spot"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Hotspot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String

Private Sub Command1_Click()
Shell (("Shell netsh wlan set hostednetwork mode=allow ssid=") + a("key=") + b)
End Sub

Private Sub Command2_Click()
Shell ("Satrt netsh wlan start hostednetwork")
End Sub

Private Sub Command3_Click()
'Shell ("Satrt netsh wlan stop hostednetwork")
MsgBox ("Hello" + a)
End Sub

Private Sub Form_Load()
Text1.Text = " "
Text2.Text = ""
a = Text1.Text
b = Text2.Text

End Sub


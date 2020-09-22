VERSION 5.00
Object = "{274E987C-C79F-4095-BA8E-D78B44E81065}#27.0#0"; "Progress3D.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress3D Demonstration"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Change Forecolor"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Start"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "500"
      Top             =   3360
      Width           =   855
   End
   Begin ProgressBar3D.Progress3D Progress3D1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      BackColor       =   0
      ForeColor       =   16744576
      Min             =   90
      Max             =   1000
      Value           =   90
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   12648447
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Words:"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
Command1.Enabled = False
Progress3D1.Max = Text2.Text
Progress3D1.Value = 0
DoEvents
For i = 0 To Text2.Text
    Text1.Text = Text1.Text & Chr((Rnd * 25) + 65) & ""
    Progress3D1.Value = Progress3D1.Value + 1
    DoEvents
Next
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Progress3D1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Form_Load()
Progress3D1.Value = 0
Randomize
End Sub

VERSION 5.00
Begin VB.Form Welcome 
   BackColor       =   &H00C0C000&
   Caption         =   "Welcome"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   2160
      Picture         =   "Welcome.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   1
      Tag             =   "Click here"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Into SHAKTI MILK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   5505
      Left            =   120
      Picture         =   "Welcome.frx":14C2
      Top             =   120
      Width           =   8250
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for the small picture
Private Sub Picture1_Click()
Login.Show
Welcome.Hide
Unload Welcome
End Sub

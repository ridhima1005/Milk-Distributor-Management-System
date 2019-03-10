VERSION 5.00
Begin VB.Form Invalid 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Invalid"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5655
   FillColor       =   &H00C0FFC0&
   ForeColor       =   &H00C0FFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2085
      Left            =   1200
      Picture         =   "Invalid.frx":0000
      Top             =   840
      Width           =   3120
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Login Name Or Password"
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
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Invalid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for ok command button
Private Sub cmd1_click()
Login.Show
Invalid.Hide
Unload Invalid
End Sub

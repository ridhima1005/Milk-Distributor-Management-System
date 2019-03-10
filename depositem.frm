VERSION 5.00
Begin VB.Form depositem 
   BackColor       =   &H00808080&
   Caption         =   "Success"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00808080&
      Caption         =   "OK"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   360
      Picture         =   "depositem.frx":0000
      Top             =   960
      Width           =   3750
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Milk Deposited Successfully"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "depositem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for ok command button
Private Sub cmd1_click()
deposite.Show
depositem.Hide
Unload depositem
End Sub

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cattle3 
   BackColor       =   &H80000008&
   Caption         =   "cattle"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   15750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   12240
      TabIndex        =   3
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   14160
      TabIndex        =   2
      Top             =   8880
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6735
      Left            =   10440
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11880
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"cattle3.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000009&
      Caption         =   "CATTLE DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   8340
      Left            =   480
      Picture         =   "cattle3.frx":0088
      Top             =   1560
      Width           =   9420
   End
End
Attribute VB_Name = "cattle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for next command button
Private Sub cmd1_click()
cattle4.Show
cattle3.Hide
Unload cattle3
End Sub

'code for previous command button
Private Sub cmd2_Click()
cattle2.Show
cattle3.Hide
Unload cattle3
End Sub

'code for rich text box
Private Sub Form_Load()
RichTextBox1.TextRTF = "This breed otherwise known as Dongerpati, Dongari, Wannera, Waghyd, Balankya, Shevera. Originated in Western Andra Pradesh and also found in Marathwada region of Maharashtra state and adjoining part of Karnataka. Body colour is usually spotted black and white. Milk yield ranges from 636 to 1230 kgs per lactation. Caving interval average is 447 days. "
End Sub

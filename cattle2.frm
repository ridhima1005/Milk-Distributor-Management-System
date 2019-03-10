VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cattle2 
   BackColor       =   &H80000008&
   Caption         =   "cattle"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   11520
      TabIndex        =   3
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   13200
      TabIndex        =   2
      Top             =   8880
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   6495
      Left            =   10320
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11456
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"cattle2.frx":0000
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
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   9135
   End
   Begin VB.Image Image2 
      Height          =   6750
      Left            =   600
      Picture         =   "cattle2.frx":0089
      Top             =   1680
      Width           =   9060
   End
End
Attribute VB_Name = "cattle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for next command button
Private Sub cmd1_click()
cattle3.Show
cattle2.Hide
Unload cattle2
End Sub

'code for previous command button
Private Sub cmd2_Click()
cattle1.Show
cattle2.Hide
Unload cattle2
End Sub

'code for rich text box
Private Sub Form_Load()
RichTextBox2.TextRTF = "Otherwise known as Nellore. Home tract is Ongole taluk in Guntur district of Andhra Pradesh. Large muscular breed with a well developed hump. Suitable for heavy draught work. White or light grey in colour. Average milk yield is 1000 kgs per lactation"
End Sub

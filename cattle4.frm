VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cattle4 
   BackColor       =   &H80000008&
   Caption         =   "cattle"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   10800
      TabIndex        =   3
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   12600
      TabIndex        =   2
      Top             =   9000
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6735
      Left            =   10560
      TabIndex        =   0
      Top             =   2040
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11880
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"cattle4.frx":0000
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
   Begin VB.Image Image1 
      Height          =   6180
      Left            =   360
      Picture         =   "cattle4.frx":008E
      Top             =   2160
      Width           =   9420
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
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "cattle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for previous command button
Private Sub cmd2_Click()
cattle3.Show
cattle4.Hide
Unload cattle4
End Sub

'code for rich text box
Private Sub Form_Load()
RichTextBox1.TextRTF = "Mehsana is a dairy breed of buffalo found in Mehsana, Sabarkanda and Banaskanta districts in Gujarat and adjoining Maharashtra state. The breed is evolved out of crossbreeding between the Surti and the Murrah. Body is longer than Murrah but limbs are lighter. The horns are less curved than in Murrah and are irregular. The milk yield is 1200-1500 kgs per lactation."
End Sub

'code for next text box
Private Sub cmd1_click()
cattle5.Show
cattle4.Hide
Unload cattle4
End Sub

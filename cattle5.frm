VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cattle5 
   BackColor       =   &H80000008&
   Caption         =   "cattle"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   5520
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"cattle5.frx":0000
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
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   3090
      Left            =   240
      Picture         =   "cattle5.frx":008B
      Top             =   1680
      Width           =   4560
   End
End
Attribute VB_Name = "cattle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for previous command button
Private Sub cmd2_Click()
cattle4.Show
cattle5.Hide
Unload cattle5
End Sub

'code for rich textbox
Private Sub Form_Load()
RichTextBox1.TextRTF = "This buffalo is named after an ancient tribe, Toda of Nilgiris Hills of south India and it is a semi-wild breed. The predominate coat colours are fawn and ash-grey. Thick hair coat is found all over the body. They are gregarious in nature. The body is long and deep and the chest is deep. The legs are short and strong. The horns are set wide apart curving inward, outward and forward forming a characteristic crescent shape. The average milk yield is 500 kgs per lactation with high fat content of 8%."
End Sub

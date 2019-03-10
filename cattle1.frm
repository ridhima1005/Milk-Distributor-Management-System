VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cattle1 
   BackColor       =   &H80000008&
   Caption         =   "cattle"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   2895
      Left            =   6600
      TabIndex        =   4
      Top             =   7320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"cattle1.frx":0000
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
   Begin VB.CommandButton cmd1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   11760
      TabIndex        =   1
      Top             =   9600
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3015
      Left            =   6600
      TabIndex        =   0
      Top             =   2040
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5318
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"cattle1.frx":008E
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
   Begin VB.Label lbl2 
      Caption         =   "Krishna Valley Cow"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   6120
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   3180
      Left            =   720
      Picture         =   "cattle1.frx":011F
      Top             =   7320
      Width           =   4515
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
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   720
      Picture         =   "cattle1.frx":2EE01
      Top             =   1560
      Width           =   4560
   End
End
Attribute VB_Name = "cattle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for next command button
Private Sub cmd1_click()
cattle2.Show
cattle1.Hide
Unload cattle1
End Sub

'code for rich textbox
Private Sub Form_Load()
RichTextBox1.TextRTF = "Originated from black cotton soil of the water shed of the river Krishna in Karnataka and also found in border districts of Maharashtra. Animals are large, having a massive frame with deep, loosely built short body. Tail almost reaches the ground. Generally grey white in color with a darker shade on fore quarters and hind quarters in male. Adults females are more whitish in appearance. The bullocks are powerful animals useful for slow ploughing, and valued for their good working qualities. The average yield is about 900 kgs per lactation."
RichTextBox2.TextRTF = "This breed is also called as Elitchpuri or Barari.  The breeding tract of this breed is Nagpur, Akola and Amrawati districts of Maharashtra. These are black coloured animal with white patches on face, legs and tail. The horns are long, flat and curved, bending backward on each side of the back. (Swaord shaped horns). The milk yield ranges from 700 to 1200 kgs per lactation."
End Sub


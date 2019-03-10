VERSION 5.00
Begin VB.MDIForm shakti 
   BackColor       =   &H8000000C&
   Caption         =   "MDMS"
   ClientHeight    =   8955
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   18315
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000008&
      Height          =   10215
      Left            =   0
      ScaleHeight     =   10155
      ScaleWidth      =   18255
      TabIndex        =   0
      Top             =   0
      Width           =   18315
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H00808000&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Sunshiney"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14520
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lbl1 
         BackColor       =   &H80000009&
         Caption         =   "SHAKTI MILK"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   9735
      End
      Begin VB.Image Image3 
         Height          =   960
         Left            =   15000
         Picture         =   "MDIForm1.frx":0000
         Top             =   5640
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "SHAKTI MILK, PUNE-MUMBAI HIGHWAY, PUNE-411010              020-26565835"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   13320
         TabIndex        =   2
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "CONTACT US:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   13320
         TabIndex        =   1
         Top             =   6120
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   6660
         Left            =   600
         Picture         =   "MDIForm1.frx":14C2
         Top             =   2880
         Width           =   9540
      End
   End
   Begin VB.Menu FILE 
      Caption         =   "&FILE"
      Begin VB.Menu Processing 
         Caption         =   "Milk Processing"
      End
      Begin VB.Menu Analysis 
         Caption         =   "Milk Analysis"
      End
      Begin VB.Menu Nutrition 
         Caption         =   "Milk Nutrition"
      End
      Begin VB.Menu cattle 
         Caption         =   "Cattle Details"
      End
   End
   Begin VB.Menu EXIT 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "shakti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for milk analysis in menu bar
Private Sub Analysis_Click()
manalysis.Show
End Sub

'code for cattle details
Private Sub cattle_Click()
cattle1.Show
End Sub

'code for login command button
Private Sub cmdlogin_Click()
Welcome.Show
End Sub

'code for exit in menu bar
Private Sub EXIT_Click()
End
End Sub

'code for milk nutrition in menu bar
Private Sub Nutrition_Click()
mnutrition.Show
End Sub

'code for milk processing in menu bar
Private Sub Processing_Click()
process.Show
End Sub

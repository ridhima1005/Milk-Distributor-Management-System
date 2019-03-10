VERSION 5.00
Begin VB.Form empd 
   BackColor       =   &H0080C0FF&
   Caption         =   "employee details"
   ClientHeight    =   8850
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13215
   FillColor       =   &H00C0E0FF&
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   6960
      TabIndex        =   32
      Top             =   3120
      Width           =   3495
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fill All The Fields"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image4 
         Height          =   960
         Left            =   600
         Picture         =   "emp.frx":0000
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   33
         Top             =   2640
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Success"
      Height          =   1815
      Left            =   4440
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   4
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Employee Added Successfully"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add new"
      Height          =   6735
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   6735
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4455
         Begin VB.Label Label10 
            BackColor       =   &H0080C0FF&
            Caption         =   "Alpabets And Expressions Not Allowed!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   3975
         End
         Begin VB.Label Label11 
            BackColor       =   &H0080C0FF&
            Caption         =   "  OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   30
            Top             =   2280
            Width           =   855
         End
         Begin VB.Image Image5 
            Height          =   960
            Left            =   480
            Picture         =   "emp.frx":14C2
            Top             =   1680
            Width           =   1500
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox txt8 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   5280
         Width           =   2295
      End
      Begin VB.TextBox txt7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   25
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox txt6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   24
         Top             =   4080
         Width           =   3975
      End
      Begin VB.TextBox txt5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2520
         TabIndex        =   23
         Top             =   3000
         Width           =   3975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txt4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   20
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txt3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   19
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txt2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txt1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Date Of Joining"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label lbl7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Contact Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label lbl8 
         BackColor       =   &H0080C0FF&
         Caption         =   "E-mail Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label lbl5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lbl6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lbl4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Date Of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbl2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Employee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Employee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   7560
      Picture         =   "emp.frx":2984
      Top             =   1920
      Width           =   6750
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label lb1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "EMPLOYEE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   6555
   End
End
Attribute VB_Name = "empd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txt1.Text = "" Or txt2.Text = "" Or txt3.Text = "" Or txt4.Text = "" Or txt5.Text = "" Or txt6.Text = "" Or txt7.Text = "" Or txt8.Text = "" Then
Frame4.Visible = True
ElseIf Not IsNumeric(txt6.Text) Then
Frame3.Visible = True
Else
Frame1.Visible = True
End Sub

Private Sub Command2_Click()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
txt4.Text = ""
txt5.Text = ""
txt6.Text = ""
txt7.Text = ""
txt8.Text = ""
Option1.Value = 0
Option2.Value = 0

End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub mnucustomer_Click()
cust.Show
End Sub

Private Sub Label11_Click()
Frame3.Visible = False
empd.Show
End Sub

Private Sub Label3_Click()
Menu.Show
empd.Hide
End Sub

Private Sub Label5_Click()
empd.Show
Frame1.Visible = False
End Sub

Private Sub Label6_Click()
empd.Hide
EmpDetails.Show
End Sub

Private Sub Label7_Click()
Frame2.Visible = True
End Sub

Private Sub Label9_Click()
Frame4.Visible = False
empd.Show
End Sub

Private Sub txt1_Change()
txt1.Text = txt1.Text
End Sub



VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form emp2 
   BackColor       =   &H0080C0FF&
   Caption         =   "emp"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   7335
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   6735
      Begin VB.CommandButton cmd3 
         BackColor       =   &H0080C0FF&
         Caption         =   "VIEW"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6000
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
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
         Left            =   4920
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
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
         Left            =   2640
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmd2 
         BackColor       =   &H0080C0FF&
         Caption         =   "HOME"
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H0080C0FF&
         Caption         =   "BACK"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6000
         Width           =   1455
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
         Locked          =   -1  'True
         TabIndex        =   7
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
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   5280
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
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3600
         Width           =   3975
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
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Top             =   3000
         Width           =   855
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
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   3
         Top             =   240
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   101515265
         CurrentDate     =   41554
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Gender"
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
         TabIndex        =   19
         Top             =   1800
         Width           =   1815
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
         TabIndex        =   18
         Top             =   5280
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
         Height          =   375
         Left            =   240
         TabIndex        =   17
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
         Left            =   480
         TabIndex        =   16
         Top             =   3000
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
         TabIndex        =   15
         Top             =   3840
         Width           =   1095
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
         Left            =   360
         TabIndex        =   14
         Top             =   360
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
         Left            =   480
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "DOB"
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
         Left            =   480
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   5415
      Index           =   0
      Left            =   8400
      Picture         =   "emp2.frx":0000
      Top             =   2400
      Width           =   6750
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
      Left            =   -120
      TabIndex        =   20
      Top             =   480
      Width           =   6555
   End
End
Attribute VB_Name = "emp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'when form loads
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=emp"

End Sub

'code for view command button
Private Sub cmd3_Click()
Dim view As Integer

If (MsgBox("Do you want to view...", vbYesNo) = vbYes) Then
On Error GoTo l1

view = InputBox("enter the employee id")
On Error GoTo l1

rs.Open "select e_id,e_name,e_gender,e_dob,e_age,e_address,e_email,e_contact from emp where e_id=' " & view & " '", con, adOpenDynamic, adLockOptimistic
On Error GoTo l1
txt2.Text = rs.Fields("e_id")
txt1.Text = rs.Fields("e_name")
If (rs.Fields("e_gender") = "male") Then
Option1.Value = True
ElseIf (rs.Fields("e_gender") = "female") Then
Option2.Value = True
End If
DTPicker1.Value = rs.Fields("e_dob")
txt4.Text = rs.Fields("e_age")
txt5.Text = rs.Fields("e_address")
txt7.Text = rs.Fields("e_email")
txt6.Text = rs.Fields("e_contact")
rs.Close
Exit Sub
l1:
MsgBox "Rec not found"
End If

End Sub


'code for back command button
Private Sub cmd1_click()
txt2.Text = ""
txt1.Text = ""
Option1.Value = False
Option2.Value = False
txt4.Text = ""
txt5.Text = ""
txt7.Text = ""
txt6.Text = ""
menue.Show
emp2.Hide
Unload emp2
End Sub

'code for menu command button
Private Sub cmd2_Click()
txt2.Text = ""
txt1.Text = ""
Option1.Value = False
Option2.Value = False
txt4.Text = ""
txt5.Text = ""
txt7.Text = ""
txt6.Text = ""
Menu.Show
emp2.Hide
Unload emp2
End Sub

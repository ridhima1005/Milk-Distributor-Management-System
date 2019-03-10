VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "EMPLOYEE REPORT"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   7335
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   6735
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   141426689
         CurrentDate     =   41554
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
         TabIndex        =   18
         Top             =   960
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
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   17
         Top             =   240
         Width           =   3855
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
         MaxLength       =   2
         TabIndex        =   16
         Top             =   3000
         Width           =   855
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
         TabIndex        =   15
         Top             =   3600
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   14
         Top             =   5400
         Width           =   3975
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
         Left            =   2400
         TabIndex        =   13
         Top             =   4800
         Width           =   3975
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H0080C0FF&
         Caption         =   "SAVE"
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
         TabIndex        =   12
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmd2 
         BackColor       =   &H0080C0FF&
         Caption         =   "CLEAR"
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
         TabIndex        =   11
         Top             =   6000
         Width           =   1575
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
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   27
         Top             =   2520
         Width           =   1335
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
         TabIndex        =   25
         Top             =   960
         Width           =   1215
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
         TabIndex        =   24
         Top             =   360
         Width           =   1575
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
         TabIndex        =   23
         Top             =   3840
         Width           =   1095
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
         TabIndex        =   22
         Top             =   3000
         Width           =   975
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
         TabIndex        =   21
         Top             =   4800
         Width           =   1695
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
         Left            =   240
         TabIndex        =   20
         Top             =   5400
         Width           =   2055
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9360
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=emp"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=emp"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd3 
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   1455
   End
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
      Begin VB.CommandButton cmd4 
         BackColor       =   &H0080C0FF&
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
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   720
         Picture         =   "emplo.frx":0000
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Label lbl11 
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
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Success"
      Height          =   2055
      Left            =   2040
      TabIndex        =   1
      Top             =   3360
      Width           =   4455
      Begin VB.CommandButton cmd6 
         BackColor       =   &H0080C0FF&
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   960
         Left            =   600
         Picture         =   "emplo.frx":14C2
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lbl12 
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
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Image Image1 
      Height          =   5415
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   6750
   End
   Begin VB.Image Image1 
      Height          =   5415
      Index           =   0
      Left            =   6480
      Picture         =   "emplo.frx":2984
      Top             =   1680
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

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'code when the form opens
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

Adodc1.Visible = False

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=emp"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
con.Close

Exit Sub
End Sub

'code for data report
Private Sub Command1_Click()
DataEnvironment2.Command1
DataReport2.Show
End Sub

'code of add employee command button
Private Sub cmd1_click()
Dim id As Integer
Dim gen As String

If (Option1.Value = True) Then
gen = "male"
ElseIf (Option2.Value = True) Then
gen = "female"
End If

If txt1.Text = " " Or txt4.Text = "  " Or txt5.Text = " " Or txt6.Text = " " Or txt7.Text = " " Then
MsgBox "fields empty"

Else

con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'auto increment of id
cmd.CommandText = "select max(e_id) from emp"
On Error GoTo l1
Set rs = cmd.Execute
id = rs.Fields(0)
txt2.Text = id + 1

'insert into table
cmd.CommandText = "insert into emp values('" & txt2.Text & "','" & txt1.Text & "','" & gen & "','" & DTPicker1 & "','" & txt4.Text & "','" & txt5.Text & "','" & txt7.Text & "','" & txt6.Text & "')"

cmd.Execute
MsgBox "Inserted"
con.Close
End If

Exit Sub
l1:
End Sub

'code of clear command button
Private Sub cmd2_Click()
txt1.Text = " "
txt2.Text = " "
txt4.Text = " "
txt5.Text = " "
txt6.Text = " "
txt7.Text = " "
Option1.Value = 0
Option2.Value = 0
End Sub

'code of home command button
Private Sub cmd3_Click()
menue.Show
empd.Hide
Unload empd
End Sub

'code of valdiation for name
Private Sub txt1_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 97 And KeyASCII <= 122) Or KeyASCII = 127 Or (KeyASCII >= 65 And KeyASCII <= 90) Or KeyASCII = 32 Or KeyASCII = 46 Or KeyASCII = 8) Then
MsgBox "numbers or characters not allowed!"
KeyASCII = 0
txt1.Text = ""
End If
End Sub

'code for age
Private Sub txt4_Click()
Dim age As Integer
Dim dob As String

'age calculation
dob = DTPicker1.Value
age = DateDiff("yyyy", dob, Now)
txt4.Text = age

'age limit
If Val(txt4.Text) < 18 Then
MsgBox " Enter valid dob!"
txt4.Text = " "
Else
 If Val(txt4.Text) > 60 Then
 MsgBox "enter a valid dob!"
 txt4.Text = " "
 Else
 MsgBox "Valid dob!"
 End If
End If

End Sub

'code of valdiation for address
Private Sub txt5_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 97 And KeyASCII <= 122) Or KeyASCII = 127 Or (KeyASCII >= 65 And KeyASCII <= 90) Or KeyASCII = 32 Or KeyASCII = 46 Or KeyASCII = 8 Or KeyASCII = 44 Or KeyASCII = 127 Or (KeyASCII >= 48 And KeyASCII <= 57)) Then
MsgBox "Enter valid address!"
KeyASCII = 0
txt5.Text = ""
End If
End Sub

'code of valdiation for contact
Private Sub txt6_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
MsgBox "Only numbers are allowed!"
txt6.Text = ""
End If
End Sub

'code of valdiation for contact
Private Sub txt6_lostfocus()
If Len(txt6.Text) <> 10 Then
MsgBox "Contact no should be 10 digits"
txt6.Text = ""
End If
End Sub

'code for email validation
Private Sub txt7_lostfocus()
Dim str1 As String
str1 = txt7.Text
If ((InStr(str1, "@")) And ((InStr(str1, "gmail.com")) Or (InStr(str1, "hotmail.com")) Or (InStr(str1, "google.com")) Or (InStr(str1, "yahoo.com")) Or (InStr(str1, "yahoo.in")))) Then
MsgBox "valid email id"
Else
MsgBox "not valid email id...plz enter vaild email id!"
txt7.Text = ""
End If
End Sub

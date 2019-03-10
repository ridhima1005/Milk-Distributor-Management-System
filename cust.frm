VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form customer 
   BackColor       =   &H00C0FFC0&
   Caption         =   " "
   ClientHeight    =   9420
   ClientLeft      =   300
   ClientTop       =   990
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   15960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CUSTOMER REPORT"
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
      Left            =   6840
      TabIndex        =   29
      Top             =   8640
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11400
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=cus"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=cus"
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Address"
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
      Left            =   4920
      TabIndex        =   26
      Top             =   3000
      Width           =   4455
      Begin VB.CommandButton cmd9 
         BackColor       =   &H00C0FFC0&
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Expressions Not Allowed!"
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
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   3975
      End
      Begin VB.Image Image10 
         Height          =   960
         Left            =   600
         Picture         =   "cust.frx":0000
         Top             =   1560
         Width           =   1500
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name"
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
      Left            =   4920
      TabIndex        =   20
      Top             =   3240
      Width           =   4455
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00C0FFC0&
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Image Image6 
         Height          =   960
         Left            =   480
         Picture         =   "cust.frx":14C2
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lbl10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Numbers And Expressions Not Allowed!"
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
         TabIndex        =   22
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Success"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   4680
      TabIndex        =   18
      Top             =   2640
      Width           =   4455
      Begin VB.CommandButton cmd8 
         BackColor       =   &H00C0FFC0&
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
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Added Successfully"
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
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image Image3 
         Height          =   2400
         Left            =   360
         Picture         =   "cust.frx":2984
         Top             =   720
         Width           =   3750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
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
      Left            =   5280
      TabIndex        =   16
      Top             =   2640
      Width           =   3495
      Begin VB.CommandButton cmd7 
         BackColor       =   &H00C0FFC0&
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   17
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image4 
         Height          =   960
         Left            =   600
         Picture         =   "cust.frx":847A
         Top             =   1320
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
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
      Left            =   4800
      TabIndex        =   14
      Top             =   2520
      Width           =   4455
      Begin VB.CommandButton cmd6 
         BackColor       =   &H00C0FFC0&
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   15
         Top             =   360
         Width           =   3975
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   480
         Picture         =   "cust.frx":993C
         Top             =   1680
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00C0FFC0&
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
      Left            =   9120
      TabIndex        =   13
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0FFC0&
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
      Left            =   4560
      TabIndex        =   12
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   2280
      TabIndex        =   11
      Top             =   8640
      Width           =   1815
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
      Left            =   5520
      TabIndex        =   10
      Top             =   5520
      Width           =   2895
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
      Height          =   855
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   9
      Top             =   6600
      Width           =   2895
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
      Height          =   735
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4440
      Width           =   2895
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
      Height          =   855
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
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
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   5070
      Left            =   10200
      Picture         =   "cust.frx":ADFE
      Top             =   2400
      Width           =   4500
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "E-mail Id:"
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
      Left            =   840
      TabIndex        =   5
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Contact Number:"
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
      Left            =   840
      TabIndex        =   4
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Address:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Id No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name:"
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
      Left            =   840
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   9855
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

Private Sub Command1_Click()
DataEnvironment3.Command1
DataReport3.Show
End Sub

'code when form loads
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

'for making all the frames invisible when the form opens
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False

Adodc1.Visible = False

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=cus"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
con.Close

Exit Sub
End Sub

'code for save command button
Private Sub cmd1_click()
Dim id As Integer

'if any field is empty
If txt1.Text = "" Or txt3.Text = "" Or txt4.Text = "" Or txt5.Text = "" Then
Frame2.Visible = True

Else

con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'autoincrement of id
cmd.CommandText = "select max(c_id) from cus"
On Error GoTo l1
Set rs = cmd.Execute
id = rs.Fields(0)
txt2.Text = id + 1

'insert values into table
cmd.CommandText = "insert into cus values('" & txt2.Text & "','" & txt1.Text & "','" & txt3.Text & "','" & txt5.Text & "','" & txt4.Text & "')"

cmd.Execute
MsgBox "Inserted"
con.Close
End If

Exit Sub
l1:
End Sub

'code for clear command button
Private Sub cmd3_Click()
txt1.Text = " "
txt2.Text = " "
txt3.Text = " "
txt4.Text = " "
txt5.Text = " "
End Sub

'code for home command button
Private Sub cmd4_Click()
menuc.Show
customer.Hide
Unload customer
End Sub

'code of name frame's command button
Private Sub cmd5_Click()
Frame4.Visible = False
customer.Show
End Sub

'code of address frame's command button
Private Sub cmd9_Click()
Frame5.Visible = False
customer.Show
End Sub

'code of contact frame's command button
Private Sub cmd6_Click()
Frame3.Visible = False
customer.Show
End Sub

'code of success frame's command button
Private Sub cmd8_Click()
Frame1.Visible = False
customer.Show
End Sub

'code of empty frame's command button
Private Sub cmd7_Click()
Frame2.Visible = False
customer.Show
End Sub

'code of valdiation for name
Private Sub txt1_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 97 And KeyASCII <= 122) Or KeyASCII = 127 Or (KeyASCII >= 65 And KeyASCII <= 90) Or KeyASCII = 32 Or KeyASCII = 46 Or KeyASCII = 8) Then
Frame4.Visible = True
KeyASCII = 0
txt1.Text = ""
End If
End Sub

'code of valdiation for address
Private Sub txt3_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 97 And KeyASCII <= 122) Or KeyASCII = 127 Or (KeyASCII >= 65 And KeyASCII <= 90) Or KeyASCII = 32 Or KeyASCII = 46 Or KeyASCII = 8 Or KeyASCII = 44 Or KeyASCII = 127 Or (KeyASCII >= 48 And KeyASCII <= 57)) Then
Frame5.Visible = True
KeyASCII = 0
txt3.Text = ""
End If
End Sub

'code of valdiation for contact number
Private Sub txt4_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
Frame3.Visible = True
txt4.Text = ""
End If
End Sub

'code of valdiation for contact number length
Private Sub txt4_lostfocus()
If Len(txt4.Text) <> 10 Then
MsgBox "Contact no should be 10 digits"
txt4.Text = ""
End If
End Sub

'code for validation of email
Private Sub txt5_lostfocus()
Dim str1 As String

str1 = txt5.Text
If ((InStr(str1, "@")) And ((InStr(str1, "gmail.com")) Or (InStr(str1, "hotmail.com")) Or (InStr(str1, "google.com")) Or (InStr(str1, "yahoo.com")) Or (InStr(str1, "yahoo.in")))) Then
MsgBox "Valid email id"
Else
MsgBox "Not valid email id.. Plz enter a valid email id"
txt5.Text = ""
End If
End Sub

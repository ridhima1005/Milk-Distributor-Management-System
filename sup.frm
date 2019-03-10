VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form supplierd 
   BackColor       =   &H00C0E0FF&
   Caption         =   "supplier details"
   ClientHeight    =   8865
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12315
   FillColor       =   &H00C0E0FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0E0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SUPPLIER REPORT"
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
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "sup.frx":0000
      Left            =   5400
      List            =   "sup.frx":000A
      TabIndex        =   34
      Text            =   "Quantity"
      Top             =   6240
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8280
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=sup"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=sup"
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
   Begin VB.Frame Frame8 
      BackColor       =   &H0080C0FF&
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
      Left            =   3600
      TabIndex        =   31
      Top             =   1800
      Width           =   4455
      Begin VB.CommandButton cmd11 
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Image Image10 
         Height          =   960
         Left            =   480
         Picture         =   "sup.frx":0015
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lbl16 
         BackColor       =   &H0080C0FF&
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Daily Supply"
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
      Left            =   3600
      TabIndex        =   28
      Top             =   1800
      Width           =   4455
      Begin VB.CommandButton cmd10 
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl15 
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
         TabIndex        =   30
         Top             =   360
         Width           =   3975
      End
      Begin VB.Image Image9 
         Height          =   960
         Left            =   480
         Picture         =   "sup.frx":14D7
         Top             =   1680
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00C0E0FF&
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
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
      Left            =   3480
      TabIndex        =   21
      Top             =   1800
      Width           =   4455
      Begin VB.CommandButton cmd3 
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl10 
         BackColor       =   &H0080C0FF&
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
      Begin VB.Image Image6 
         Height          =   960
         Left            =   480
         Picture         =   "sup.frx":2999
         Top             =   1680
         Width           =   1500
      End
   End
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
      Left            =   3360
      TabIndex        =   19
      Top             =   1680
      Width           =   4455
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
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   480
         Picture         =   "sup.frx":3E5B
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lbl11 
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
         TabIndex        =   20
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   3360
      TabIndex        =   17
      Top             =   1440
      Width           =   3495
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   960
         Left            =   600
         Picture         =   "sup.frx":531D
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lbl12 
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
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
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
      Left            =   3600
      TabIndex        =   15
      Top             =   1080
      Width           =   4455
      Begin VB.CommandButton cmd5 
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   2400
         Left            =   360
         Picture         =   "sup.frx":67DF
         Top             =   720
         Width           =   3750
      End
      Begin VB.Label lbl9 
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
         TabIndex        =   16
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ADD"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txt7 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txt5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      TabIndex        =   10
      Top             =   4920
      Width           =   4095
   End
   Begin VB.TextBox txt4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   9
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3600
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   5565
      Left            =   8160
      Picture         =   "sup.frx":C2D5
      Top             =   1920
      Width           =   4995
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   1200
      Picture         =   "sup.frx":28ACE
      Top             =   7440
      Width           =   5850
   End
   Begin VB.Label lbl8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Suppy"
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
      Left            =   600
      TabIndex        =   11
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Id"
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
      Left            =   600
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Id No"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "supplierd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'code when form loads
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

'making all the frames invisible when the form opens
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame7.Visible = False
Frame8.Visible = False

Adodc1.Visible = False

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=sup"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
con.Close

Exit Sub
End Sub

'code for datareport
Private Sub Command1_Click()
DataEnvironment4.Command1
DataReport4.Show
End Sub

'code of add supplier command button
Private Sub cmd1_click()
Dim id As Integer

'empty fields
If txt1.Text = " " Or txt3.Text = " " Or txt4.Text = " " Or txt5.Text = " " Or txt7.Text = " " Then
'Frame2.Visible = True
MsgBox "fields are empty"
Else

con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'autoincrement of id
cmd.CommandText = "select max(s_id) from sup"
On Error GoTo l1
Set rs = cmd.Execute
id = rs.Fields(0)
txt2.Text = id + 1

'insert in table
cmd.CommandText = "insert into sup values('" & txt2.Text & "','" & txt1.Text & "','" & txt3.Text & "','" & txt4.Text & "','" & txt5.Text & "','" & txt7.Text & "','" & Combo1.Text & "')"
cmd.Execute
Frame1.Visible = True
con.Close
End If

Exit Sub
l1:
End Sub

'code of daily supply frame's command button
Private Sub cmd10_Click()
Frame7.Visible = False
supplierd.Show
End Sub

'code of address frame's command button
Private Sub cmd11_Click()
Frame8.Visible = False
supplierd.Show
End Sub

'code of clear command button
Private Sub cmd2_Click()
txt1.Text = " "
txt2.Text = " "
txt3.Text = " "
txt4.Text = " "
txt5.Text = " "
txt7.Text = " "
End Sub

'code of name frame's command button
Private Sub cmd3_Click()
Frame4.Visible = False
supplierd.Show
End Sub

'code of contact frame's command button
Private Sub cmd4_Click()
Frame3.Visible = False
supplierd.Show
End Sub

'code of success frame's command button
Private Sub cmd5_Click()
Frame1.Visible = False
supplierd.Show
End Sub

'code of empty frame's command button
Private Sub cmd6_Click()
Frame2.Visible = False
supplierd.Show
End Sub

'code for home command button
Private Sub cmd7_Click()
menus.Show
supplierd.Hide
Unload supplierd
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
Frame8.Visible = True
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

'code for email validation
Private Sub txt5_lostfocus()
Dim str1 As String
str1 = txt5.Text
If (InStr(str1, "@")) Then
MsgBox "valid email id"
Else
MsgBox "not valid email id..Plz enter valid email id!"
txt5.Text = ""
End If
End Sub

'code of valdiation for daily supply
Private Sub txt7_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
Frame7.Visible = True
txt7.Text = ""
End If
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form milkp 
   BackColor       =   &H00C0FFFF&
   Caption         =   "item"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9705
   FillColor       =   &H00C0FFFF&
   ForeColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MILK DETAILS 2"
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
      TabIndex        =   23
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "VIEW RECORDS"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5520
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102170625
      CurrentDate     =   41554
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102170625
      CurrentDate     =   41554
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "item.frx":0000
      Left            =   3360
      List            =   "item.frx":000A
      TabIndex        =   19
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DELETE"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk3"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk3"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Success"
      Height          =   3375
      Left            =   2160
      TabIndex        =   14
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lbl9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Successful"
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3855
      End
      Begin VB.Image Image2 
         Height          =   2010
         Left            =   360
         Picture         =   "item.frx":001D
         Top             =   840
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmd4 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lbl8 
         BackColor       =   &H00C0FFFF&
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
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image4 
         Height          =   960
         Left            =   600
         Picture         =   "item.frx":4D8F
         Top             =   1320
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txt5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
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
      Left            =   3360
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txt1 
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lbl11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "animal"
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
      TabIndex        =   17
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   6360
      Picture         =   "item.frx":6251
      Top             =   1320
      Width           =   4515
   End
   Begin VB.Label lbl7 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Left            =   2880
      TabIndex        =   8
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   600
      TabIndex        =   6
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date  (mm/dd/yyyy)"
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Date (mm/dd/yyyy)"
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
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Milk ID"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Milk Details 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "milkp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'code for making frames invisible when form opens
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

Frame1.Visible = False
Frame2.Visible = False

Adodc1.Visible = False

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk3"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
con.Close

Exit Sub
End Sub

'code for previous form
Private Sub cmd8_Click()
milka.Show
milkp.Hide
Unload milkp
End Sub

'code for view button
Private Sub cmd7_Click()
dmilk.Show
End Sub
'code for delete button
Private Sub cmd6_Click()
Dim view As Integer

If (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
view = InputBox("enter the milk id")
Set rs = New ADODB.Recordset
On Error GoTo l1

rs.Open "select * from milk3 where m_id = '" & view & "'", con, adOpenKeyset, adLockPessimistic

rs.Delete
con.Execute "commit"
rs.Close
Set rs = Nothing
MsgBox "Deleted Succesfully..."
Exit Sub
l1:
MsgBox "Rec not found"
End If

End Sub

'code for save command button
Private Sub cmd1_click()
Dim id As Integer

If txt5.Text = "" Or Combo1.Text = "" Then
MsgBox "Fields empty!"

Else

con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'autoincrement of id
cmd.CommandText = "select max(m_id) from milk3"
On Error GoTo l1
Set rs = cmd.Execute
id = rs.Fields(0)
txt1.Text = id + 1

cmd.CommandText = "insert into milk3 values('" & txt1.Text & "','" & Combo1.Text & "','" & DTPicker1.Value & "','" & DTPicker2.Value & "','" & txt5.Text & "')"
cmd.Execute
Frame1.Visible = True
con.Close
End If

Exit Sub
l1:
End Sub

'code for clear button
Private Sub cmd2_Click()
txt1.Text = " "
txt5.Text = " "
End Sub

'code of success frame's command button
Private Sub cmd5_Click()
milkp.Show
Frame1.Visible = False
End Sub

'code for home command button
Private Sub cmd3_Click()
Menu.Show
milkp.Hide
Unload milkp
End Sub


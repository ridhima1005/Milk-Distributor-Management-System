VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form milkd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "milk details"
   ClientHeight    =   7125
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10935
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MILK DETAILS1"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "milkd.frx":0000
      Left            =   2640
      List            =   "milkd.frx":000A
      TabIndex        =   11
      Text            =   "Animal"
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "milkd.frx":001C
      Left            =   2640
      List            =   "milkd.frx":0026
      TabIndex        =   10
      Text            =   "Packings"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "milkd.frx":0037
      Left            =   2640
      List            =   "milkd.frx":0044
      TabIndex        =   9
      Text            =   "Packs"
      Top             =   2280
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7560
      Top             =   6480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk2"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk2"
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
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAVE AND  CONTINUE"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
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
      Height          =   420
      ItemData        =   "milkd.frx":005E
      Left            =   2640
      List            =   "milkd.frx":009E
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sale Record No:"
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
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6180
      Left            =   5880
      Picture         =   "milkd.frx":00E9
      Top             =   240
      Width           =   4860
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number Of Packs"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Animal"
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
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Packings"
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
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Packs"
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
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Milk Details 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "milkd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'code for making all the forms invisible when the form opens
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

Adodc1.Visible = False

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk2"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

Exit Sub
End Sub

'code for prevoius form
Private Sub cmd6_Click()
milka.Show
milkd.Hide
Unload milkd
End Sub

'code for view details command button
Private Sub cmd4_Click()
milkdd.Show
End Sub

'code for delete command button
Private Sub cmd5_Click()
Dim view As Integer

If (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
view = InputBox("enter the sale record number")
Set rs = New ADODB.Recordset
On Error GoTo l1

rs.Open "select * from milk2 where s_r_no = '" & view & "'", con, adOpenKeyset, adLockPessimistic

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

'code of continue frame's command button
Private Sub cmd1_click()
Dim id As Integer

If Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
MsgBox "Fields are empty!"

Else

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'autoincrement of id
cmd.CommandText = "select max(s_r_no) from milk2"
On Error GoTo l1
Set rs = cmd.Execute
id = rs.Fields(0)
txt1.Text = id + 1

'insert values
cmd.CommandText = "insert into milk2 values('" & txt1.Text & "','" & Combo4.Text & "','" & Combo2.Text & "','" & Combo3.Text & "','" & Combo1.Text & "')"
cmd.Execute

milkp.Show
milkd.Hide
Unload milkd
con.Close
End If
Exit Sub
l1:

End Sub

'code for clear command button
Private Sub cmd2_Click()
txt1.Text = ""
End Sub

'code for home command button
Private Sub cmd3_Click()
Menu.Show
milkd.Hide
Unload milkd
End Sub


VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form milka 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   675
      Left            =   4440
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txt7 
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmd10 
      BackColor       =   &H00FFC0FF&
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
      Left            =   2280
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
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
      ItemData        =   "milka.frx":0000
      Left            =   4560
      List            =   "milka.frx":000A
      TabIndex        =   9
      Text            =   "Quantity"
      Top             =   2280
      Width           =   1695
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
      ItemData        =   "milka.frx":0015
      Left            =   4560
      List            =   "milka.frx":001F
      TabIndex        =   8
      Text            =   "Quantity"
      Top             =   1560
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5040
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk1"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk1"
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
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SAVE AND CONTINUE"
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
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00FFC0FF&
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
      Left            =   1320
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFC0FF&
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
      Left            =   3480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1335
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
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
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
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Record No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Milk For Sale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Milk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Milk Details 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   5460
      Left            =   0
      Picture         =   "milka.frx":002A
      Top             =   0
      Width           =   7110
   End
End
Attribute VB_Name = "milka"
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

Dim id As Integer

Adodc1.Visible = False

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=milk1"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
con.Close

Exit Sub
End Sub

'code for view command button
Private Sub Command1_Click()
milkad.Show
End Sub

'code for delete button
Private Sub cmd10_Click()
Dim view As Integer

If (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
On Error GoTo l1
view = InputBox("enter the record number")
Set rs = New ADODB.Recordset
On Error GoTo l1

rs.Open "select * from milk1 where r_no = '" & view & "'", con, adOpenKeyset, adLockPessimistic

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

'code of home command button
Private Sub cmd1_click()
Menu.Show
milka.Hide
Unload milka
End Sub

'code for clear command button
Private Sub cmd3_Click()
txt1.Text = ""
txt2.Text = ""
End Sub

'code of empty frame's command button
Private Sub cmd4_Click()
Frame2.Visible = False
milka.Show
End Sub

'code of milk frame's command button
Private Sub cmd5_Click()
Frame3.Visible = False
milka.Show
End Sub

'code of continue command button
Private Sub cmd7_Click()
If txt1.Text = "" Or txt2.Text = "" Then
MsgBox "Fields empty"

Else
'checking the values
If ((Val(txt2.Text) And Combo2.Text = "l") > (Val(txt1.Text) And Combo1.Text = "l")) Then
MsgBox "enter appropriate values"
Else
    If ((Val(txt2.Text) And Combo2.Text = "ml") > (Val(txt1.Text) And Combo1.Text = "ml")) Then
    MsgBox "enter appropriate values"
    Else
        If (Combo1.Text = "ml" And Combo2.Text = "l") Then
        MsgBox "enter appropriate values"
        Else
        MsgBox "correct values are entered!"
        'End If
    'End If
'End If
con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'autoincrement of id
cmd.CommandText = "select max(r_no) from milk1"
On Error GoTo l1
Set rs = cmd.Execute
id = rs.Fields(0)
txt7.Text = id + 1

'insert values
cmd.CommandText = "insert into milk1 values('" & txt7.Text & "','" & txt1.Text & "','" & Combo1.Text & "','" & txt2.Text & "','" & Combo2.Text & "')"
cmd.Execute
MsgBox "Success!"
con.Close
milkd.Show
milka.Hide
Unload milka
End If
End If
End If
End If

Exit Sub
l1:


End Sub

'validation for available milk
Private Sub txt1_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
MsgBox " Numbers only allowed!"
txt1.Text = ""
End If
End Sub

'validation for milk for sale
Private Sub txt2_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
MsgBox " Numbers only allowed!"
txt2.Text = ""
End If
End Sub

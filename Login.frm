VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Login"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6240
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=login"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=login"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
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
      Left            =   4200
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CLEAR"
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
      Left            =   2400
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "LOGIN"
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
      Left            =   720
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txt2 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6660
      Left            =   120
      Picture         =   "Login.frx":0000
      Top             =   120
      Width           =   9540
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'form loads
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=login"

Adodc1.Visible = False
End Sub

'code for login and password
Private Sub cmd1_click()
rs.Open "insert into login values('" & txt1.Text & "','" & txt2.Text & "')", con, adOpenDynamic, adLockOptimistic

If (rs.State = 1) Then
rs.Close
End If

If txt1.Text = "shakti" And txt2.Text = "password" Then
 Menu.Show
 Login.Hide
 Unload Login
Else
Invalid.Show
Login.Hide
Unload Login
End If

End Sub

'code for clear command button
Private Sub cmd2_Click()
txt1.Text = " "
txt2.Text = " "
End Sub

'code for exit button
Private Sub cmd3_Click()
End
End Sub



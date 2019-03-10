VERSION 5.00
Begin VB.Form sup2 
   BackColor       =   &H00C0E0FF&
   Caption         =   "supplier"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16035
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   16035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8280
      Width           =   1215
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
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8280
      Width           =   1215
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   4695
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
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   12
      Top             =   1680
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   3855
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
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4560
      Width           =   2775
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5280
      Width           =   4095
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
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
      ItemData        =   "sup2.frx":0000
      Left            =   7080
      List            =   "sup2.frx":000A
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Quantity"
      Top             =   6600
      Width           =   1575
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
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
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
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
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
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
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
      Left            =   480
      TabIndex        =   3
      Top             =   4560
      Width           =   2415
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
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
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
      Left            =   480
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   5565
      Left            =   10560
      Picture         =   "sup2.frx":0015
      Top             =   1560
      Width           =   4995
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
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
End
Attribute VB_Name = "sup2"
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

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=sup"

End Sub

'code for view command button
Private Sub cmd1_click()
Dim view As Integer

If (MsgBox("Do you want to view...", vbYesNo) = vbYes) Then
On Error GoTo l1
view = InputBox("enter the supplier id")
On Error GoTo l1

rs.Open "select s_id,s_name,s_address,s_contact,s_email,s_dsupply,s_quantity from sup where s_id=' " & view & " '", con, adOpenDynamic, adLockOptimistic
txt2.Text = rs.Fields("s_id")
txt1.Text = rs.Fields("s_name")
txt3.Text = rs.Fields("s_address")
txt4.Text = rs.Fields("s_contact")
txt5.Text = rs.Fields("s_email")
txt7.Text = rs.Fields("s_dsupply")
Combo1.Text = rs.Fields("s_quantity")
rs.Close
Exit Sub
l1:
MsgBox "Rec not found"
End If

End Sub

'code for back command button
Private Sub cmd2_Click()
menus.Show
sup2.Hide
Unload sup2
End Sub

'code for menu command button
Private Sub cmd7_Click()
sup2.Hide
Menu.Show
Unload Menu
End Sub

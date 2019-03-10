VERSION 5.00
Begin VB.Form deposite2 
   BackColor       =   &H00808080&
   Caption         =   "depositeview"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      Begin VB.CommandButton cmd1 
         BackColor       =   &H00808080&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   0
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "deposite2.frx":0000
         Left            =   3480
         List            =   "deposite2.frx":000A
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Quantity"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "deposite2.frx":0015
         Left            =   2040
         List            =   "deposite2.frx":001F
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Animal"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00808080&
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3720
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00808080&
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text6 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text5 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Id no"
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
         Left            =   480
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Depositing value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Percentage Of Fat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Animal"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   3810
      Left            =   7320
      Picture         =   "deposite2.frx":0031
      Top             =   1440
      Width           =   5715
   End
End
Attribute VB_Name = "deposite2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'when the form loads
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=dep"

End Sub

'code for view command button
Private Sub cmd1_click()
Dim view As Integer

If (MsgBox("Do you want to view...", vbYesNo) = vbYes) Then
On Error GoTo l1

view = InputBox("enter the depositor id")
On Error GoTo l1

rs.Open "select d_id,d_name,d_animal,d_fat,d_value,d_quantity from dep where d_id=' " & view & " '", con, adOpenDynamic, adLockOptimistic
Text8.Text = rs.Fields("d_id")
Text5.Text = rs.Fields("d_name")
Combo1.Text = rs.Fields("d_animal")
Text6.Text = rs.Fields("d_fat")
Text7.Text = rs.Fields("d_value")
Combo2.Text = rs.Fields("d_quantity")
rs.Close

Exit Sub
l1:
MsgBox "Rec not found"
End If

End Sub

'code for back command button
Private Sub Command5_Click()
menud.Show
deposite2.Hide
Unload deposite2
End Sub

'code for menu command button
Private Sub Command6_Click()
Menu.Show
deposite2.Hide
Unload deposite2
End Sub

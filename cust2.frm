VERSION 5.00
Begin VB.Form customer2 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Customer Details"
   ClientHeight    =   9735
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   16050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3000
      TabIndex        =   13
      Top             =   8640
      Width           =   1815
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
      Left            =   7560
      TabIndex        =   12
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0FFC0&
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
      Left            =   5280
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      TabIndex        =   7
      Top             =   3240
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
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   5070
      Left            =   10320
      Picture         =   "cust2.frx":0000
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
      Left            =   720
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
      Left            =   720
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
      Top             =   3360
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
      Top             =   2280
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
Attribute VB_Name = "customer2"
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

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=cus"

End Sub

'code to view the details
Private Sub cmd1_click()
Dim view As Integer

If (MsgBox("Do you want to view...", vbYesNo) = vbYes) Then
On Error GoTo l1
view = InputBox("enter the customer id")
On Error GoTo l1

rs.Open "select c_id,c_name,c_address,c_email,c_contact from cus where c_id=" & view, con, adOpenDynamic, adLockOptimistic
txt2.Text = rs.Fields("c_id")
txt1.Text = rs.Fields("c_name")
txt3.Text = rs.Fields("c_address")
txt5.Text = rs.Fields("c_email")
txt4.Text = rs.Fields("c_contact")
rs.Close
Exit Sub
l1:
MsgBox "Rec not found"
End If

End Sub

'code for back button
Private Sub cmd3_Click()
menuc.Show
customer2.Hide
Unload customer2
End Sub

'code for home command button
Private Sub cmd4_Click()
Menu.Show
customer.Hide
Unload customer
End Sub




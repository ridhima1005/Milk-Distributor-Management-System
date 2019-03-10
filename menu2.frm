VERSION 5.00
Begin VB.Form menuc 
   BackColor       =   &H80000008&
   Caption         =   "menuc"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10185
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "HOME"
      Height          =   855
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "VIEW A RECORD"
      Height          =   855
      Left            =   7080
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "RECORD DETAILS"
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Caption         =   "ADD NEW  ENTERY"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   3600
      Picture         =   "menu2.frx":0000
      Top             =   1680
      Width           =   2625
   End
End
Attribute VB_Name = "menuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'code for making all the frames invisible when the form opens
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=cus"

End Sub

'code for record details command button
Private Sub Command3_Click()
cdetails.Show
End Sub

'code for view command button
Private Sub Command4_Click()
customer2.Show
customer.Hide
Unload customer
End Sub

'code for add new entry command button
Private Sub Command1_Click()
customer.Show
menuc.Hide
Unload menuc
End Sub

'code for delete command button
Private Sub Command2_Click()
Dim view As Integer

If (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
On Error GoTo l1
view = InputBox("enter the customer id")
Set rs = New ADODB.Recordset
On Error GoTo l1

rs.Open "select * from cus where c_id = '" & view & "'", con, adOpenKeyset, adLockPessimistic

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

'code for home command button
Private Sub Command5_Click()
Menu.Show
menuc.Hide
Unload menuc
End Sub

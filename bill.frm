VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Bill 
   BackColor       =   &H00FFFFFF&
   Caption         =   "bill"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "bill.frx":0000
      Left            =   12480
      List            =   "bill.frx":0028
      TabIndex        =   41
      Text            =   "Products"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
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
      Left            =   12120
      TabIndex        =   40
      Top             =   6720
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9360
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=bill"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=bill"
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
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID No"
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
      Left            =   4440
      TabIndex        =   37
      Top             =   2400
      Width           =   4455
      Begin VB.CommandButton cmd12 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   38
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Image Image10 
         Height          =   960
         Left            =   480
         Picture         =   "bill.frx":0152
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   39
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.TextBox txt11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      MaxLength       =   5
      TabIndex        =   36
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contents quantity"
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
      Left            =   4440
      TabIndex        =   32
      Top             =   2400
      Width           =   4455
      Begin VB.CommandButton cmd11 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   33
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   34
         Top             =   480
         Width           =   3975
      End
      Begin VB.Image Image9 
         Height          =   960
         Left            =   480
         Picture         =   "bill.frx":1614
         Top             =   1680
         Width           =   1500
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5400
      TabIndex        =   29
      Top             =   2040
      Width           =   3495
      Begin VB.CommandButton cmd9 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   30
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lbl12 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   31
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image7 
         Height          =   960
         Left            =   600
         Picture         =   "bill.frx":2AD6
         Top             =   1320
         Width           =   1500
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4920
      TabIndex        =   26
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton cmd8 
         BackColor       =   &H00FFFFFF&
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
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         TabIndex        =   28
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image Image4 
         Height          =   2400
         Left            =   360
         Picture         =   "bill.frx":3F98
         Top             =   720
         Width           =   3750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contents cost"
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
      Left            =   4680
      TabIndex        =   23
      Top             =   2160
      Width           =   4455
      Begin VB.CommandButton cmd7 
         BackColor       =   &H00FFFFFF&
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
      Begin VB.Image Image3 
         Height          =   960
         Left            =   480
         Picture         =   "bill.frx":9A8E
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   25
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bill No"
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
      Left            =   4440
      TabIndex        =   20
      Top             =   2520
      Width           =   4455
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00FFFFFF&
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
      Begin VB.Label lbl11 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   22
         Top             =   480
         Width           =   3975
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   480
         Picture         =   "bill.frx":AF50
         Top             =   1680
         Width           =   1500
      End
   End
   Begin VB.TextBox txt10 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6840
      Width           =   3735
   End
   Begin VB.TextBox txt7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   5775
      Left            =   11520
      TabIndex        =   9
      Top             =   4200
      Width           =   2775
      Begin VB.CommandButton cmd4 
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
         Left            =   600
         TabIndex        =   19
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "CALCULATE"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmd3 
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
         Left            =   600
         TabIndex        =   11
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton cmd1 
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
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
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
      Height          =   735
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   8880
      Width           =   3495
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
      Height          =   1215
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3600
      Width           =   3855
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
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
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
      Height          =   615
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Customer ID:"
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
      Left            =   8400
      TabIndex        =   35
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lbl10 
      BackColor       =   &H8000000E&
      Caption         =   "Total Rs"
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
      TabIndex        =   17
      Top             =   6960
      Width           =   3615
   End
   Begin VB.Label lbl9 
      BackColor       =   &H8000000E&
      Caption         =   "Contents cost Rs:"
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
      Left            =   480
      TabIndex        =   16
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label lbl8 
      BackColor       =   &H8000000E&
      Caption         =   "Products:"
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
      Left            =   9120
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lbl7 
      BackColor       =   &H8000000E&
      Caption         =   "Contents quantity:"
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
      TabIndex        =   12
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   13080
      Picture         =   "bill.frx":C412
      Top             =   480
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   600
      Picture         =   "bill.frx":D8D4
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label lbl6 
      BackColor       =   &H8000000E&
      Caption         =   "Total Cost After Discount  Rs:"
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
      TabIndex        =   7
      Top             =   9000
      Width           =   3735
   End
   Begin VB.Label lbl5 
      BackColor       =   &H8000000E&
      Caption         =   "*5% Discount If The Amount Is  Rs.1000 or More!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   7920
      Width           =   8895
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Customer Name:"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lbl2 
      BackColor       =   &H8000000E&
      Caption         =   "Bill No:"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lbl1 
      BackColor       =   &H8000000E&
      Caption         =   " SHAKTI MILK BILL"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "Bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

'variable declaration
Dim intquantity As Integer

Dim curcost As Currency
Dim curtotc As Currency
Dim curtot As Currency
Dim curdis As Currency

'code  when form loads
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

'connection open
con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=bill"

'code for for making all the frames invisible even the form opens
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame8.Visible = False
Frame9.Visible = False

Adodc1.Visible = False

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
con.Close

Exit Sub
End Sub

'datareport
'code for print command button
Private Sub Command1_Click()
Dim b

b = txt1.Text
DataEnvironment1.Command1 b
DataReport1.Show
End Sub

'code for save command button
Private Sub cmd1_click()

Dim bl As Integer

If txt2.Text = " " Or txt3.Text = " " Or txt5.Text = " " Or txt10.Text = " " Or txt7.Text = " " Then
Frame6.Visible = True

Else

con.CursorLocation = adUseClient
'con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

'auto increment of bill number
cmd.CommandText = "select max(b_no) from bill"
On Error GoTo l1
Set rs = cmd.Execute
bl = rs.Fields(0)
txt1.Text = bl + 1

'inserting values in the table
cmd.CommandText = "insert into bill values('" & txt1.Text & "','" & txt11.Text & "','" & txt2.Text & "','" & Combo1.Text & "','" & txt3.Text & "','" & txt7.Text & "','" & txt10.Text & "','" & txt5.Text & "')"

Set rs = cmd.Execute
Frame5.Visible = True
con.Close

End If

Exit Sub
l1:
End Sub

'code of bill no frame's command button
Private Sub cmd10_Click()
Frame7.Visible = False
Bill.Show
End Sub

'code of contents quantity frame's command button
Private Sub cmd11_Click()
Frame8.Visible = False
Bill.Show
End Sub

'code for id frame's command
Private Sub cmd12_Click()
Frame9.Visible = False
Bill.Show
End Sub

'code for calculate command button
Private Sub cmd2_Click()
intquantity = Val(txt3.Text)
curcost = Val(txt7.Text)
curtot = intquantity * curcost
txt10.Text = curtot

'discount
curdis = 0

If (curtot >= 1000) Then
curdis = (curtot * 0.05)
curtotc = curtot - curdis
txt5.Text = curtotc
Else
txt5.Text = curtot
End If

End Sub

'code for clear command button
Private Sub cmd3_Click()
txt1.Text = " "
txt2.Text = " "
txt3.Text = " "
txt5.Text = " "
txt7.Text = " "
txt10.Text = " "
txt11.Text = " "
End Sub

'code for home command button
Private Sub cmd4_Click()
Menu.Show
Bill.Hide
Unload Bill
End Sub

'code of empty frame's command button
Private Sub cmd5_Click()
Frame2.Visible = False
Bill.Show
End Sub

'code of name frame's command button
Private Sub cmd6_Click()
Frame3.Visible = False
Bill.Show
End Sub

'code of contents cost frame's command button
Private Sub cmd7_Click()
Frame4.Visible = False
Bill.Show
End Sub

'code of success frame's command button
Private Sub cmd8_Click()
Frame5.Visible = False
Bill.Show
End Sub

'code of empty frame's command button
Private Sub cmd9_Click()
Frame6.Visible = False
Bill.Show
End Sub

'code of validation for cust id
Private Sub txt11_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
Frame9.Visible = True
txt11.Text = ""
End If
End Sub

'code for customer name text box
Private Sub txt2_Click()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=bill""Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=bill"

con.CursorLocation = adUseClient
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
On Error GoTo l1

cmd.CommandText = "select c_name from cus where c_id=" & txt11.Text

Set rs = cmd.Execute
txt2.Text = rs.Fields(0)
rs.Close
Exit Sub
l1:
MsgBox "Rec not found"

End Sub

'code of valdiation for name
Private Sub txt2_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 97 And KeyASCII <= 122) Or KeyASCII = 127 Or (KeyASCII >= 65 And KeyASCII <= 90) Or KeyASCII = 32 Or KeyASCII = 46 Or KeyASCII = 8) Then
Frame3.Visible = True
KeyASCII = 0
txt2.Text = ""
End If
End Sub

'code of valdiation for contents quantity
Private Sub txt3_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
Frame8.Visible = True
txt3.Text = ""
End If
End Sub

'code of valdiation for contents cost
Private Sub txt7_KeyPress(KeyASCII As Integer)
If Not ((KeyASCII >= 48 And KeyASCII <= 58) Or KeyASCII = 127 Or KeyASCII = 8) Then
Frame4.Visible = True
txt7.Text = ""
End If
End Sub

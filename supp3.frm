VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form supp3 
   Caption         =   "Supplier Details"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SUPPLIER DETAILS"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   8520
      Picture         =   "supp3.frx":0000
      Top             =   600
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "supp3.frx":14C2
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label lbl1 
      Caption         =   "SUPPLIER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "supp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'loading of form
Private Sub Form_Load()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

con.ConnectionString = "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=sup"

con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText

cmd.CommandText = "select * from sup"

Set rs = cmd.Execute
Set DataGrid1.DataSource = rs

End Sub


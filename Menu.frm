VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00FF8080&
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
   FillColor       =   &H00C000C0&
   ForeColor       =   &H00400040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Depositor Details"
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
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmd10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bill"
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
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Home"
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
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Details"
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
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exit"
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
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Employee Details"
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
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Supplier Details"
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
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Milk Details"
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
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4065
      Left            =   2640
      Picture         =   "Menu.frx":0000
      Top             =   2760
      Width           =   3600
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Welcome To SHAKTI Milk , Pune"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   7815
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   3720
      Picture         =   "Menu.frx":33CD
      Top             =   1560
      Width           =   1500
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'code for home command button
Private Sub cmd1_click()
shakti.Show
Menu.Hide
Unload Menu
End Sub

'code for bill command button
Private Sub cmd10_Click()
Bill.Show
Menu.Hide
Unload Menu
End Sub

'code for exit command button
Private Sub cmd2_Click()
End
End Sub

'code for depositors command button
Private Sub cmd3_Click()
menud.Show
Menu.Hide
Unload Menu
End Sub

'code for customer details command button
Private Sub cmd6_Click()
menuc.Show
Menu.Hide
Unload Menu
End Sub

'code for supplier details command button
Private Sub cmd7_Click()
menus.Show
Menu.Hide
Unload Menu
End Sub

'code for employee details command button
Private Sub cmd8_Click()
menue.Show
Menu.Hide
Unload Menu
End Sub

'code for milk details command button
Private Sub cmd9_Click()
milka.Show
Menu.Hide
Unload Menu
End Sub

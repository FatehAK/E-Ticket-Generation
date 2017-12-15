VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo3 
      Height          =   390
      Left            =   5640
      TabIndex        =   15
      Text            =   " -Select-"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   390
      Left            =   3000
      TabIndex        =   14
      Text            =   "Chennai"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   495
      Left            =   11040
      TabIndex        =   11
      Top             =   10440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   495
      Left            =   9720
      TabIndex        =   10
      Top             =   10440
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form5.frx":E6EF
      Height          =   4335
      Left            =   960
      OleObjectBlob   =   "Form5.frx":E703
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   16290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\NyteShady\Desktop\OOAD\E-Ticket\sea.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "disp"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   5640
      TabIndex        =   6
      Text            =   " -Select-"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   5640
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000011&
      Caption         =   " Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   20655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Results    :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Class                                     :"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Depature Date                  :"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "From"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search criteria      :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text3.Text = "" Then
MsgBox "Please enter the search keys", vbCritical, "Error"
Else
Data1.RecordSource = "Select * from disp where Source='" + Combo2.Text + "' and Destination='" + Combo3.Text + "' and Class='" + Combo1.Text + "'"
Data1.Refresh
DBGrid1.Visible = True
End If
End Sub

Private Sub Command2_Click()
If Text3.Text = "" Then
MsgBox "Please enter search keys", vbCritical, "Error"
Else
Me.Hide
Form6.Show
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
End Sub

Private Sub Form_Load()
Combo1.AddItem "Sleeper Class"
Combo1.AddItem "First AC"
Combo1.AddItem "Second AC"
Combo1.AddItem "Chair Car"
Combo3.AddItem "Bangalore"
Combo3.AddItem "New Delhi"
Combo3.AddItem "Hyderabad"
Combo3.AddItem "Agra"
End Sub


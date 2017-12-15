VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   -480
      Width           =   20295
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   9
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Users\NyteShady\Desktop\OOAD\E-Ticket\reg.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Register"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   495
         Left            =   18960
         TabIndex        =   8
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   6
         Top             =   7800
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   7800
         Width           =   975
      End
      Begin VB.TextBox Text2 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   9360
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   6720
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         X1              =   7320
         X2              =   8520
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Register here"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   7320
         TabIndex        =   7
         Top             =   7920
         Width           =   2655
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password             :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7320
         TabIndex        =   3
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name          :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7320
         TabIndex        =   2
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "E-Ticket System"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   1
         Top             =   1320
         Width           =   20295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Please enter details", vbCritical, "Error"
End If
Data1.RecordSource = "Select * from Register where uname='" + Text1.Text + "' and passwd='" + Text2.Text + "'"
Data1.Refresh
If Data1.Recordset.EOF Then
MsgBox "Login Failed, re-enter details ", vbCritical, "Error"
Text1.Text = ""
Text2.Text = ""
Else
MsgBox "Login Successful", vbInformation, "Success"
Me.Hide
Form4.Show
End If
End Sub

Private Sub Command2_Click()
Combo1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Label4_Click()
Unload Me
Form3.Show
End Sub

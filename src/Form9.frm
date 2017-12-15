VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form9"
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
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No. of Seats to Cancel              :"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected Seats                             :"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter PNR No.                             :"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   " Cancellation Form"
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
      TabIndex        =   0
      Top             =   0
      Width           =   20655
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim m As String
Dim n As String
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
MsgBox "Please enter the details", vbCritical, "Error"
Else
n = Text3.Text
m = MsgBox("Confirm cancellation of " + n + " tickets?", vbYesNo, "Confirmation")
If m = vbYes Then
Me.Hide
Form4.Show
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form4.Show
End Sub

VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form7"
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
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   13200
      TabIndex        =   26
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   13200
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\NyteShady\Desktop\OOAD\E-Ticket\sea.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "disp"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6600
      TabIndex        =   24
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   5400
      TabIndex        =   23
      Top             =   9000
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   390
      Left            =   6720
      TabIndex        =   22
      Top             =   8040
      Width           =   1935
   End
   Begin VB.ComboBox Combo5 
      Height          =   390
      Left            =   6480
      TabIndex        =   20
      Text            =   "Year"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      Height          =   390
      Left            =   5400
      TabIndex        =   19
      Text            =   "Month"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   390
      Left            =   4560
      TabIndex        =   18
      Text            =   "Day"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   16
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   15
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   14
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   13
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   12
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   10
      Text            =   "-Select-"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4560
      TabIndex        =   9
      Text            =   "-Select-"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount to be Paid      :"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000011&
      Caption         =   "  Payment Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   20535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IFSC Code                           :"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bank Name                        :"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Account No.                      :"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CVV No.                              :"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Expiration Date                 :"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name on Card                   :"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Card No.                              :"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Card Type                            :"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment Type                   :"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim res1 As Integer
Dim res2 As Integer
Dim res3 As Integer
Dim res4 As Integer
Dim m As String
Dim n As String
If Text1.Text = "" And Text2.Text = "" And Text4.Text = "" And Text5.Text = "" Then
MsgBox "Please enter the details", vbCritical, "Error"
Else
Data1.RecordSource = "Select * from disp where TrainName='" + Text3.Text + "'"
Data1.Refresh
Data1.Recordset.Edit
Randomize
res1 = Int((2500 * Rnd) + 1500)
Data1.Recordset!PNRNo = res1
Text8.Text = res1
Data1.Recordset.Update
Data1.Refresh
Data1.Recordset.Edit
Randomize
res2 = Int((9000 * Rnd) + 1500)
Data1.Recordset!CoachNo = res2
Data1.Recordset.Update
Data1.Refresh
Data1.Recordset.Edit
Randomize
res3 = Int((5 * Rnd) + 1)
Data1.Recordset!BerthNo = res3
Data1.Recordset.Update
Data1.Refresh
Data1.Recordset.Edit
Randomize
res4 = Int((64 * Rnd) + 1)
Data1.Recordset!SeatNo = res4
Data1.Recordset.Update
Data1.Refresh
n = Text8.Text
m = MsgBox("Ticket booked successfully your PNR No. is  " + n, vbOK, "Sucess")
If m = vbOK Then
Me.Hide
Form4.Show
End If
End If
End Sub

Private Sub Command2_Click()
Me.Hide
Form6.Show
End Sub

Private Sub Form_Load()
Form7.Text3.Text = Form6.Text8.Text
Form7.Text13.Text = Form6.Text13.Text
Combo1.AddItem "Credit Card"
Combo1.AddItem "Debit Card"
Combo1.AddItem "NetBanking"
Combo2.AddItem "MasterCard"
Combo2.AddItem "Visa"
Combo2.AddItem "SBIMaestro"
Combo2.AddItem "Rupay"
Dim i As Integer
For i = 1 To 31
Combo3.AddItem (i)
Next
For i = 1 To 12
Combo4.AddItem (i)
Next
For i = 2016 To 2030
Combo5.AddItem (i)
Next
End Sub


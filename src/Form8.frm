VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form8"
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
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8760
      TabIndex        =   35
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   7560
      TabIndex        =   34
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   33
      Text            =   "CHART PREPARED"
      Top             =   7080
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox Text13 
      DataField       =   "SeatNo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9960
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      DataField       =   "BerthNo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8280
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      DataField       =   "CoachNo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      DataField       =   "Departure Time"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   14280
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataField       =   "Arrival Time"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12360
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      DataField       =   "Class"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10680
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "Destination"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "Source"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "TrainName"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "TrainNo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\NyteShady\Desktop\OOAD\E-Ticket\sea.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "disp"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "Get Status"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   390
      Left            =   5400
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label19 
      Caption         =   "Charting Status       "
      Height          =   495
      Left            =   1920
      TabIndex        =   32
      Top             =   7080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Booking Details    :"
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
      Left            =   120
      TabIndex        =   26
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label18 
      Caption         =   "Seat No."
      Height          =   495
      Left            =   9960
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Berth No."
      Height          =   495
      Left            =   8280
      TabIndex        =   24
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "Coach No."
      Height          =   495
      Left            =   6600
      TabIndex        =   23
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Total Fare"
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label14 
      Caption         =   "No. of Passengers"
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Departure Time"
      Height          =   495
      Left            =   14280
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Arrival Time"
      Height          =   495
      Left            =   12360
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Class"
      Height          =   495
      Left            =   10680
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Destination"
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Source"
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Train Name"
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Train No."
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000011&
      Caption         =   " Status Information"
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
      TabIndex        =   6
      Top             =   0
      Width           =   20535
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000011&
      Caption         =   " Passenger Current Status Enquiry"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   20295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jorurney Details    :"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter PNR No.                                    :"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check Status        :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
If Text1.Text = "" Then
MsgBox "Please enter the PNR No.", vbCritical, "Error"
Else
Data1.RecordSource = "Select * from disp where PNRNo='" + Text1.Text + "'"
Data1.Refresh
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
Text11.Visible = True
Text12.Visible = True
Text13.Visible = True
Text14.Visible = True
End If
End Sub

Private Sub Command2_Click()
Me.Hide
Form4.Show
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub

Private Sub Form_Load()
Form8.Text9.Text = Form6.Text11.Text
Form8.Text10.Text = Form6.Text13.Text
End Sub

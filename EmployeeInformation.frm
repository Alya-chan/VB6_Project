VERSION 5.00
Begin VB.Form EmployeeInformation 
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\elize\Desktop\VB6_Project\employeeinformationDB.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employees"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "First Record"
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Record"
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Previous Record"
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Last Record"
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "LastName"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   12
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      DataField       =   "Region"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "PostalCode"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      DataField       =   "Country"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      DataField       =   "EmployeeID"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      DataField       =   "ReportsTo"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      DataField       =   "Extension"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      DataField       =   "Notes"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text14 
      DataField       =   "HomePhone"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      TabIndex        =   1
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      DataField       =   "HireDate"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   7200
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   495
      Left            =   840
      TabIndex        =   33
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   495
      Left            =   840
      TabIndex        =   32
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   495
      Left            =   840
      TabIndex        =   31
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "City"
      Height          =   495
      Left            =   840
      TabIndex        =   30
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Region"
      Height          =   495
      Left            =   840
      TabIndex        =   29
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Postal Code"
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Country"
      Height          =   495
      Left            =   840
      TabIndex        =   27
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Employee ID"
      Height          =   495
      Left            =   5760
      TabIndex        =   26
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Title"
      Height          =   495
      Left            =   5760
      TabIndex        =   25
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Report To"
      Height          =   495
      Left            =   5760
      TabIndex        =   24
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Hire Date"
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Extension"
      Height          =   495
      Left            =   5760
      TabIndex        =   22
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Note"
      Height          =   495
      Left            =   5760
      TabIndex        =   21
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Phone No"
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2648
      TabIndex        =   19
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "EmployeeInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
If Not Data1.Recordset.EOF Then
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MoveLast
End If
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.MovePrevious
End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
End
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Student_Information 
   Caption         =   "student information"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   ScaleHeight     =   8985
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton firstbtn 
      Caption         =   "First"
      Height          =   615
      Left            =   8040
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Prevtbtn 
      Caption         =   "Previous"
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton nxtbtn 
      Caption         =   "Next"
      Height          =   615
      Left            =   8040
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "Last"
      Height          =   615
      Left            =   8040
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8040
      TabIndex        =   10
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton clearbtn 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton delbtn 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "Update"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton addbtn 
      Caption         =   "Add"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtphone 
      DataField       =   "Phone"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   5640
      Width           =   4215
   End
   Begin VB.TextBox txtemail 
      DataField       =   "Email"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   4800
      Width           =   4215
   End
   Begin VB.TextBox txtadd 
      DataField       =   "Address"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox txtclass 
      DataField       =   "Class"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtroll 
      DataField       =   "RollNo"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
   End
   Begin MSAdodcLib.Adodc Studentdb 
      Height          =   615
      Left            =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elize\Desktop\VB6_Project\StudentDatabase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elize\Desktop\VB6_Project\StudentDatabase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from student_info"
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
   Begin VB.Label Label1 
      Caption         =   "Roll No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   21
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Adress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   18
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Phone No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2333
      TabIndex        =   15
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Student_Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addbtn_Click()
Studentdb.Recordset.AddNew
End Sub

Private Sub clearbtn_Click()
txtroll.Text = “”
txtname.Text = “”
txtclass.Text = “”
txtadd.Text = “”
txtemail.Text = “”
txtphone.Text = “”
End Sub

Private Sub Command9_Click()
End
End Sub

Private Sub delbtn_Click()
Confirmation = MsgBox("Do you want to delete this record?", vbYesNo + vbCritical, "Delete Record Confirmation")
If Confirmation = vbYes Then
Studentdb.Recordset.Delete
MsgBox " Record has been Deleted Successfully", vbYesNo, "Message"
Else
MsgBox " Record Not Deleted!!!", vbInformation, “Message”
End If
End Sub



Private Sub firstbtn_Click()
Studentdb.Recordset.MoveFirst
End Sub

Private Sub lastbtn_Click()
Studentdb.Recordset.MoveLast
End Sub

Private Sub nxtbtn_Click()
If Not Studentdb.Recordset.EOF Then
Studentdb.Recordset.MoveNext
If Studentdb.Recordset.EOF Then
Studentdb.Recordset.MoveLast
End If
End If
End Sub

Private Sub Prevtbtn_Click()
If Not Studentdb.Recordset.BOF Then
Studentdb.Recordset.MovePrevious
If Studentdb.Recordset.BOF Then
Studentdb.Recordset.MoveFirst
End If
End If
End Sub

Private Sub updatebtn_Click()
Studentdb.Recordset.Fields("RollNo") = txtroll.Text
Studentdb.Recordset.Fields("Name") = txtname.Text
Studentdb.Recordset.Fields("Class") = txtclass.Text
Studentdb.Recordset.Fields("Address") = txtemail.Text
Studentdb.Recordset.Fields("Phone") = txtphone.Text
Studentdb.Recordset.Update
MsgBox "Data Save Successfully", vbInformation, "Message"
End Sub

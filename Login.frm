VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   360
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elize\Desktop\VB6_Project\LoginDatabase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elize\Desktop\VB6_Project\LoginDatabase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * From LOGINTAB"
      Caption         =   "DATA"
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
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   8640
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "PASSWORD"
      DataSource      =   "Adodc1"
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "USERNAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   4920
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "USERNAME"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "LOGIN FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.RecordSource = "select * from LOGINTAB where USERNAME ='" + Text1.Text + "' and PASSWORD ='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "INVALID USERNAME OR PASSWORD"
Else
MDIForm1.Show
Login.Hide
End If
End Sub

Private Sub Command2_Click()
End
End Sub


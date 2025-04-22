VERSION 5.00
Begin VB.Form list_box 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Total NO .of Customer "
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Cutomer Lists"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter The Name of new Customer"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "list_box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
List1.AddItem Text1.Text
Text1.Text = ""
Label4.Caption = List1.ListCount
End Sub

Private Sub Command2_Click()
Dim ind As Integer
ind = List1.ListIndex
If ind >= 0 Then
List1.RemoveItem ind
Label4.Caption = List1.ListCount
End If
End Sub

Private Sub Command3_Click()
List1.Clear
Label4.Caption = List1.ListCount
End Sub

Private Sub Text1_Change()
Command1.Enabled = (Len(Text1.Text) > 0)
End Sub

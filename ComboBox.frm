VERSION 5.00
Begin VB.Form ComboBox 
   Caption         =   "Combo Box"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Select"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Combo Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   953
      TabIndex        =   4
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "ComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" Then
Combo1.AddItem Text1.Text
MsgBox "Ietm Added"
Text1.Text = ""
Else
MsgBox "No input"
End If
End Sub

Private Sub Command2_Click()
If Combo1.ListCount > 0 Then
Combo1.RemoveItem ListIndex
Else
MsgBox "Empty List"
End If
End Sub



Private Sub Command3_Click()
End
End Sub

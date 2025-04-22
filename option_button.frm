VERSION 5.00
Begin VB.Form option_button 
   Caption         =   "Option Button"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "UNDERLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ITALIC"
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "BOLD"
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "GREEN"
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "BLUE"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "RED"
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ENETR  STRING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "option_button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.FontBold = True
Else
Text1.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text1.FontItalic = True
Else
Text1.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Text1.FontUnderline = True
Else
Text1.FontUnderline = False
End If

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text1.ForeColor = vbRed
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text1.ForeColor = vbBlue
End If
End Sub


Private Sub Option3_Click()
If Option3.Value = True Then
Text1.ForeColor = vbGreen
End If
End Sub














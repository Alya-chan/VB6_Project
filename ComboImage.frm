VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ComboImage 
   Caption         =   "Combo_Image"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ComboImage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ComboImage.frx":1AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ComboImage.frx":400C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ComboImage.frx":9D95
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ComboImage.frx":ED11
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "Select"
   End
End
Attribute VB_Name = "ComboImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Set ImageCombo1.ImageList = ImageList1
ImageCombo1.ComboItems.Add , , "Administrator", 1
ImageCombo1.ComboItems.Add , , "Staff", 2
ImageCombo1.ComboItems.Add , , "Office Assistant", 3
ImageCombo1.ComboItems.Add , , "Librarian", 4
ImageCombo1.ComboItems.Add , , "Storekeeper", 5
End Sub

Private Sub ImageCombo1_Click()
Text1.Text = ImageCombo1.Text
End Sub

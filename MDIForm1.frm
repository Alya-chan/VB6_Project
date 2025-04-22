VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu project1 
      Caption         =   "Project 1"
      Begin VB.Menu option 
         Caption         =   "Option Button"
      End
      Begin VB.Menu list 
         Caption         =   "List Box"
      End
      Begin VB.Menu box 
         Caption         =   "Combo Box"
      End
   End
   Begin VB.Menu project2 
      Caption         =   "Project 2"
      Begin VB.Menu cbimage 
         Caption         =   "Combo Image"
      End
      Begin VB.Menu studentinformation 
         Caption         =   "Student Information"
      End
      Begin VB.Menu emp_info 
         Caption         =   "Employee Information"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub box_Click()
ComboBox.Show
End Sub

Private Sub cbimage_Click()
ComboImage.Show
End Sub

Private Sub emp_info_Click()
EmployeeInformation.Show
End Sub

Private Sub list_Click()
list_box.Show
End Sub

Private Sub option_Click()
option_button.Show
End Sub

Private Sub studentinformation_Click()
Student_Information.Show
End Sub

Option Compare Database
Option Explicit


' ----------------------------------------------------------------
' Procedure Name:   example_Git_Clone
' Purpose:          Get and import objects from given repositories
' Procedure Kind:   Sub
' Author:           Thomas Braun
' Date:             12.07.2023
' Procedure Access: Private
' ----------------------------------------------------------------
Private Sub example_Git_Clone()
  
  git_Import "https://github.com/berlin-braun/my_VbaRegistry"
  git_Import "https://github.com/berlin-braun/my_VbaDrive"
  git_Import "https://github.com/berlin-braun/my_VbaObject"
  
End Sub
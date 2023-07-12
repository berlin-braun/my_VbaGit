Option Compare Database
Option Explicit

' factory for my_Git_Object

' ----------------------------------------------------------------
' Procedure Name:   git_Import
' Purpose:          clone git-repository and import object to current database
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             12.07.2023
' Procedure Access: Public
' Parameter str_Repository (String): URL or filepath to repository
' Definition:       scripts for objects are in specific folder (please set to definition in the repository)
' ----------------------------------------------------------------
Public Function git_Import(ByVal str_Repository As String)
  
  Dim git As New my_Git_Object
  
  git.repository_URL = str_Repository                           ' reference to repository
  
  git.Clone                                                     ' clone to local working directory
  
  git.module_Subfolder = "source\module"                        ' subfolder for module-scripts
  git.form_Subfolder = "source\form"                            ' subfolder for form-scripts
  git.query_Subfolder = "source\query"                          ' subfolder for query-scripts
  git.report_Subfolder = "report\form"                          ' subfolder for report-scripts
  
  git.Import                                                    ' import object from working-directory to current database
  
  git.Cleanup                                                   ' cleanup
  
  Set git = Nothing
  
End Function
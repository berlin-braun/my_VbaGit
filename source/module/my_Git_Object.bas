Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_repository_URL            As String
Private m_repository_Directory      As String
Private m_form_Subfolder            As String
Private m_query_Subfolder           As String
Private m_report_Subfolder          As String
Private m_module_Subfolder          As String

Private m_import_object             As New Collection


'
' start: resources for 'ExecCmd'
' source
' https://learn.microsoft.com/de-de/office/vba/access/concepts/windows-api/determine-when-a-shelled-process-ends
'
Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessID As Long
  dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long _
                                                           , ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long _
                                                      , ByVal lpCommandLine As String _
                                                      , ByVal lpProcessAttributes As Long _
                                                      , ByVal lpThreadAttributes As Long _
                                                      , ByVal bInheritHandles As Long _
                                                      , ByVal dwCreationFlags As Long _
                                                      , ByVal lpEnvironment As Long _
                                                      , ByVal lpCurrentDirectory As Long _
                                                      , lpStartupInfo As STARTUPINFO _
                                                      , lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
' end: resources for 'ExecCmd'
'
'


' start: class

Private Sub Class_Initialize()

' ToDo

End Sub

Private Sub Class_Terminate()

'  Cleanup

End Sub

' end: class

    
' start: public property
    
Public Property Get repository_URL() As String                      ' URL or Path of the repository

  repository_URL = m_repository_URL

End Property

Public Property Let repository_URL(ByVal sNewValue As String)       ' URL or path of the repository

  m_repository_URL = sNewValue

End Property

Public Property Get repository_Name() As String                     ' name of the repository
  Dim str_Ret As String
  
  str_Ret = repository_URL                                          ' URL/filepath
  
  If InStr(str_Ret, "\") > 0 Then                                   ' extract from filepath
    str_Ret = Mid(str_Ret, InStrRev(str_Ret, "\") + 1)
  End If
  If InStr(str_Ret, "/") > 0 Then                                   ' extract from URL
    str_Ret = Mid(str_Ret, InStrRev(str_Ret, "/") + 1)
  End If
  
  repository_Name = str_Ret

End Property

Public Property Get repository_Directory() As String                ' working directory

  If Len(m_repository_Directory) = 0 Then                           ' not set
    m_repository_Directory = CurrentProject.Path & _
                             "\" & repository_Name & _
                             "_" & format(Now, "yyyymmdd_hhnnss")   ' add directory in current directory
    MkDir m_repository_Directory
  End If
  
  repository_Directory = m_repository_Directory

End Property

Public Property Let repository_Directory(ByVal sNewValue As String) ' working directory

  m_repository_Directory = sNewValue

End Property

Public Property Get form_Subfolder() As String                      ' folder for form-scripts

  form_Subfolder = m_form_Subfolder

End Property

Public Property Let form_Subfolder(ByVal sNewValue As String)       ' folder for form-scripts

  m_form_Subfolder = sNewValue

End Property

Public Property Get query_Subfolder() As String                     ' folder for query-scripts

  query_Subfolder = m_query_Subfolder

End Property

Public Property Let query_Subfolder(ByVal sNewValue As String)      ' folder for query-scripts

  m_query_Subfolder = sNewValue

End Property

Public Property Get report_Subfolder() As String                    ' folder for report-scripts

  report_Subfolder = m_report_Subfolder

End Property

Public Property Let report_Subfolder(ByVal sNewValue As String)     ' folder for report-scripts

  m_report_Subfolder = sNewValue

End Property

Public Property Get module_Subfolder() As String                    ' folder for module-scripts

  module_Subfolder = m_module_Subfolder

End Property

Public Property Let module_Subfolder(ByVal sNewValue As String)     ' folder for module-scripts

  m_module_Subfolder = sNewValue

End Property

' end: public property


' start: private property

Private Property Get import_Folder(ByVal object_Type As AcObjectType) As String
  Dim str_Ret As String
  
  str_Ret = ""
  
  Select Case object_Type
    Case Is = acQuery
      str_Ret = query_Subfolder
      
    Case Is = acForm
      str_Ret = form_Subfolder
      
    Case Is = acReport
      str_Ret = report_Subfolder
      
    Case acModule, acClassModule
      str_Ret = module_Subfolder
    
  End Select
  
  If Len(str_Ret) > 0 Then
    str_Ret = "\" & str_Ret
  End If
  
  str_Ret = repository_Directory & str_Ret
  
  import_Folder = str_Ret

End Property

' end: private property


' start: public function

Public Function Clone()
  
  ExecCmd "git clone " & _
          """" & m_repository_URL & """ " & _
          """" & repository_Directory & """"                      ' Execute clone-command
  
End Function

Public Function Import()
  
  Set m_import_object = Nothing                                   ' reset collection

  collect_Files_Folder m_import_object _
                     , repository_Directory _
                     , True                                       ' collect files from working directory
                          
  import_Object acQuery                                           ' import collected files to objects
  import_Object acForm
  import_Object acReport
  import_Object acModule
  
End Function

Public Function Cleanup()

  ' Todo: delete local working-directory
  
  m_repository_Directory = ""

End Function

' end: public function


' start: private function

Private Sub ExecCmd(cmdline As String)
  ' source:
  ' https://learn.microsoft.com/de-de/office/vba/access/concepts/windows-api/determine-when-a-shelled-process-ends
  Dim proc        As PROCESS_INFORMATION
  Dim start       As STARTUPINFO
  Dim ReturnValue As Integer
  
  start.cb = Len(start)                                             ' Initialize the STARTUPINFO structure:
  
  ReturnValue = CreateProcessA(0&, cmdline$, 0&, 0&, 1& _
                             , NORMAL_PRIORITY_CLASS, 0& _
                             , 0&, start, proc)                     ' Start the shelled application:
  
  Do                                                                ' Wait for the shelled application to finish:
    ReturnValue = WaitForSingleObject(proc.hProcess, 0)
    DoEvents
  Loop Until ReturnValue <> 258
  
  ReturnValue = CloseHandle(proc.hProcess)

End Sub

Private Function import_Object(ByVal object_Type As AcObjectType)
  Dim str_Name    As String
  Dim str_File    As String
  Dim str_Path    As String
  Dim str_Folder  As String
  Dim c           As Long
  
  str_Folder = import_Folder(object_Type)                                 ' load relevant folder for object
  
  For c = 1 To m_import_object.Count                                      ' iterate files
    
    If Left(m_import_object(c), Len(str_Folder)) = str_Folder Then        ' file is in folder
      
      str_File = m_import_object(c)                                       ' get file
      
      Select Case Right(str_File, 4)                                      ' only '.bas', '.txt', '.cls' files
        Case Is = ".txt", ".bas", ".cls"
        
        str_Name = Dir(str_File)                                          ' filename
        str_Path = Left(str_File, Len(str_File) - Len(str_Name) - 1)      ' filepath
        
        If str_Path = str_Folder Then                                     ' filepath matches folder
        
          str_Name = Left(str_Name, InStr(str_Name, ".") - 1)             ' extract object-name from file
          Application.LoadFromText object_Type, str_Name, str_File        ' import object
          
        End If
          
      End Select
      
    End If
  Next
  
End Function

Private Function collect_Files_Folder(ByRef myFile As Collection _
                                    , ByVal str_Folder As String _
                           , Optional ByVal include_Subfolder As Boolean = True)
  Dim fso_Folder      As Object
  Dim fso_SubFolder   As Object
  Dim fso_File        As Object
  
  Set fso_Folder = CreateObject("Scripting.FileSystemObject").GetFolder(str_Folder)   ' get foldern
  
  For Each fso_File In fso_Folder.Files                                   ' iterate files
    myFile.Add fso_File.Path                                              ' add file to collection
  Next
  
  If include_Subfolder = True Then                                        ' search in subfolders
    For Each fso_SubFolder In fso_Folder.SubFolders                       ' iterate subfolder
      If Left(fso_SubFolder.Name, 1) <> "." Then
        collect_Files_Folder myFile _
                           , fso_SubFolder.Path _
                           , include_Subfolder                            ' add files recursive
      End If
    Next
  End If
  
  Set fso_Folder = Nothing
  Set fso_SubFolder = Nothing
  Set fso_File = Nothing
  
End Function

' end: private function
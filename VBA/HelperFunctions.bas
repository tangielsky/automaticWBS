Option Explicit

Function GetFolder(s As String) As String
Dim result As String
Dim i As Integer

  result = ""
  For i = Len(s) To 1 Step -1
    If Mid$(s, i, 1) = Application.PathSeparator Then
      result = Left$(s, i)
      Exit For
    End If
  Next
  
  GetFolder = result
End Function

Sub OpenDocument(document As String)
Dim path As String

On Error Resume Next

  'URL?
  If LCase(Left$(document, 4)) <> "http" Then
  
    'relativer oder absoluter Pfad?
    If Left$(document, 2) = "." & Application.PathSeparator Then
      path = Sheets("Setup").Range("PATH_DOCUMENTS")
      If path = "" Then path = ActiveWorkbook.path
      If Right$(path, 1) <> Application.PathSeparator Then
        path = path & Application.PathSeparator
      End If
      document = path & Right$(document, Len(document) - 2)
    End If
  End If
  
  ActiveWorkbook.FollowHyperlink document
End Sub

Function ExtractFilename(FilePath As String) As String
Dim result As String

  With CreateObject("Scripting.FileSystemObject")
    result = .GetFileName(FilePath)
    'extName = .GetExtensionName(FilePath)
    result = .GetBaseName(FilePath)
    'parentName = .GetParentFolderName(FilePath)
  End With
  
  ExtractFilename = result
End Function

Function ConvertControlCharsToSpace(s As String) As String
Dim result As String
Dim i As Integer

  result = ""
  For i = 1 To Len(s)
    If Asc(Mid$(s, i, 1)) < 32 Then
      result = result & " "
    Else
      result = result & Mid$(s, i, 1)
    End If
  Next
  
  ConvertControlCharsToSpace = result
End Function

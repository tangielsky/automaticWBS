VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinkedDocumentsForm 
   Caption         =   "Verknüpfte Dokumente"
   ClientHeight    =   3516
   ClientLeft      =   156
   ClientTop       =   588
   ClientWidth     =   5268
   OleObjectBlob   =   "LinkedDocumentsForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "LinkedDocumentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Zerlegt einen Dokument Text aus mehreren Titel und verknüpften Dokumente (Pfade, URLs) in 2 Arrays
'Splits the documents text into 2 arrays
Sub UpdateDocuments(documents As String)
Dim docArray As Variant
Dim i As Integer
Dim j As Integer
  
  'split everything into one helper array
  docArray = Split(documents, "|")
  
  j = (UBound(docArray, 1) + 1) / 2
  If j < 1 Then j = 1
  ReDim document(j)
  ReDim title(j)
  
  'fill the title and document array
  j = 0
  For i = LBound(docArray, 1) To UBound(docArray, 1) Step 2
    j = j + 1
    title(j) = docArray(i)
    document(j) = docArray(i + 1)
    If title(j) = "" Then
      title(j) = ExtractFilename(document(j))
    End If
  Next
  
  ListBox1.Clear
  For i = 1 To j
    ListBox1.AddItem title(i)
  Next
  
  If ListBox1.ListCount > 0 Then
    ListBox1.Selected(0) = True
  End If
End Sub

'Öffnet ein Dokument oder URL
'Open a document or link
Private Sub CommandButton1_Click()
Dim doc As String

  If ListBox1.Value = "" Then Exit Sub
  
  doc = document(ListBox1.ListIndex + 1)
  Unload Me
  Call OpenDocument(doc)
End Sub

'Öffnet den Ordner im Explorer
'Opens the folder in Explorer
Private Sub CommandButton3_Click()
Dim doc As String

  If ListBox1.Value = "" Then Exit Sub
  
  doc = GetFolder(document(ListBox1.ListIndex + 1))
  
  Unload Me
  Call OpenDocument(doc)
End Sub


Private Sub ListBox1_Change()
On Error Resume Next

  TextBox1 = document(ListBox1.ListIndex + 1)
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call CommandButton1_Click
End Sub


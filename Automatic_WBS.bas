Attribute VB_Name = "AutomaticWBS"
'Macro:  Automatic work-breakdown structure WBS / Projektstrukturplan PSP
'Author: Thomas Angielsky

'https://techpluscode.de/projekt-struktur-plan-psp-in-3-minuten/
'https://techpluscode.de/work-breakdown-structure-wbs-in-3-minutes/

'Version 05: 08.07.2020
'            reduced links only to:
'            - Visual Basic for Applications
'            - Microsoft Excel 16.0 Object Library
'            - Microsoft Office 16.0 Object Library
'            - Microsoft Forms 2.0 Object Library
'            for better support of Office for Mac
'
'Version 04: 06.04.2020
'            corrections:
'            - clear x,y area before routine starts
'

'Version 03: 02.02.2020
'            - functions for progress visualization
'            - expanded to 10 user fields
'
'Version 02: 04.01.2020
'            - more parameters in sheet setup
'            - some corrections



Option Explicit



'Datentyp für Koordinaten und Anzahl Kindelemente
'Datatype for coordinates and number of children items
Type Position
  x As Double
  y As Double
  count As Integer
End Type



'Positionen in der Tabelle "Start"
'Positions of sheet "Start"
Const COL_CODE = 1
Const COL_NAME = 2
Const COL_PROGRESS = 3
Const COL_FIELDS = 4
Const COL_X = 15
Const COL_Y = 16
Const COL_COUNT = 17
Const ROW_START = 5




'Löscht die PSP-Struktur im aktuellen Tabellenblatt
'Deletes the WBS structure in the current sheet
Sub DeleteWorkBreakdownStructure()
Dim i As Integer
Dim shape1
  
  For Each shape1 In ActiveSheet.Shapes
    If Left$(shape1.name, 2) = "N_" Then
      shape1.Delete
    End If
  Next
End Sub



'Hauptfunktion zum Erzeugen der PSP-Struktur im aktuellen Tabellenblatt
'Main function to create the WBS structure in the current sheet
Sub CreateWorkBreakdownStructure()
Dim row As Integer
Dim wbsCode As String
Dim wbsCodeParent As String
Dim wbsOld As String
Dim wbsLevel1Old As String
Dim level As Integer
Dim levelOld As Integer
Dim p As Position
Dim pParent As Position
Dim caption As String
Dim w As Double
Dim h As Double
Dim d As Double
Dim spaceYLevel0 As Double
Dim spaceXLevel1 As Double
Dim spaceYLevel3 As Double
Dim spaceXLevel3 As Double
Dim i As Integer
Dim progress As Double
Dim firstshape As String


  Call DeleteWorkBreakdownStructure
  Call DeleteInternalValues
  

  'Einstellungen einlesen
  'Get setup values
  spaceYLevel0 = Sheets("Setup").Range("LEVEL0_SPACE_Y")
  spaceXLevel1 = Sheets("Setup").Range("LEVEL1_SPACE_X")
  spaceYLevel3 = Sheets("Setup").Range("LEVEL3_SPACE_Y")
  spaceXLevel3 = Sheets("Setup").Range("LEVEL3_SPACE_X")

  'Startbedingungen setzen
  'Set initialization values
  p.x = 0
  p.y = 0
  row = ROW_START
  wbsOld = ""
  wbsLevel1Old = ""
  levelOld = 0
  
  'Struktur in Tabelle "Start" durchlaufen
  'Run through structure in "Start" sheet
  Do
    'Eigenschaften jedes Elements einlesen
    'Read parameters of each item
    wbsCode = Sheets("Start").Cells(row, COL_CODE)
    wbsCodeParent = GetParentWBS(wbsCode)
    pParent = GetLastPosition(wbsCodeParent)
    
    If Sheets("Start").Cells(row, COL_PROGRESS) = "" Then
      progress = 0
    Else
      progress = CDbl(Sheets("Start").Cells(row, COL_PROGRESS))
    End If
    
    level = CountPoints(wbsCode)
    If level = 0 Then 'usualy projectname
      caption = Sheets("setup").Shapes("LEVEL_1").TextFrame.Characters.Text
      w = Sheets("Setup").Shapes("LEVEL_1").Width
      h = Sheets("Setup").Shapes("LEVEL_1").Height
      p.y = h
    ElseIf level = 1 Then 'phase or part project
      caption = Sheets("setup").Shapes("LEVEL_2").TextFrame.Characters.Text
      w = Sheets("Setup").Shapes("LEVEL_2").Width
      h = Sheets("Setup").Shapes("LEVEL_2").Height
      If pParent.count = 0 Then
        p.x = 0
        p.y = p.y + h * spaceYLevel0
      Else
        d = GetLastMaxPosition(wbsLevel1Old)
        If d <> 0 Then
          p.x = d + spaceXLevel1 * w
        Else
          p.x = pParent.x + spaceXLevel1 * w
        End If
        p.y = pParent.y
      End If
    Else
      caption = Sheets("setup").Shapes("LEVEL_3").TextFrame.Characters.Text
      w = Sheets("Setup").Shapes("LEVEL_3").Width
      h = Sheets("Setup").Shapes("LEVEL_3").Height
      
      If pParent.count = 0 Then
        p.x = pParent.x + w * spaceXLevel3
        If level = 2 Then
          d = Sheets("Setup").Shapes("LEVEL_2").Height
        Else
          d = h
        End If
        p.y = pParent.y + d * spaceYLevel3
      Else
        p.x = pParent.x
        p.y = pParent.y + h * spaceYLevel3
      End If
    End If
    
    
    'Element-Eigenschaften in Tabelle "Start" speichern
    'Save item properties in sheet "Start"
    p.count = pParent.count + 1
    Call SetPositions(wbsCodeParent, p)
    
    Call SetPosition(wbsCodeParent, p)
    
    p.count = 0
    Call SetPosition(wbsCode, p)
    
    
    'Soll Formatierung über Leistungsfortschritt erfolgen?
    'Should formatting be based on performance progress?
    If UCase(Sheets("Setup").Range("PROGRESS_COLORS")) = "J" Then
      If progress = 0 Then
        caption = Replace(caption, "$PROGRESS", Sheets("Setup").Range("FORMAT_NOT_STARTED"))
      ElseIf progress = 1 Then
        caption = Replace(caption, "$PROGRESS", Sheets("Setup").Range("FORMAT_COMPLETED"))
      Else
        caption = Replace(caption, "$PROGRESS", Sheets("Setup").Range("FORMAT_IN_PROGRESS"))
      End If
    End If
   
   
    'Variablen austauschen mit echten Werten
    'Replace variables with real values
    caption = Replace(caption, "$CODE", Sheets("Start").Cells(row, COL_CODE))
    caption = Replace(caption, "$NAME", Sheets("Start").Cells(row, COL_NAME))
    caption = Replace(caption, "$PROGRESS", Format(progress, "0%"))
    
    For i = 10 To 1 Step -1
      caption = Replace(caption, "$F" & CStr(i), Sheets("Start").Cells(row, COL_FIELDS + i - 1))
    Next
    
    
    'Rechteck-Shape und Verbindungslinie einfügen
    'Insert rectangle shape and connector
    Call InsertRectangle(wbsCode, caption, p.x, p.y, w, h, level, progress)
    If level > 0 Then Call InsertConnector(wbsCodeParent, wbsCode, level)
    
    'Nächste Zeile vorbereiten
    'Prepare next line
    If row = ROW_START Then firstshape = "N_" & wbsCode
    
    row = row + 1
    levelOld = level
    wbsOld = wbsCode
    If level = 1 Then wbsLevel1Old = wbsCode
    wbsCode = Sheets("Start").Cells(row, COL_CODE)
    
  Loop Until wbsCode = ""

  'Erstes Element zentrieren
  'Center first item
  If firstshape <> "" Then
    d = FindXmax() + Sheets("Setup").Shapes("LEVEL_3").Width
    ActiveSheet.Shapes("N_1").Left = (d - Sheets("Setup").Shapes("LEVEL_1").Width) / 2
  End If

  ActiveSheet.Cells(1, 1).Select
End Sub

'Löscht die internen Werte für X, Y, Count vor der Neuberechnung des PSP
'Deletes the internal values of X, Y, Count before a recalculation of WBS
Sub DeleteInternalValues()
Dim row As Integer
Dim result As Double
Dim w As String

  row = ROW_START
  Do
    w = Sheets("Start").Cells(row, COL_CODE)
    If w <> "" Then
      Sheets("Start").Cells(row, COL_X) = ""
      Sheets("Start").Cells(row, COL_Y) = ""
      Sheets("Start").Cells(row, COL_COUNT) = ""
    End If

    row = row + 1
  Loop Until w = ""

End Sub

'Zählt die Anzahl der Punkte im String => Strukturebene
'Counts the amount of points in the string => structure level
Function CountPoints(s As String) As Integer
Dim result As Integer
Dim i As Integer

  result = 0
  For i = 1 To Len(s)
    If Mid$(s, i, 1) = "." Then result = result + 1
  Next
  CountPoints = result
End Function



'Findet den maximalen X Wert
'Get the max X value
Function FindXmax() As Double
Dim row As Integer
Dim result As Double
Dim w As String

  row = ROW_START
  result = 0
  Do
    w = Sheets("Start").Cells(row, COL_CODE)
    If w <> "" Then
      If Sheets("Start").Cells(row, COL_X) > result Then
        result = Sheets("Start").Cells(row, COL_X)
      End If
    End If

    row = row + 1
  Loop Until w = ""
  FindXmax = result
End Function



'Findet then übergeordneten PSP-Code
'Gets the parent WBS code
Function GetParentWBS(wbsCode As String) As String
Dim i As Integer
Dim result As String

  result = ""
  For i = Len(wbsCode) To 1 Step -1
    If Mid$(wbsCode, i, 1) = "." Then
      result = Left$(wbsCode, i - 1)
      Exit For
    End If
  Next
  
  GetParentWBS = result
End Function



'Findet den PSP-Code der 2. Ebene
'Gets the WBS-Code of the 2nd level
Function GetLevel2WBS(wbsCode As String) As String
Dim i As Integer
Dim j As Integer
Dim result As String

  result = ""
  j = 0
  For i = 1 To Len(wbsCode)
    If Mid$(wbsCode, i, 1) = "." Then
      j = j + 1
      If j = 2 Then
        result = Left$(wbsCode, i - 1)
        Exit For
      End If
    End If
  Next
  
  GetLevel2WBS = result
End Function



'Erzeugt ein Rechteckshape
'Creates a rectangle shape
Sub InsertRectangle(name As String, caption As String, x As Double, y As Double, w As Double, h As Double, level As Integer, progress As Double)
Dim s As String
Dim c As Long
Dim id As String


'If you want to use other shapes, have a look at:
'https://docs.microsoft.com/de-de/office/vba/api/office.msoautoshapetype
  ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h).Select
  id = "N_" & name
  Selection.name = id
 
  If level > 1 Then
    s = "3"
  Else
    s = CStr(level + 1)
  End If
  
  Sheets("Setup").Shapes("LEVEL_" & s).PickUp
  ActiveSheet.Shapes(id).Apply
  
  If UCase(Sheets("Setup").Range("PROGRESS_COLORS")) = "J" Then
    If progress = 0 Then
     Selection.ShapeRange.Fill.ForeColor.RGB = Sheets("Setup").Range("PROGRESS_NOT_STARTED").Cells.Interior.Color
    ElseIf progress = 1 Then
      Selection.ShapeRange.Fill.ForeColor.RGB = Sheets("Setup").Range("PROGRESS_COMPLETED").Cells.Interior.Color
    Else
      Selection.ShapeRange.Fill.ForeColor.RGB = Sheets("Setup").Range("PROGRESS_IN_PROGRESS").Cells.Interior.Color
    End If
  End If
      
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = caption
End Sub



'Fügt eine Verbindungslinie zwischen 2 PSP-Elementen ein
'Adds a connrector between to rectangle shapes
Sub InsertConnector(wbsCodeFrom As String, wbsCodeTo As String, level As Integer)
Dim pFrom As Integer
Dim pTo As Integer

  If level = 1 Then
    pFrom = 3
    pTo = 1
  Else
    pFrom = 3
    pTo = 2
  End If

   ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100).Select
   Selection.name = "N_" & wbsCodeFrom & "_" & wbsCodeTo
   Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
   Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes("N_" & wbsCodeFrom), pFrom
   Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("N_" & wbsCodeTo), pTo
      
   Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadStealth
   With Selection.ShapeRange.Line
     .EndArrowheadLength = msoArrowheadLong
     .EndArrowheadWidth = msoArrowheadWide
     .Visible = msoTrue
     .Weight = 1
     .Transparency = 0
   
   End With
   
  Sheets("Setup").Shapes("CONNECTOR").PickUp
  ActiveSheet.Shapes("N_" & wbsCodeFrom & "_" & wbsCodeTo).Apply
End Sub
      
      
      
'Findet die Koordinaten der letzten Position
'Gets the coordinates of the last position
Function GetLastPosition(wbsCode As String) As Position
Dim row As Integer
Dim result As Position
Dim found As Boolean
Dim w As String

  found = False
  row = ROW_START
  Do
    w = Sheets("Start").Cells(row, COL_CODE)
    If wbsCode = w Then
      found = True
      result.x = Sheets("Start").Cells(row, COL_X)
      result.y = Sheets("Start").Cells(row, COL_Y)
      result.count = Sheets("Start").Cells(row, COL_COUNT)
      
    End If

    row = row + 1
  Loop Until w = "" Or found = True

  GetLastPosition = result
End Function



'Findet die Koordinaten der letzten max. Position
'Gets the coordinates of the last max. position
Function GetLastMaxPosition(wbsCode As String) As Double
Dim row As Integer
Dim result As Double
Dim w As String
Dim wbsLevel2 As String
  
  row = ROW_START
  Do
    w = Sheets("Start").Cells(row, COL_CODE)
    If (wbsCode = w) Or ((wbsCode = Left$(w, Len(wbsCode))) And (wbsCode <> "")) Then
      If result < Sheets("Start").Cells(row, COL_X) Then result = Sheets("Start").Cells(row, COL_X)
    End If

    row = row + 1
  Loop Until w = ""
  
  GetLastMaxPosition = result
End Function



'Setzt die Koordinaten eines PSP-Elements
'Saves the coordinates of a WBS item
Sub SetPosition(wbsCode As String, p As Position)
Dim row As Integer
Dim result As Position
Dim found As Boolean
Dim w As String

  found = False
  row = ROW_START
  Do
    w = Sheets("Start").Cells(row, COL_CODE)
    If wbsCode = w Then
      found = True
      Sheets("Start").Cells(row, COL_X) = p.x
      Sheets("Start").Cells(row, COL_Y) = p.y
      Sheets("Start").Cells(row, COL_COUNT) = p.count
    End If

    row = row + 1
  Loop Until w = "" Or found = True
End Sub



'Setzt alle übergeordneten Koordinaten eines PSP-Elements
'Saves the parent coordinates of a WBS item
Sub SetPositions(wbsCode As String, p As Position)
Dim row As Integer
Dim result As Position
Dim w As String
Dim wbsLevel2 As String


  wbsLevel2 = GetLevel2WBS(wbsCode)
  
  row = ROW_START
  Do
    w = Sheets("Start").Cells(row, COL_CODE)
    If (wbsCode = w) Or ((wbsLevel2 = Left$(w, Len(wbsLevel2))) And (wbsLevel2 <> "")) Then
      Sheets("Start").Cells(row, COL_COUNT) = p.count
      Sheets("Start").Cells(row, COL_Y) = p.y
    End If

    row = row + 1
  Loop Until w = ""
End Sub




Attribute VB_Name = "ModuleVariousScripts"
Option Private Module

Sub LoadFile(wb, o, s, n)
    ' Goal: Open Workbooks
    
    ' wb: Workbook
    ' o:  Worksheet
    ' s:  Address
    ' n:  Name or index of worksheet
    
    Dim b As Boolean
    Dim t As String
    
    ' Remove path from filename
    t = Right(s, Len(s) - InStrRev(s, "\"))
    b = False
    
    ' Go through all open workbooks and address the specific one. If it is not open, open and address it.
    For Each Workbook In Workbooks
        If Workbook.Name = t Then
            b = True
            Set wb = Workbook
            Set o = wb.Sheets(n)
            Exit For
        End If
    Next
    If b = False Then
        Workbooks.Open s
        Set wb = ActiveWorkbook
        Set o = wb.Sheets(n)
    End If
End Sub

Function FindColumn(o, s, i)
    ' Goal: Find and address column
    
    ' o: Worksheet
    ' s: Heading of column
    ' i: Row where the header is supposed to be
    
    Dim j As Integer
    
    j = 1
    Do Until o.Cells(i, j) = s Or o.Cells(i, j) = ""
        j = j + 1
    Loop
    If o.Cells(i, j) = "" Then
        MsgBox "Column '" & s & "' not found, Check data or macro and adjust if necessary."
        End
    End If
    FindColumn = j
End Function

Function FindStringInColumn(o, s, t, i)
    ' Goal: Find and Address column. However, compared to "FindColumn" we do search for substrings because of how the content data files are structured.
    
    ' o: Worksheet
    ' s: Substring of header
    ' t: Key for identification in substring; Search for Product-No. ==> will be "."; search for "BD-" ==> "-"
    ' i: Row where the header is supposed to be
    
    Dim j As Integer
    
    j = 1
    Do Until o.Cells(i, j) = "" Or Left(o.Cells(i, j), InStr(o.Cells(i, j), t)) = s
        j = j + 1
    Loop
    If o.Cells(i, j) = "" Then
        MsgBox "Spalte mit Eintrag '" & s & "' nicht gefunden. Programm bricht ab."
        End
    End If
    FindStringInColumn = j
End Function

Function FindSubstringInColumn(o, s, i)
    ' Goal: Before this macro existed, many product data sheets were made manually. As a consequence, some headers are written wrong or contain " ".
    ' So check if substring is in string.
    
    Dim j As Integer
    
    j = 1
    Do Until o.Cells(i, j) = "" Or InStr(o.Cells(i, j), s) > 0
        j = j + 1
    Loop
    If o.Cells(i, j) = "" Then
        MsgBox "Spalte mit Eintrag '" & s & "' nicht gefunden. Datenblatt genau auf Sonderzeichen sowie Leerzeichen prüfen."
        End
    End If
    FindSubstringInColumn = j
End Function

Sub Rand(o, no)
    ' Goal: Create cell borders
    
    Dim b(5) As Variant
    
    b(0) = xlEdgeTop
    b(1) = xlEdgeBottom
    b(2) = xlEdgeLeft
    b(3) = xlEdgeRight
    b(4) = xlInsideVertical
    b(5) = xlInsideHorizontal
    
    With o.Borders(b(no))
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

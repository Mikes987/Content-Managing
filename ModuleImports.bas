Attribute VB_Name = "ModuleImports"
Option Private Module
Sub ImportValues(p, q, si)
    ' Goal: Insert default values and their IDs into the specific datasheets if we handle iPIM lables
    
    ' Variables
    Dim wb As Workbook
    Dim o As Object
    Dim s, t1, t2 As String
    Dim b, b1 As Boolean
    Dim i, k As Integer
    Dim pid, pwe, pli As Integer
    Dim j, x As Long
    
    ' Open and address data file
    Call LoadFile(wb, o, si, "Legend")
    p.Activate
    
    ' In this sheet, we need 3 Columns
    ' pid: Identifier
    ' pwe: Default Values
    ' pli: Lookup-Identifier
    pid = FindColumn(o, "Identifier", 1)
    pwe = FindColumn(o, "Default Values", 1)
    pli = FindColumn(o, "Lookup-Identifier", 1)
    
    ' We can use the identifier of the attribute names which are also within this sheet as a reference which default values need to be copied.
    ' The IDs will be copied into the third data sheet but on the same position as their default value.
    i = 2
    Do Until p.Cells(2, i) = ""
        s = p.Cells(2, i)
        j = 2
        b = False
        ' Go through each row of the data file and look for a match
        Do Until b = True Or o.Cells(j, 1) = ""
            If s = o.Cells(j, pid) Then
                ' If a match is made, copy default values into sheet 2 and ID into sheet 3 on the same position
                b = True
                k = 6
                Do While s = o.Cells(j, pid)
                    p.Cells(k, i) = o.Cells(j, pwe)
                    q.Cells(k, i) = o.Cells(j, pli)
                    k = k + 1
                    o.Rows(CStr(j) & ":" & CStr(j)).Delete
                Loop
            Else
                j = j + 1
            End If
        Loop
        i = i + 1
    Loop
    
    ' Done close data file with default values
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Sub

Sub PrepareImport(o)
    ' Ziel: After receiving the product data sheet with information by the supplier and transforming default values into their IDs, handle the sheet for automatic import
    
    ' Variables
    Dim i, j, k As Integer
    Dim i1, i2 As Integer
    Dim x As Integer
    Dim s, t, t1, u As String
    Dim b As Boolean
    
    ' Unhide all cells
    o.Cells.EntireRow.Hidden = False
    
    ' Count rows
    i = 6
    Do Until o.Cells(i + 1, 1) = ""
        i = i + 1
    Loop
    
    ' We need the first row with no content
    i1 = i + 1
    
    ' EAN is not needed any longer, the entire column can be deleted. I pretend as if I don't know the column and look for it. Then the column will be deleted.
    j = FindColumn(o, "EAN", 6)
    s = Split(o.Cells(1, j).Address, "$")(1)
    o.Columns(s & ":" & s).Delete
    
    ' Product and article based attributes must be prositioned in a certain way, first for product, then article. For that, product numbers must be copied and inserted just below their
    ' last entry.
    j = FindColumn(o, "Product Number", 6)
    o.Range(o.Cells(7, j), o.Cells(i, j)).Copy Destination:=o.Cells(i + 1, j)
    
    ' Just in case, save the position of the last content.
    i2 = i1
    Do Until o.Cells(i2 + 1, j) = ""
        i2 = i2 + 1
    Loop
    
    ' Now move the article numbers in the positions below
    j = FindSubstringInColumn(o, "Article Number", 6)
    o.Range(o.Cells(7, j), o.Cells(i, j)).Cut Destination:=o.Cells(i1, j)
    
    ' If we have columns with multi default values, combine them with " | " as the delimiter
    j = j + 1
    
    Do Until o.Cells(6, j) = ""
        If o.Cells(5, j) = "Value, multi" Then
            For k = 1 To 6
                o.Cells(k, j).MergeCells = False
            Next
            ' Check if the supplier actually used multiple values
            For k = 7 To i
                b = True
                x = j
                s = ""
                Do Until x > j + 2 Or b = False Or o.Cells(5, j + 1) <> ""
                    If o.Cells(k, x) <> "" Then
                        If s <> "" Then s = s & " | "
                        s = s & o.Cells(k, x)
                        x = x + 1
                    Else
                        x = x + 1
                    End If
                Loop
                If o.Cells(5, j + 1) = "" Then o.Cells(k, j) = s
            Next
            ' Remove remaining columns, usually there are 3 for multiple default values, for import, we need 1.
            t = Split(o.Cells(1, j + 1).Address, "$")(1)
            Do While o.Cells(5, j + 1) = ""
                o.Columns(t & ":" & t).Delete
            Loop
        ElseIf o.Cells(5, j) = "Value" And o.Cells(6, j - 1) = "Percentage" Then
            u = o.Cells(4, j)
            ' Some Products require to give their composition in percentage. That will also be combined with " | "
            For k = 7 To i
                b = True
                x = j
                s = ""
                Do Until b = False Or o.Cells(4, x) <> u
                    If o.Cells(k, x) <> "" Then
                        If s <> "" Then s = s & " | "
                        s = s & CStr(o.Cells(k, x - 1)) & "# " & o.Cells(k, x)
                        x = x + 2
                    Else
                        b = False
                    End If
                Loop
                If s <> "" Then o.Cells(k, j) = s
            Next
            t = Split(o.Cells(1, j + 2).Address, "$")(1)
            t1 = Split(o.Cells(1, j + 1).Address, "$")(1)
            Do While o.Cells(4, j + 2) = u
                o.Columns(t1 & ":" & t).Delete
            Loop
            t = Split(o.Cells(1, j - 1).Address, "$")(1)
            o.Columns(t & ":" & t).Delete
            j = j - 1
        End If
        ' If the content is article oriented, move all content downwards
        If InStr(o.Cells(4, j), "dim") > 0 Or InStr(o.Cells(4, j), "_Artikel") > 0 Or o.Cells(1, j) = "A" Or o.Cells(1, j) = "Article" Or o.Cells(1, j) = "V" Or o.Cells(4, j) = "PrimaryColor" Then
            o.Range(o.Cells(7, j), o.Cells(i, j)).Cut Destination:=o.Cells(i1, j)
        End If
        
        j = j + 1
    Loop
    
    ' Delete rows that are not needed: 1,2,3,5,6
    o.Rows("5:6").Delete
    o.Rows("1:3").Delete
End Sub

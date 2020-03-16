Attribute VB_Name = "ModuleValueIDs"
Option Private Module

Sub DefaultValuesIDs(p, q, si)
    ' Copy Default Value IDs, for PBK Labels only
    
    ' Why is this necessary?
    ' iPIM stores every single default Value, but we only need a certain percentage and specific ones.
    ' The primary data file stores the specific ones. However, we need the File with default values to get their associated IDs
    
    ' Variables
    Dim wb As Workbook
    Dim r As Object
    Dim b, b1 As Boolean
    Dim s As String
    Dim a1, a2, a3 As Integer
    Dim i, j As Integer
    Dim x, y As Long
    
    ' Open and reference File
    Call LoadFile(wb, r, si, "Legend")
    
    ' As in iPIM Labels, we need 3 Columns
    ' a1: Identifier
    ' a2: Default Values
    ' a3: Lookup-Identifier
    a1 = FindColumn(r, "Identifier", 1)
    a2 = FindColumn(r, "Wertemenge", 1)
    a3 = FindColumn(r, "Lookup-Identifier", 1)
    
    q.Activate
    
    ' We match the attribute IDs in this file and the product data sheet, default values in row 2.
    i = 2
    Do Until p.Cells(2, i) = ""
        ' We have a characteristic sort. Column A is the primary one, so we look for a match, go down each line until there is no match anymore
        x = 2
        s = p.Cells(2, i)
        b1 = False
        Do Until r.Cells(x, a1) = "" Or b1 = True
            If s = r.Cells(x, a1) Then
                b1 = True
                ' We then match the default values and copy the IDs only.
                j = 6
                Do Until p.Cells(j, i) = ""
                    ' Because of the characteristic sort, we always begin in line x or its reference respectively.
                    y = x
                    b = False
                    Do Until s <> r.Cells(y, a1) Or r.Cells(y, 1) = "" Or b = True
                        If p.Cells(j, i) = r.Cells(y, a2) Then
                            b = True
                            q.Cells(j, i) = r.Cells(y, a3)
                        Else
                            y = y + 1
                        End If
                    Loop
                    j = j + 1
                Loop
            Else
                x = x + 1
            End If
        Loop
        i = i + 1
    Loop
    
    ' Close file.
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Sub

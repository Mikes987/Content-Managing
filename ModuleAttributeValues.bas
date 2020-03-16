Attribute VB_Name = "ModuleAttributeValues"
Option Private Module

Sub AttributeValues(o, p, q)
    ' Goal: Copy attribute names from product data sheet into the other sheets if they shall contain default values
    
    ' Variables
    Dim i, j As Integer
    
    i = 1
    j = 2
    
    ' Not necessary but active sheet with default values to see if the code worked.
    p.Activate
    
    Do Until o.Cells(6, i) = ""
        If o.Cells(5, i) = "Value, single" Or o.Cells(5, i) = "Value, multi" Then
            ' unit
            p.Cells(1, j) = o.Cells(3, i)
            q.Cells(1, j) = o.Cells(3, i)
            ' ID
            p.Cells(2, j) = o.Cells(4, i)
            q.Cells(2, j) = o.Cells(4, i)
            ' Data Type
            p.Cells(3, j) = o.Cells(5, i)
            q.Cells(3, j) = o.Cells(5, i)
            ' Attribute-name
            p.Cells(4, j) = o.Cells(6, i)
            q.Cells(4, j) = o.Cells(6, i)
            j = j + 1
        End If
        i = i + 1
    Loop
End Sub


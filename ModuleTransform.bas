Attribute VB_Name = "ModuleTransform"
Sub Transform(wb, o, p, q)
    ' Ziel: Transform default values into their IDs
    
    ' o: Product Data Sheet
    ' p: Default Values
    ' q: Default Values IDs
    
    Dim wb1 As Workbook
    Dim r As Object
    Dim i, j, k, l, x As Integer
    Dim s, t, u As String
    Dim b As Boolean
    
    ' First we need the number of all products within the object o. I consider that for every product there is a EAN, thus I go through every row of column A
    i = 6
    Do Until o.Cells(i, 1) = ""
        i = i + 1
    Loop
    If i = 6 Then
        MsgBox "No EAN found"
        End
    Else
        i = i - 1
    End If
    
    ' In order not to overwrite any information, a copy of object o and insert the IDs into the Copy
    s = "Product Data Sheets with IDs"
    For Each Worksheet In wb.Sheets
        If Worksheet.Name = s Then
            b = True
            Set r = wb.Sheets(s)
            Exit For
        End If
    Next
    If b = False Then
        o.Copy after:=o
        Set r = wb.ActiveSheet
        r.Name = s
    End If
    
    ' Within a Loop, go through all columns in row 6 and check if this column contains attributes with default values.
    ' If yes, then take a look at the attribute ID and look for its counterpart within sheet p. Then call function "DoTransform"
    j = 1
    Do Until o.Cells(6, j) = ""
        If o.Cells(5, j) = "Value, single" Then
            s = o.Cells(4, j)
            k = FindColumn(p, s, 2)
            Call DoTransform(o, p, q, r, s, t, u, i, j, k, l)
            j = j + 1
        ElseIf o.Cells(5, j) = "Value, multi" Then
            ' Default values with multiple Choices always contain 3 columns, so we go through the function "DoTransform" 3 times.
            s = o.Cells(4, j)
            k = FindColumn(p, s, 2)
            For x = 1 To 3
                Call DoTransform(o, p, q, r, s, t, u, i, j, k, l)
                j = j + 1
            Next
        Else
            j = j + 1
        End If
    Loop
End Sub

Sub DoTransform(o, p, q, r, s, t, u, i, j, k, l)
    ' Remember attribute ID and look for its counterpart in sheet "Default Values"
    ' The IDs of the default values are on the same position on on the data "Default Values IDs"
    ' Use For Loop to check if the supplier has inserted any default values
    For l = 7 To i
        ' Only do something  if there is any content in a cell
        If o.Cells(l, j) <> "" Then
            ' Excel or VBA respectively do have some issues recognizing and interpreting numbers or their datatype correctly. Even if considered to be set as strings in the cell,
            ' they are sometimes interpreted as integers or float. As a consequence, all default values are transformed into strings first and then checked for a match
            t = CStr(o.Cells(l, j))
            m = 6
            b = False
            Do Until p.Cells(m, k) = "" Or b = True
                u = CStr(p.Cells(m, k))
                If t = u Then
                    b = True
                    r.Cells(l, j) = q.Cells(m, k)
                    ' The database of primary is newer then iPIM, so some default values do not exist in iPIM and therefore, there is no ID, either.
                    ' If there is a default value but no associated ID, the interior of the cell will be marked yellow, so the operator knows immediately
                    ' that a value is missing.
                    If r.Cells(l, j) = "" Then
                        r.Cells(l, j) = t
                        r.Cells(l, j).Interior.Color = 65535
                    End If
                Else
                    m = m + 1
                End If
            Loop
            ' Sometimes, suppliers do not use any default value because they do not match their requirements for descriptions and thus, they manually type certain information.
            ' In this case, there will be no match with any default value and the interior of the cell will turn red.
            If b = False Then r.Cells(l, j).Interior.Color = 255
        End If
    Next
End Sub

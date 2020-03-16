Attribute VB_Name = "ModuleDropDown"
Option Private Module
Sub DropDown(o, p, q)
    ' Goal: Create a DropDown Menue in the columns that contain default values
    
    Dim i, j, k, m, x, y, z As Integer
    Dim s, t As String
    
    o.Activate
    
    ' As always, we go through row 6
    i = 1
    Do Until o.Cells(6, i) = ""
        ' First look for Selling Point 5 to create range and execute autofit.
        If o.Cells(6, i) = "Selling Point 5" Then
            x = i
        End If
        ' Check if single or multi values
        If o.Cells(5, i) = "Wertemenge, einfach" Or o.Cells(5, i) = "Wertemenge, mehrfach" Then
            ' Match the attribute IDs in product data sheet and sheet with default values
            j = 2
            Do Until o.Cells(4, i) = p.Cells(2, j)
                j = j + 1
            Loop
            ' Suche die Position des letzten Eintrages, das Zeilenende wird benötigt.
            k = 5
            Do Until p.Cells(k + 1, j) = ""
                k = k + 1
            Loop
            ' Certain executions if we have multi values
            If o.Cells(5, i) = "Value, multi" Then
                With o.Cells(3, i)
                    .Value = "Multiple Choices"
                    .WrapText = True
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                End With
                ' Create 2 extra columns
                For z = 1 To 2
                    o.Cells(1, i + 1).EntireColumn.Insert , copyorigin:=xlFormatFromLeftOrAbove
                Next
                o.Range(o.Cells(1, i), o.Cells(1, i + 2)).Merge
                For m = 3 To 6
                    o.Range(o.Cells(m, i), o.Cells(m, i + 2)).Merge
                    o.Cells(m, i).HorizontalAlignment = xlCenter
                Next
                For y = 0 To 3
                    Call Rand(o.Range(o.Cells(3, i), o.Cells(3, i + 2)), y)
                Next
                Set r = o.Range(o.Cells(7, i), o.Cells(307, i + 2))
                i = i + 2
            Else
                Set r = o.Range(o.Cells(7, i), o.Cells(307, i))
            End If
            s = Split(p.Cells(1, j).Address, "$")(1)
            ' Create the Dropdown here
            t = "=" & p.Name & "!$" & s & "$" & CStr(5) & ":$" & s & "$" & CStr(k)
            With r.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=t
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = False
            End With
        ' iPIM Labals contains boolean-like attributes
        ElseIf o.Cells(5, i) = "Boolean" Then
            Set r = o.Range(o.Cells(7, i), o.Cells(307, i))
            With r.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = False
            End With
        End If
        i = i + 1
    Loop
    
    ' Colored markings and borders
    i = i - 1
    Set r = o.Range(o.Cells(5, 1), o.Cells(6, i))
    For j = 0 To 4
        Call Rand(r, j)
    Next
    r.Interior.Color = 15921906
    o.Range(o.Cells(6, x), o.Cells(6, i)).Columns.EntireColumn.AutoFit
    o.Rows("1:1").EntireRow.Hidden = True
    o.Rows("4:4").EntireRow.Hidden = True
    
    ' Hide sheets with default values and their IDs
    p.Visible = False
    q.Visible = False
End Sub

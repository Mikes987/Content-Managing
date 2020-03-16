Attribute VB_Name = "ModuleCopyPBKValues"
Option Private Module

Sub DefaultValues(o, p, si)
    ' Goal:
    ' 1. Copy Default values into the product sheet, default values, here for PBK Labels
    ' 2. Copy their associated IDs into the specific sheet
    
    ' Creating 2 objects for addressing 2 data sheets of the primary file
    ' q: default values
    ' r: Info-Sheet.
    ' Info for info-Sheet: It does not have a constant name but refers to its label. However, the index of its sheet appears to always be 1.
    
    ' Variables
    Dim wb1, wb2 As Workbook
    Dim q, r As Object
    Dim s As String
    Dim b As Boolean
    Dim i, j, k, x, y As String
    
    ' Open and addressing the sheets
    Call LoadFile(wb1, q, si, "Vorgabewerte")
    Call LoadFile(wb2, r, si, 1)
    
    ' First, in r, we need the row where the specific attributes and categorys are listed
    ' It begins, if we go through each line in column F until we have a match with "Merkmal" (Characteristic in English)
    x = 1
    Do Until r.Cells(x, 6) = "Merkmal" Or r.Cells(x, 6) = ""
        x = x + 1
    Loop
    If r.Cells(x, 2) = "" Then
        MsgBox "No Match with 'Merkmal', please check file."
        End
    End If
    x = x + 1
    
    ' Going through each column in row 6 in the product data sheet and check for a match.
    ' Check if it is product based or article based.
    i = 1
    Do Until o.Cells(6, i) = ""
        y = x
        b = False
        Do Until b = True Or (r.Cells(y, 6) = "" And r.Cells(y, 6).MergeCells = False)
            If o.Cells(6, i) = r.Cells(y, 6) Then
                b = True
                s = r.Cells(y, 2)
                If s = "A" Or s = "V" Then
                    s = "Artikel"
                ElseIf s = "P" Then
                    s = "Produkt"
                End If
                o.Cells(1, i) = s
            Else
                y = y + 1
            End If
        Loop
        i = i + 1
    Loop
    
    ' We need a triple loop for copying default values in q
    ' Loop 1: Go through the header (Attributes) in the sheet "Default Values"
    ' Loop 2: Go through row 1 within the default values of the primary data file and find match
    ' Loop 3: Copy all default values
    
    p.Activate
    
    ' Loop 1
    i = 2
    Do Until p.Cells(4, i) = ""
        j = 1
        b = False
        ' Loop 2
        Do Until b = True Or q.Cells(1, j) = ""
            If p.Cells(4, i) = q.Cells(1, j) Then
                b = True
                k = 2
                l = 6
                ' Loop 3
                Do Until q.Cells(k, j) = ""
                    p.Cells(l, i) = q.Cells(k, j)
                    k = k + 1
                    l = l + 1
                Loop
            Else
                j = j + 1
            End If
        Loop
        i = i + 1
    Loop
    
    ' Primary data file can be closed
    Application.DisplayAlerts = False
    wb1.Close
    Application.DisplayAlerts = True
End Sub

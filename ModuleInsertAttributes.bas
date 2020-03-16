Attribute VB_Name = "ModuleInsertAttributes"
Option Private Module
Sub InsertAttributes(o, sa)
    ' Goal: Insert attributes as headers from attribute file into the rows 3, 4, 5 and 6
    
    ' Variables
    Dim wb As Workbook
    Dim q As Object
    Dim s, t1, t2 As String
    Dim b As Boolean
    Dim aid, atr, att, ate, apf As Integer
    Dim i, oi, ai As Integer
    
    ' Open and set reference to  attribute data file
    Call LoadFile(wb, q, sa, 1)
    
    ' All in all, 5 columns have to be addressed:
    ' aid: Attribut-ID
    ' atr: Attribut
    ' att: Attributtyp
    ' ate: Attribut-Einheit
    ' apf: Pflichttyp
    
    aid = FindColumn(q, "Attribute-ID", 1)
    atr = FindColumn(q, "Attribute", 1)
    att = FindColumn(q, "Attributtype", 1)
    ate = FindColumn(q, "Attribute-Unit", 1)
    apf = FindColumn(q, "Mandatory", 1)
    
    ' First we need to know the last columns in the product data sheet to know where to insert the attributes
    oi = 1
    Do Until o.Cells(6, oi) = ""
        oi = oi + 1
    Loop
    
    ' Every Attribute does contain further information in brackets. Delete the entire further information
    ' Then, copy attributes
    i = 2
    Do Until q.Cells(i, 1) = ""
        q.Cells(i, atr) = Left(q.Cells(i, atr), InStr(q.Cells(i, atr), "(") - 2)
        o.Cells(6, oi) = q.Cells(i, atr)
        If q.Cells(i, apf) = "Mandatory" Then o.Cells(6, oi).Font.Color = -16776961
        o.Cells(5, oi) = q.Cells(i, att)
        o.Cells(4, oi) = q.Cells(i, aid)
        With o.Cells(3, oi)
            .Value = q.Cells(i, ate)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        oi = oi + 1
        i = i + 1
    Loop
    
    ' Data file with attributes can be closed. The data file has been changed, so we set "Display-Alerts" off to skip messenger box for saving
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Sub

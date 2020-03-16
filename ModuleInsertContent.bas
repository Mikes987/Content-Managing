Attribute VB_Name = "ModuleInsertContent"
Option Private Module
Sub InsertContent(o, sc, bp, t)
    ' Goal: Copy EAN, product no. etc into the product data sheet if selected
    
    ' Variables
    Dim wb As Workbook
    Dim p As Object
    Dim b As Boolean
    Dim s, t1, t2 As String
    Dim ea1, ea2 As Integer
    Dim pr1, pr2 As Integer
    Dim ar1, ar2 As Integer
    Dim li1, li2 As Integer
    Dim ma1, ma2 As Integer
    Dim i1, i2, j, z As Integer
    
    ' Open and address content file
    Call LoadFile(wb, p, sc, "Content Query")
    
    ' We need 5 Columns on both sheets
    ' ea1, ea2: EAN
    ' pr1, pr2: Product number
    ' ar1, ar2: Article number
    ' li1, li2: Supplier number
    ' ma1, ma2: Brand
    
    ' Again, due to the strange arrangement of the content of the headers, we sometimes need to look for substrings
    
    ' EAN
    ea1 = FindColumn(p, "EAN", 3)
    ea2 = FindColumn(o, "EAN", 6)
    
    ' Product number
    pr1 = FindStringInColumn(p, "Product-No.", ".", 3)
    pr2 = FindColumn(o, "Product number", 6)
    
    ' Article number
    ar1 = FindStringInColumn(p, "BD-", "-", 3)
    ar2 = FindColumn(o, "Article number", 6)
    
    ' Supplier number
    li1 = FindStringInColumn(p, "Supplier number/", "/", 3)
    li2 = FindColumn(o, "Supp.-Art.-Description", 6)
    
    ' Brand
    ma1 = FindColumn(p, "BRAND", 3)
    ma2 = FindColumn(o, "Brand", 6)
    
    ' All important columns are addressed, with the help of the boolan BP, we know if a iPIM or PBK Label is chosen
    If bp = False Then
        s = "exact location in iPIM"
    Else
        s = "PBK"
    End If
    
    ' Look for the header in content file
    j = FindColumn(p, s, 3)
    
    ' Copy. For that we need indices for the rows in both data sheets
    i1 = 4
    i2 = 7
    Do Until p.Cells(i1, j) = ""
        If p.Cells(i1, j) = t Then
            o.Cells(i2, ea2) = p.Cells(i1, ea1)
            o.Cells(i2, pr2) = p.Cells(i1, pr1)
            o.Cells(i2, ar2) = p.Cells(i1, ar1)
            o.Cells(i2, li2) = p.Cells(i1, li1)
            o.Cells(i2, ma2) = p.Cells(i1, ma1)
            With o.Range(o.Cells(i2, ea2), o.Cells(i2, li2))
                .HorizontalAlignment = xlLeft
                .NumberFormat = "0"
            End With
            i2 = i2 + 1
        End If
        i1 = i1 + 1
    Loop
End Sub

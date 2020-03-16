Attribute VB_Name = "ModuleBuildTable"
Option Private Module
Public Sub BuildTable(o, p, q)
    ' Goal: Initializing and creating first content of table
    
    ' Variables
    Dim r As Object
    Dim z(1) As Variant
    Dim a(2, 15), b(5) As String
    Dim i, j, a1, a2, e1, e2 As Integer
    Dim s1, s2 As String
    Dim t1, t2, t3 As String
    
    'Set o = ThisWorkbook.ActiveSheet
    'o.Name = "Product Data sheet"
    
    o.Rows("2:2").RowHeight = 35
    o.Rows("3:3").RowHeight = 40
    o.Rows("4:4").Font.Size = 10
    o.Cells.ColumnWidth = 28
    
    ' Datatypes
    t1 = "String"
    t2 = "Value, single"
    t3 = "Value, multi"
    
    ' 2d-Array with standard headers
    a(0, 0) = "ARTICLEEAN"
    a(0, 1) = "IPIM_PRODUCT_NUMBER"
    a(0, 2) = "IPIM_ARTICLE_NUMBER"
    a(0, 3) = "SUPP_ART_DESCRIPTION"
    a(0, 4) = "Brand"
    a(0, 5) = "Producttype"
    a(0, 6) = "ProductName"
    a(0, 7) = "Addition_Short_name"
    a(0, 8) = "SpecialFeatures_Str_Compliance"
    a(0, 9) = "Set-Type"
    a(0, 10) = "SerialName"
    a(0, 11) = "Selling Point 1"
    a(0, 12) = "Selling Point 2"
    a(0, 13) = "Selling Point 3"
    a(0, 14) = "Selling Point 4"
    a(0, 15) = "Selling Point 5"
    
    a(1, 0) = ""
    a(1, 1) = "BD"
    a(1, 2) = "BD"
    a(1, 3) = ""
    a(1, 4) = t1
    a(1, 5) = t2
    a(1, 6) = t1
    a(1, 7) = t1
    a(1, 8) = t1
    a(1, 9) = t2
    a(1, 10) = t1
    a(1, 11) = t1
    a(1, 12) = t1
    a(1, 13) = t1
    a(1, 14) = t1
    a(1, 15) = t1
    
    a(2, 0) = "EAN"
    a(2, 1) = "Product Number"
    a(2, 2) = "Article Number"
    a(2, 3) = "Supp.-Art.-Description"
    a(2, 4) = "Brand"
    a(2, 5) = "Producttype"
    a(2, 6) = "Product-Name"
    a(2, 7) = "Addition Short Name"
    a(2, 8) = "Special Features"
    a(2, 9) = "Set-Type"
    a(2, 10) = "Serienname"
    a(2, 11) = "Selling Point 1"
    a(2, 12) = "Selling Point 2"
    a(2, 13) = "Selling Point 3"
    a(2, 14) = "Selling Point 4"
    a(2, 15) = "Selling Point 5"
    
    ' Create Initial standard table
    For i = 0 To UBound(a, 1)
        For j = 0 To UBound(a, 2)
            o.Cells(i + 4, j + 1) = a(i, j)
            ' A superior header has to be created above Brand until Addition Short Name, so save their column positions
            If o.Cells(i + 4, j + 1) = "Brand" Then e1 = j + 1
            If o.Cells(i + 4, j + 1) = "Addition Short Name" Then e2 = j + 1
            ' Likwise with Selling points
            If o.Cells(i + 4, j + 1) = "Selling Point 1" Then a1 = j + 1
            If o.Cells(i + 4, j + 1) = "Selling Point 5" Then a2 = j + 1
            ' Some headers are mandatory headers, make the font red.
            If o.Cells(6, j + 1) = "Selling Point 1" Or o.Cells(6, j + 1) = "Selling Point 2" Or o.Cells(6, j + 1) = "Selling Point 3" Or o.Cells(6, j + 1) = "Produkttyp" Then o.Cells(6, j + 1).Font.Color = -16776961
        Next
    Next
    
    ' Row 6 contains all attributes as headers and the font shall be bold.
    o.Rows("6:6").Font.Bold = True
    
    ' Create the content of the superior headers
    s1 = "Content leads to online title and apperance of product!"
    s2 = "(valid for all variants of the product)"
    
    ' String s1 shall be bold
    ' String s2 shall be italic
    o.Cells(2, e1) = s1 & vbNewLine & s2
    i = InStr(o.Cells(2, e1), "!")
    With o.Cells(2, e1)
        .Characters(Start:=1, Length:=i).Font.FontStyle = "Bold"
        .Characters(Start:=1, Length:=i).Font.Size = 12
        .Characters(Start:=i + 1).Font.FontStyle = "Italic"
    End With
    ' Create borders
    Set r = o.Range(o.Cells(2, e1), o.Cells(3, e2))
    r.Interior.Color = 12379352
    For i = 0 To UBound(b)
        Call Rand(r, i)
    Next
    ' Additional content
    o.Cells(3, e2 - 1) = "ONLY name of the product!"
    o.Cells(3, e2) = "E.g. measurements, specific features, material"
    With o.Range(o.Cells(3, e2 - 1), o.Cells(3, e2))
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    With o.Range(o.Cells(2, e1), o.Cells(2, e2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Merge
    End With
    
    ' Now do the same for the selling points
    s1 = "Unique Selling Points that show how the products differs from competitors."
    s2 = "Short and concise (only 55 characters per selling point!)"
    
    With o.Cells(2, a1)
        .Value = s1
        .Font.Bold = True
        .Font.Size = 12
    End With
    With o.Cells(3, a1)
        .Value = s2
        .Font.Italic = True
    End With
    
    Set r = o.Range(o.Cells(2, a1), o.Cells(3, a2))
    With r
        .Interior.Color = 15849925
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    For i = 0 To 3
        Call Rand(r, i)
    Next
    o.Range(o.Cells(2, a1), o.Cells(2, a2)).Merge
    o.Range(o.Cells(3, a1), o.Cells(3, a2)).Merge
    
    ' Now concentrate on the second and third sheet for default values. Create legend to show how the position should be.
    With p
        .Cells(1, 1) = "Attribut-Einheit"
        .Cells(2, 1) = "Attribut-ID"
        .Cells(3, 1) = "Attributtyp"
        .Cells(4, 1) = "Attribut"
        .Cells(5, 1) = "Attributswerte"
        .Range("A1:A5").Font.Bold = True
        .Cells(1, 1).EntireColumn.AutoFit
        .Rows("4:4").Font.Bold = True
    End With
    With q
        .Cells(1, 1) = "Attribut-Einheit"
        .Cells(2, 1) = "Attribut-ID"
        .Cells(3, 1) = "Attributtyp"
        .Cells(4, 1) = "Attribut"
        .Cells(5, 1) = "Attributswerte"
        .Range("A1:A5").Font.Bold = True
        .Cells(1, 1).EntireColumn.AutoFit
        .Rows("4:4").Font.Bold = True
    End With
End Sub

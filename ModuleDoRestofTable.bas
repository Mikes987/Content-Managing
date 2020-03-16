Attribute VB_Name = "ModuleDoRestofTable"
Option Private Module
Sub RemainingData(o)
    ' Goal: Insert remaining data as attributes that do not appear in attributes data files
    
    ' Variables
    Dim r As Object
    Dim i, j, k, l, m As Integer
    Dim a(2, 8) As String
    
    
    ' 2d-Array
    a(0, 0) = "PrimaryColor"
    a(0, 1) = "SupplierComment"
    a(0, 2) = "SEOMarketingtext"
    a(0, 3) = "XSELL"
    a(0, 4) = "ADSELL"
    a(0, 5) = "SERIE"
    a(0, 6) = "VARIANT"
    a(0, 7) = "SET"
    a(0, 8) = "SETPART"

    a(1, 0) = "Value, single"
    a(1, 1) = "Item related"
    a(1, 2) = "Item related"
    a(1, 3) = "String"
    a(1, 4) = "String"
    a(1, 5) = "String"
    a(1, 6) = "String"
    a(1, 7) = "String"
    a(1, 8) = "String"

    a(2, 0) = "Primary Color"
    a(2, 1) = "Supplier Comment"
    a(2, 2) = "Seo Marketing Text"
    a(2, 3) = "Fits with that"
    a(2, 4) = "Equipment"
    a(2, 5) = "Serial"
    a(2, 6) = "Variant"
    a(2, 7) = "Set"
    a(2, 8) = "Set Component"
    
    ' First look for last entry in row 6
    i = 1
    Do Until o.Cells(6, i) = ""
        i = i + 1
    Loop
    
    ' Then insert content of array
    For j = 0 To UBound(a, 2)
        l = 4
        For k = 0 To UBound(a, 1)
            o.Cells(l, i) = a(k, j)
            l = l + 1
        Next
        ' Primary Color is mandatory, change font to red
        If o.Cells(6, i) = "Primary Color" Then o.Cells(6, i).Font.Color = -16776961
        ' Marketing SEO needs further Header, save position
        If o.Cells(6, i) = "SEO Marketing Text" Then m = i
        i = i + 1
    Next
    
    ' Create superior Header for SEO
    s = Split(o.Cells(1, m).Address, "$")(1)
    o.Columns(s & ":" & s).ColumnWidth = 55
    With o.Cells(3, m)
        .Value = "Valid for all variants of product."
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = 15597051
        .WrapText = True
    End With
    
    ' Create another superior Header, here in line 2 and 3
    m = m + 1
    With o.Cells(2, m)
        .Value = "Product relationships if there are any."
        .Font.Bold = True
        .Font.Size = 12
    End With
    With o.Cells(3, m)
        .Value = "Please insert BD Article number with comma as delimiter"
        .Font.Italic = True
    End With
    
    i = i - 1
    Set r = o.Range(o.Cells(2, m), o.Cells(3, i))
    With r
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = 15849925
    End With
    For j = 0 To 3
        Call Rand(r, j)
    Next
    
    o.Range(o.Cells(2, m), o.Cells(2, i)).Merge
    o.Range(o.Cells(3, m), o.Cells(3, i)).Merge
End Sub

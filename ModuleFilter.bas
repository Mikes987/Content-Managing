Attribute VB_Name = "ModuleFilter"
Option Private Module

Sub Filter(o)
    ' Goal: Filter attributes from the product data sheet that are not necessary
    
    ' Variables
    Dim fi(13) As String
    Dim i, j As Integer
    Dim s As String
    Dim b As Boolean
    
    ' Filterarray
    fi(0) = "PBK-Valueset"
    fi(1) = "Care Instructions"
    fi(2) = "Product Labeling"
    fi(3) = "Relevance Battery Law"
    fi(4) = "Relevance CE-Obligation"
    fi(5) = "Relevance Guideline"
    fi(6) = "Languages CE-Declaration of Conformity"
    fi(7) = "Sprachen Warnings"
    fi(8) = "Catalog text"
    fi(9) = "Special Features"
    fi(10) = "Languages on the Product"
    fi(11) = "CE-Marking on the Product"
    fi(12) = "Manufacturer Address"
    fi(13) = "Marketing text SEO"
    
    ' We check the attributes at row 6. Delete the entire column if there is a match. The attributes begin after the selling points, so
    ' we first look for them and begin filtering as soon as we pass them.
    i = 1
    Do Until o.Cells(6, i) = "Selling Point 5"
        i = i + 1
    Loop
    i = i + 1
    
    ' Now we can filter.
    Do Until o.Cells(6, i) = ""
        j = 0
        b = False
        Do Until b = True Or j > UBound(fi)
            If o.Cells(6, i) = fi(j) Then
                b = True
                s = Split(o.Cells(1, i).Address, "$")(1)
                o.Columns(s & ":" & s).Delete
            Else
                j = j + 1
            End If
        Loop
        If b = False Then i = i + 1
    Loop
End Sub

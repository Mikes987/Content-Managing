VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormProduct 
   Caption         =   "Insert Product Number"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490.001
   OleObjectBlob   =   "UserFormProduct.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Apply_Click()
    ' Check first if Data files have been inserted.
    If ContentAddress.Caption = "" Or ContentAddress.Caption = "False" Then
        MsgBox "Content File is missing"
        Exit Sub
    End If
    If ProductNoAddress = "" Or ProductNoAddress = "False" Then
        MsgBox "Data file with product numbers is missing"
        Exit Sub
    End If
    
    ' Variablen
    Dim wb1, wb2 As Workbook
    Dim o, p As Object
    Dim b As Boolean
    Dim s, t As String
    Dim i1, i2 As Integer
    Dim ar1, ar2 As Integer
    Dim pr1, pr2 As Integer
    Dim u As Variant
    
    ' Save addresses
    s = ContentAddress.Caption
    t = ProductNoAddress.Caption
    
    ' All information from Userform is stored, userform can be closed.
    Unload Me
    
    ' Open, activate and reference data files and specific sheets
    Call LoadFile(wb1, o, s, "Content Query")
    Call LoadFile(wb2, p, t, 1)
    
    o.Activate
    
    ' The way of how the content of the cells are written is sometimes very strange, it is divided into two lines in the path but the cells show only one line.
    ' As a consequence sometimes I'll only look for substrings if no other way is possible.
    
    ' We need the positions of the article numbers as the method of comparison both sheets, then furthermore, the positions of
    ' the product numbers
    ' First look for the article numbers
    ar1 = FindStringInColumn(o, "BD-", "-", 3)
    ar2 = FindColumn(p, "Product-/Article number", 1)
    
    ' Now look for position of product numbers
    pr1 = FindStringInColumn(o, "Product-no.", ".", 3)
    pr2 = FindColumn(p, "Product", 1)
    
    ' We need to transform the product numers in the product sheet into integers. First we need to know the last row with data.
    ' Then, a range will be set.
    i2 = 1
    Do Until p.Cells(i2 + 1, 1) = ""
        i2 = i2 + 1
    Loop
    With p.Range(p.Cells(2, ar2), p.Cells(i2, pr2))
        .NumberFormat = "0"
        .Value = .Value
    End With
    
    ' Now we go through every row in the content data sheet first, beginning in row 4 and check if there exist a product number for this specific article number
    i1 = 4
    Do Until o.Cells(i1, ar1) = ""
        i2 = 2
        b = False
        Do Until p.Cells(i2, ar2) = "" Or b = True
            If o.Cells(i1, ar1) = p.Cells(i2, ar2) Then
                b = True
                o.Cells(i1, pr1) = p.Cells(i2, pr2)
            Else
                i2 = i2 + 1
            End If
        Loop
        i1 = i1 + 1
    Loop
    
    ' Close the file with the product numbers. In order not to be asked to save any file, the application "Display-Alert" will be deactivated first and reactivated then.
    Application.DisplayAlerts = False
    wb2.Close
    Application.DisplayAlerts = True
End Sub

Private Sub Back_Click()
    Unload Me
    UserFormStart.Show
End Sub

Private Sub ButtonContent_Click()
    ContentAddress.Caption = Application.GetOpenFilename("Excel-Worksheet (*.xlsx), *.xlsx")
End Sub

Private Sub ButtonProduct_Click()
    ProductNoAddress.Caption = Application.GetOpenFilename("Excel-Worksheet (*xlsx), *.xlsx")
End Sub

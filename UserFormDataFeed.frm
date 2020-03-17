VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDataFeed 
   Caption         =   "Insert information from datafeed file into product data sheet file"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250.001
   OleObjectBlob   =   "UserFormDataFeed.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormDataFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Insert content of datafeed from iPIM into empty product data sheet. The supplier can check his content, maybe after one month or one year and
    '       check if he wants something to be changed.
    
    ' Variables
    Dim wb1, wb2 As Workbook
    Dim o, p As Object
    Dim s, t, u As String
    
    s = DataFeedAddress.Caption
    t = UserFormDataFeed.ComboBoxData.Text
    u = ProductSheetAddress.Caption
    
    ' Check for inserted data files
    If s = "" Or s = "False" Then
        MsgBox "Datafeed file missing"
        Exit Sub
    ElseIf t = "" Then
        MsgBox "Product type has not been chosen"
        Exit Sub
    ElseIf u = "" Or u = "False" Then
        MsgBox "Product Data Sheet file missing"
        Exit Sub
    End If
    Unload Me
        
    ' Load workbooks if everything is ok so far.
    Call LoadFile(wb1, o, s, 1)
    Call LoadFile(wb2, p, u, "Product Data Sheet")
    
    ' Do insert via the module below
    Call Insert(o, p, t)
End Sub

Private Sub ButtonCancel_Click()
    Debug.Print ProductSheetAddress.Caption
    Unload Me
End Sub

Private Sub ButtonLoadData_Click()
    DataFeedAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub

Private Sub ButtonProductsheet_Click()
    ProductSheetAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub

Private Sub ButtonRead_Click()
    ' Goal: Read product types and transfer them into the combobox of the userform
    
    ' Variables
    Dim wb1 As Workbook
    Dim o As Object
    Dim s As String
    Dim i, j, k As Integer
    Dim a() As String
    Dim n As Variant
    Dim b As Boolean
    
    ' Check if a data file has been chosen
    s = DataFeedAddress.Caption
    If s = "" Or s = "False" Then
        MsgBox "Datafeed file missing"
        Exit Sub
    End If
    
    ' Open and address file
    Call LoadFile(wb1, o, s, 1)
    
    ' Labels are to be found within column "Einkaufskategorie"
    j = FindColumn(o, "Einkaufskategorie", 1)
    
    ' Insert labels into array
    ReDim a(0)
    a(0) = ""
    
    i = 3
    Do Until o.Cells(i, 1) = ""
        ' Only insert values if they don't already exist.
        s = o.Cells(i, j)
        b = False
        For k = 0 To UBound(a)
            If s = a(k) Then
                b = True
                Exit For
            End If
        Next
        If b = False Then
            ReDim Preserve a(k)
            a(k) = s
        End If
        i = i + 1
    Loop
    
    UserFormDataFeed.ComboBoxData.List = a
End Sub

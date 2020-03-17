VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormImport 
   Caption         =   "Prepare Product Data Sheet for Import"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   OleObjectBlob   =   "UserFormImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Prepare product data files for import into iPIM
    
    ' Variables
    Dim wb As Workbook
    Dim o, p As Object
    Dim s As String
    Dim b As Boolean
    
    ' Check if File has been loaded/chosen
    s = ProductSheetAddress.Caption
    If s = "" Or s = "False" Then
        MsgBox "Product data file missing"
        Exit Sub
    End If
    Unload Me
    
    ' Open and address data file
    ' We will do the following:
    ' First we will reference the regular product data sheet with default values
    ' However, if that file contains a sheet with title "Product Data Sheet with IDs" that was created in (2) then reference this one.
    Call LoadFile(wb, o, s, 1)
    For Each Worksheet In wb.Worksheets
        If Worksheet.Name = "Product Data Sheet with IDs" Then
            Set o = wb.Sheets("Product Data Sheet with IDs")
            Exit For
        End If
    Next
    
    ' We don't want to make any changes into the data file directly. Instead, we create a new file and copy the sheet into that file.
    Workbooks.Add
    Set p = ActiveWorkbook.ActiveSheet

    o.Cells.Copy Destination:=ActiveWorkbook.ActiveSheet.Cells(1, 1)
    Set o = ActiveWorkbook.ActiveSheet
    'wb.Close

    Call PrepareImport(o)
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonReadAddress_Click()
    ProductSheetAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *.xlsx")
End Sub

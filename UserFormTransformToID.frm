VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTransformToID 
   Caption         =   "Transform to ID"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   OleObjectBlob   =   "UserFormTransformToID.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormTransformToID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Transform default values set by suppliers for their products into their associated IDs if they exist
    
    ' Variables
    Dim wb1, wb2, wb3 As Workbook
    Dim o, p, q As Object
    Dim s As String
    Dim s1, s2, s3 As String
    
    ' Namen of all three sheets of the Product Data File
    s1 = "Product Data Sheet"
    s2 = "Default Values"
    s3 = "Default Values IDs"
    
    ' Laden der Datei
    s = ProductsheetAddress.Caption
    
    If s = "" Or s = "False" Then
        MsgBox "Product Data File missing"
        Exit Sub
    End If
    
    Unload Me
    
    Call LoadFile(wb1, o, s, s1)
    Call LoadFile(wb2, p, s, s2)
    Call LoadFile(wb3, q, s, s3)
    
    ' Calling a new module in which transformation is happening
    Call Transform(wb1, o, p, q)
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonLoad_Click()
    ProductsheetAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *.xlsx")
End Sub

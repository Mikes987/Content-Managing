VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormStart 
   Caption         =   "Create Product Data Sheet"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   OleObjectBlob   =   "UserFormStart.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonProductNumber_Click()
    Unload Me
    UserFormProduct.Show
End Sub

Private Sub ButtonProductsheet_Click()
    Unload Me
    UserFormProductsheet.Show
End Sub

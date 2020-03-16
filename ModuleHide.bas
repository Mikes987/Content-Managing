Attribute VB_Name = "ModuleHide"
Option Private Module

Sub HideColumnsAtEnd(o)
    ' Goal: Hide Columns for Product Relationships, for content management only
        
    ' Variables
    Dim i, j As Integer
    Dim s, t As String
    
    i = 1
    Do Until o.Cells(6, i) = "Fits with that"
        i = i + 1
    Loop
    j = 1
    Do Until o.Cells(6, j) = "Set Component"
        j = j + 1
    Loop
    
    s = Split(o.Cells(1, i).Address, "$")(1)
    t = Split(o.Cells(1, j).Address, "$")(1)
    
    o.Columns(s & ":" & t).Hidden = True
End Sub

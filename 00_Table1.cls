VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub ButtonImportSheet_Click()
        UserFormImport.Show
End Sub

Private Sub ButtonLoadDataFeed_Click()
    UserFormDataFeed.Show
End Sub

Private Sub ButtonTransformToID_Click()
    UserFormTransformToID.Show
End Sub

Private Sub StartContent_Click()
    UserFormStart.Show
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormProductsheet 
   Caption         =   "Create Product Data Sheet"
   ClientHeight    =   9540.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14325
   OleObjectBlob   =   "UserFormProductsheet.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormProductsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonReadContent_Click()
    ContentAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *.xlsx")
End Sub


Private Sub ButtonApply_Click()
    
    ' Goal: Create a product data sheet for the suppliers.
    
    ' The entire process contains a lot of code, so the entire script within different modules with specific goals to create the final product sheet.
    ' As soon as every data is taken from the UserForm, it will be closed.
    
    ' Variables
    Dim wb As Workbook
    Dim o, p, q As Object
    Dim sc, sa, si, sp, t As String
    Dim bp, b, b1 As Boolean
    Dim i As Integer
    
    ' In order to keep the information of the content data sheet optional, their information and requirement to be read and copied is connected to a boolean
    bp = UseContent.Value
    If bp = True Then
        If ContentAddress.Caption = "" Or ContentAddress.Caption = "False" Then
            MsgBox "Content data sheet missing"
            Exit Sub
        End If
        sc = ContentAddress.Caption
        ' Check if iPIM or PBK label is chosen, otherwise the program will not continue.
        ' Reminder: If a content from both lists is chosen, then only the iPIM label will be read, PBK will be ignored.
        If UserFormProductsheet.ComboBoxPIM <> "" Then
            t = UserFormProductsheet.ComboBoxPIM
            b = False
        ElseIf UserFormProductsheet.ComboBoxPBK <> "" Then
            t = UserFormProductsheet.ComboBoxPBK
            b = True
        Else
            MsgBox "No label chosen."
            Exit Sub
        End If
    End If
    
    ' We need the data sheet with attributes, otherwise the program will terminate.
    ' Furthermore, we need the data sheet with the associated values.
    If AttributeAddress.Caption = "" Or AttributeAddress.Caption = "False" Then
        MsgBox "No data sheet with attributes chosen"
        Exit Sub
    End If
    If ImportAddress.Caption = "" Or ImportAddress.Caption = "False" Then
        MsgBox "No data sheet with values chosen"
        Exit Sub
    End If
    
    ' The primary data file remains optional. However, if it is loaded, the script will continue with the modules for PBK labels.
    ' If no file is chosen, the script will use the modules written for iPIM Labels
    
    ' Referencing all addresses to string variables
    sa = AttributeAddress.Caption
    si = ImportAddress.Caption
    sp = PrimaryAddress.Caption
    
    ' Boolean variable checks path of primary data sheet
    If sp = "" Or sp = "False" Then
        b1 = False
    Else
        b1 = True
    End If
    
    ' All necessary information stored, User Interface can be closed.
    Unload Me
    
    ' A new File with three sheets will be created.
    ' Product sheet for suppliers
    ' Sheet with Values as default values for Dropdown Menu
    ' Sheet with Database IDs of default values for automated import into database later
    Workbooks.Add
    Set wb = ActiveWorkbook
    Set o = ActiveWorkbook.ActiveSheet
    o.Name = "Product Data Sheet"
    wb.Sheets.Add after:=o
    Set p = wb.ActiveSheet
    p.Name = "Default Values"
    wb.Sheets.Add after:=p
    Set q = wb.ActiveSheet
    q.Name = "Default Values IDs"
    
    ' The script will now continue to call modules. Depending on the modules only specific files and sheets will be handed over.
    ' The modules "Filter", "RemainingData" and "InsertValues" can be set as comments and ignored if necessary.
    
    ' BuildTable:        Module for initial set up of the table
    ' InsertAttributes:  Module for inserting attributes (headers of the default values)
    ' Filter:            Module for ignoring specific attributes not needed for the suppliers
    ' RemainingData:     Module to insert remaining attributes that are not in the attribute datasheet
    ' AttributeValues:   Module to copy attribute (which represent headers) into the sheet "Default Values" and "Default Values IDs" if they contain default values
    ' DefaultValues:     Module for inserting default values into the sheet "Default Values" (only PBK label)
    ' DefaultValuesIDs:  Module for inserting the IDs of default Values (only PBK label)
    ' ImportValues:      Module for inserting default values and their IDs (only iPIM Label)
    ' DropDown:          Module to create DropDown Menues in the sheet "Product Data Sheet" containing the default values
    ' HideColumnsAtEnd:  Module to hide certain columns with content that are not important for the suppliers but for the operators and managers when the product data sheet will be returned
    ' InsertContent:     Module to insert EAN, article number, product number into product sheet (optional)
    
    Call BuildTable(o, p, q)
    Call InsertAttributes(o, sa)
    Call Filter(o)
    Call RemainingData(o)
    Call AttributeValues(o, p, q)
    If b1 = False Then
        Call ImportValues(p, q, si)
    Else
        Call DefaultValues(o, p, sp)
        Call DefaultValuesIDs(p, q, si)
    End If
    Call DropDown(o, p, q)
    Call HideColumnsAtEnd(o)
    
    ' If a content data sheet is chosen, then its content will be inserted into the product data sheet
    If bp = True Then Call WerteEinfügen(o, sc, b, t)
    
    ' In order to know when the script is done, a little messenger box will appear.
    MsgBox "Done"
End Sub

Private Sub ButtonAttributes_Click()
    AttributeAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *.xlsx")
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonHide_Click()
    Hide
End Sub

Private Sub ButtonImport_Click()
    ImportAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *.xlsx")
End Sub

Private Sub ButtonPrimary_Click()
    PrimaryAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *xlsx")
End Sub

Private Sub ButtonReadAttributes_Click()
    ContentAddress.Caption = Application.GetOpenFilename("Excel-Workbook (*.xlsx), *.xlsx")
End Sub


Private Sub ReadContent_Click()
    ' Goal: Read the all iPIM or PBK labels from a content data sheet
    
    ' First check if a data file has been chosen.
    If ContentAddress = "" Or ContentAddress = "False" Then
        MsgBox "Content data file missing"
        Exit Sub
    End If
    
    ' Just in case, delete all content from Comboboxes (in case you want to reload).
    UserFormProductsheet.ComboBoxPIM.Text = ""
    UserFormProductsheet.ComboBoxPBK.Text = ""
    
    ' Variables
    Dim wb1 As Workbook
    Dim o As Object
    Dim s, s1, s2, t As String
    Dim b, bp As Boolean
    Dim i1, i2, j, k1, k2 As Integer
    Dim a1(), a2() As String
    
    ' Read datafile and reference its address to a string variable.
    s = ContentAddress.Caption
    Call LoadFile(wb1, o, s, "Content query")
    
    ' Search for iPIM and PBK columns. We only seach in row 3.
    s1 = "exact location in iPIM"
    s2 = "PBK"
    
    i1 = FindColumn(o, s1, 3)
    
    ' The column search for "PBK" has to be optional, the different content teams do not use the same templates, not every team has this column.
    bp = False
    i2 = 1
    Do Until o.Cells(3, i2) = s2 Or o.Cells(3, i2) = ""
        i2 = i2 + 1
    Loop
    If o.Cells(3, i2) = s2 Then
        bp = True
    Else
        UserFormProductsheet.ComboBoxPBK.Enabled = False
    End If
    
    ' We begin to fill the arrays a1 and a2 with an empty string.
    ReDim a1(0)
    a1(0) = ""
    If bp = True Then
        ReDim a2(0)
        a2(0) = ""
    End If
    
    ' Now we start in row 4 until there is no content left and check the content in the columns iPIM and PBK. If the content is unknown it will be inserted into the array,
    ' otherwise it will be ignored.
    j = 4
    Do Until o.Cells(j, i1) = ""
        ' iPIM
        k1 = 0
        b = False
        Do Until b = True Or k1 > UBound(a1)
            If o.Cells(j, i1) = a1(k1) Then
                b = True
            Else
                k1 = k1 + 1
            End If
        Loop
        ' If the content is unkown, it will be inserted into the array
        If b = False Then
            ReDim Preserve a1(k1)
            a1(k1) = o.Cells(j, i1)
        End If
        ' PBK
        If bp = True Then
            k2 = 0
            b = False
            Do Until b = True Or k2 > UBound(a2)
                If o.Cells(j, i2) = a2(k2) Then
                    b = True
                Else
                    k2 = k2 + 1
                End If
            Loop
            If b = False Then
                ReDim Preserve a2(k2)
                a2(k2) = o.Cells(j, i2)
            End If
        End If
        j = j + 1
    Loop
    
    ' Hand over the arrays to the comboboxes of the userform.
    UserFormProductsheet.ComboBoxPIM.List = a1
    If bp = True Then UserFormProductsheet.ComboBoxPBK.List = a2
End Sub

Private Sub UseContent_Click()
    If UseContent.Value = False Then
        UseContent.Caption = "Don't use Content File"
        ButtonReadAttributes.Enabled = False
        ReadContent.Enabled = False
        ComboBoxPIM.Enabled = False
        ComboBoxPBK.Enabled = False
        Label2.ForeColor = &H80000010
        Label3.ForeColor = &H80000010
        ContentAddress.Caption = ""
    Else
        UseContent.Caption = "Use Content File"
        ButtonReadAttributes.Enabled = True
        ReadContent.Enabled = True
        ComboBoxPIM.Enabled = True
        ComboBoxPBK.Enabled = True
        Label2.ForeColor = &H80000012
        Label3.ForeColor = &H80000012
    End If
End Sub

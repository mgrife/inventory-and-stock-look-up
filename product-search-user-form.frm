VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Productsearch 
   Caption         =   "UserForm1"
   ClientHeight    =   3830
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   6300
   OleObjectBlob   =   "product-search-user-form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Productsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CancelButton_Click()
Unload Me
End
End Sub


Private Sub Gobutton_Click()
'if price option is selected
If PriceOption.Value = True Then
'find the price of the codelist value on the right sheet
With Worksheets("products")
'loop through all codes to find the matching code
For Each code In Range(Range("A2"), Range("A1").End(xlDown))
    'if the cell currently in the loop matches the value selected
    If code = CodesList.Value Then
        Mprice = code.Offset(0, 1)
        MsgBox CodesList.Value & " is priced at $" & Mprice
        Unload Me
        Exit Sub
    End If
    Next code
End With
'if inventory is selected
ElseIf InventoryOption.Value = True Then
'make sure its the correctsheet
    With Worksheets("products")
'loop through all codes to find the matching code
For Each code In Range(Range("A2"), Range("A1").End(xlDown))
    'if the cell currently in the loop matches the value selected
    If code = CodesList.Value Then
        minventory = code.Offset(0, 2)
        MsgBox CodesList.Value & " Inventory level: " & minventory & " Units"
        Unload Me
        Exit Sub
    Else
    End If
    Next code
End With

End If
    

End Sub


Private Sub UserForm_Initialize()
'make sure using the right sheet
With Worksheets("Products")
    'finding the values that begin with the right letter
    For Each Value In Range(Range("A2"), Range("A1").End(xlDown))
        'test the letters against eachother
        characterone = Left(Value, 1)
            If characterone = letter Then
                'store the value of the cell into a variable
                code1 = Value
                'with the list of codes
                With CodesList
                    'add the code to the list
                    .AddItem code1
                End With
            Else
            End If
        Next Value
End With
'making the first code selected the default
CodesList.Selected(O) = True
'making the price option selected the default
PriceOption.Value = True


End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Newproduct 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   5900
   OleObjectBlob   =   "add-new-product.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Newproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End
End Sub


Private Sub Gobutton_Click()

Price = Price.Text
inventory = InventoryLevel.Text
ProductCode = ProductCode.Text
'determine if a new product code or previously used one
isnew = True
If Range("A2") <> "" Then
    For Each x In Range(Range("A2"), Range("A1").End(xlDown))
        If UCase(x.Value) = UCase(ProductCode) Then
        isnew = False
        Exit For
        End If
    Next x
End If

'if not a valid code or or not a new code
If Not iscodevalid(ProductCode) Or isnew = False Then
    MsgBox "product code not valid or already in the list"
    ProductCode.SetFocus
    Exit Sub
    End If
'if price is not valid
If Not ispricevalid(Price) Then
    MsgBox "Price not valid"
    Price.SetFocus
    Exit Sub
  End If
  
  
'if inventory is not valid
Valid = True
If Not IsNumeric(inventory) Then
Valid = False
InventoryLevel.SetFocus
MsgBox "Inventory level isn't valid"
Exit Sub
ElseIf Int(inventory) <> CDbl(inventory) Or inventory < 0 Then
Valid = False
InventoryLevel.SetFocus
MsgBox "inventory level isn't valid"
Exit Sub
End If


'put record on the worksheet
If Range("A2") = "" Then
Range("A2") = ProductCode
Range("B2") = Price
Range("C2") = inventory
Else
Range("A1").End(xlDown).Offset(1, 0) = ProductCode
Range("B1").End(xlDown).Offset(1, 0) = FormatCurrency(Price)
Range("C1").End(xlDown).Offset(1, 0) = inventory
End If


ProductCode = ""
Price = ""
InventoryLevel = ""
ProductCode.SetFocus
      

End Sub

Private Sub UserForm_Click()

End Sub

Attribute VB_Name = "inventory_stock"
Public Price As Double
Public ProductCode As String
Public InventoryLevel As Integer
Public letter
Public allcodes



Sub Problem1()
Dim usedletter As Boolean
Dim lvalid As Boolean

'start a loop
Do
    'get value from person
    letter = InputBox("please enter a single letter")
    'uppercase the value
    letter = UCase(letter)
    'make sure its exactly one input
    If Len(letter) <> 1 Then
    'if not make the holder value false
        lvalid = False
        MsgBox "That's not a valid character length. Please try again"
    'make sure the one value is a letter
    ElseIf Asc(letter) < 65 Or Asc(letter) > 90 Then
        lvalid = False
        MsgBox "That is not a letter. Please try again"
'else it is a valid input
    Else
        lvalid = True
'end the statement
    End If

'end loop
Loop Until lvalid = True



'determine if there are codes that start with that letter
'holder variable set to false and only change it if at some point there is a cell that matches letter
userletter = False
'loop through each cell in the range
For Each cell In Range(Range("A2"), Range("A1").End(xlDown))
    characterone = Left(cell, 1)
    'make sure that the character = letter
    If characterone = letter Then
        'flip holder variable
        usedletter = True
        'variable that will store that code
        code1 = cell
        allcodes = allcodes & vbNewLine & code1
    End If
Next cell
'if there exist codes that start with that letter than show the userform
If usedletter = True Then
    'show userform
    ProductSearch.Show
    'if there are not codes that start with that letter end the code with a relevant message
Else
    MsgBox "No code begins with " & letter
    Exit Sub
End If

End Sub



Public Function iscodevalid(ProductCode) As Boolean
'function to later be called to make sure the code is valid for problem 2
If Len(ProductCode) <> 5 Then
    iscodevalid = False
    
ElseIf Not IsNumeric(Right(ProductCode, 4)) Then
    iscodevalid = False
    
ElseIf Asc(UCase(Left(ProductCode, 1))) < 65 Or Asc(UCase(Left(ProductCode, 1))) > 90 Then
    iscodevalid = False
    
Else
    iscodevalid = True
    
End If
    
End Function
Public Function ispricevalid(Price) As Boolean
'function to later be called to make sure the price is valid for problem 2

If Len(Price) < 3 Then
    ispricevalid = False
    
ElseIf IsNumeric(Price) And Mid(Price, Len(Price) - 2, 1) = "." Then
    ispricevalid = True
    
Else
    ispricevalid = False

End If
    
End Function

Sub problem_2()
Worksheets("Products").Activate
Newproduct.Show

End Sub


Sub problem3()
Attribute problem3.VB_ProcData.VB_Invoke_Func = "s\n14"
Stockoptions.Show
End Sub



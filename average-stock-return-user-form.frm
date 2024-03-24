VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Stockoptions 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   4580
   OleObjectBlob   =   "average-stock-return-user-form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Stockoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()
Unload Me
End
End Sub

Private Sub OkButton_Click()
'loop through each index of the list
For i = 0 To Stocklist.ListCount - 1
'if it is selected
If Stocklist.Selected(i) = True Then
    'go to that worksheet
    Worksheets(i + 2).Activate
      'average all the returns
        avg = WorksheetFunction.Average(Range(Range("C4"), Range("C4").End(xlDown)))
        avg = FormatPercent(avg, 2)
    'combine that variable using a new line and store
    msg = Worksheets(i + 2).Name & ": " & avg & vbNewLine
    finalmsg = msg & finalmsg
End If
Next i
Worksheets(1).Activate
MsgBox finalmsg
Unload Me
End Sub

Private Sub UserForm_Initialize()
'add all the symbols into the listbox
With Stocklist
'loop through each worksheet
numberofsheets = Worksheets.Count
For i = 2 To numberofsheets
'add the item to the list
    .AddItem Worksheets(i).Name
Next i
End With
End Sub

Attribute VB_Name = "bShare"
Option Explicit

Sub ShareP(Share As String)
Dim LastRow As Integer
Dim SP As Worksheet
Dim i As Integer

Set SP = Worksheets("SharePoint")

'set the share content to choose further
Share = UserForm1.ComboBox2

LastRow = SP.Cells(SP.Rows.count, "B").End(xlUp).Row

For i = 2 To LastRow

If Share = SP.Cells(i, 2).Value Then
Result = "SharePoint for " & SP.Cells(i, 2) & " : " & SP.Cells(i, 3)
End If

Next i
     
   
End Sub

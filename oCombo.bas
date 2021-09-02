Attribute VB_Name = "oCombo"
Sub ComboPop()
Dim DB As Worksheet
Dim BoxNr As Integer
Dim LastRow As Integer
Dim StarterCombo As Integer


'Create the object oCombo's as clsComboBox class
Dim oCombo1 As New clsComboBox
'Create the collection
Dim comboRead As Collection
'Create the collection
Dim Combo1 As New Collection

Set DB = Worksheets("Db")

'set Column number ( - 1)
StarterCombo = 4

For i = 1 To 2

'Insert items in the collection combobox
LastRow = DB.Cells(DB.Rows.count, i + StarterCombo).End(xlUp).Row

    For count = 3 To LastRow
    Combo1.Add DB.Cells(count, i + StarterCombo).Value
    Next


'Set the label collection to object
Set oCombo1.Combo1 = Combo1

Set comboRead = oCombo1.Combo1
BoxNr = DB.Cells(1, i + StarterCombo).Value

UpdateCombo comboRead, BoxNr

'clear the collection for getting empty on next comboBox
Set Combo1 = Nothing
Set oCombo1.Combo1 = Combo1

Next


End Sub

' Use the contents of a Collection
Sub UpdateCombo(c As Collection, BoxNr As Integer)

    Dim item As Variant
    For Each item In c
        
    'UserForm1.ComboBox1.AddItem item
     UserForm1.Controls("ComboBox" & BoxNr).AddItem item
       
    Next item
  ' set the first value in the combobox
    UserForm1.Controls("ComboBox" & BoxNr).ListIndex = 0
       
End Sub


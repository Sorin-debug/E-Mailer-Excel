Attribute VB_Name = "oLab"
Sub Labelpop()
Dim DB As Worksheet

Set DB = Worksheets("Db")

'Create the collection of label and add labels.
    LastRow = DB.Cells(DB.Rows.count, "B").End(xlUp).Row
    Dim label As New Collection
    For count = 3 To LastRow
    label.Add Worksheets("db").Cells(count, 2).Value
    Next

'create the object oLabel as clsLabel class
Dim oLabel As New clsLabel

'Set the label collection to object
Set oLabel.Labels = label

Dim labelRead As Collection
Set labelRead = oLabel.Labels

UpdateLabel labelRead


End Sub


' Print the contents of a Collection to the Immediate Window(Ctrl + G)
Sub UpdateLabel(c As Collection)

Dim i As Integer
i = 1
    Dim item As Variant
    For Each item In c
       ' Debug.Print item
     UserForm1.Controls("Label" & i).Caption = item
        i = i + 1
          
       
    Next item



End Sub


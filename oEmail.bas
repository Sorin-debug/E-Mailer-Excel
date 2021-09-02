Attribute VB_Name = "oEmail"
Option Explicit


Public Sub EmaConfig()
Dim DB As Worksheet
Dim count As Integer
Dim LastRow As Integer
Dim Ema As String

Set DB = Worksheets("Db")

 
    LastRow = DB.Cells(DB.Rows.count, "C").End(xlUp).Row
    For count = 3 To LastRow
    Ema = Ema + DB.Cells(count, 3).Value
    Next
    
    ' Use CreateSimpleObject function
    Dim oTo As clsEmail
    Set oTo = CreateSimpleObject(Ema)
       
   ToEmail = oTo.m_Label
   
   'Clear the Ema
   Ema = ""
   
   'Adding CC email data
    LastRow = DB.Cells(DB.Rows.count, "D").End(xlUp).Row
    For count = 3 To LastRow
    Ema = Ema + DB.Cells(count, 4).Value
    Next
    
    ' Use CreateSimpleObject function
    Dim oCc As clsEmail
    Set oCc = CreateSimpleObject(Ema)
       
   CcEmail = oCc.m_Label
    
    

End Sub

Public Function CreateSimpleObject(label As String) As clsEmail

    Dim oSimple As New clsEmail
    oSimple.Init label
    Set CreateSimpleObject = oSimple

End Function




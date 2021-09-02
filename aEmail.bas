Attribute VB_Name = "aEmail"
Option Explicit
Public ToEmail As String
Public CcEmail As String
Public YearWeek As String
Public Result As String
Public Share As String

Sub Send_Email()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim DB As Worksheet
    Set DB = Worksheets("Db")
   
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'get the week number and year
    YearWeekNumber
          
     strbody = "Hello Team," & vbNewLine & vbNewLine & "Were saved new changes on SharePoint:" & vbNewLine & vbNewLine
                
    On Error Resume Next
    
    'call procedure for Configuration for ToEmail and CcEmail
    EmaConfig
      
    
    'Call for sharepoint
    ShareP (Share)
      
    With OutMail
        .To = ToEmail
        .CC = CcEmail
        .BCC = ""
        .Subject = "Project - AUDI_Q7 - " & "Drawing Date: " & UserForm1.ComboBox1 & " - " & "Harness: " & UserForm1.ComboBox2 & " - " & " Change translations " & YearWeek
        .Body = strbody & vbNewLine & UserForm1.TextBox1 & vbNewLine & vbNewLine & Result & vbNewLine & "Thanks," & vbNewLine
        '
        'You can add a file like this
        'Attachments.Add strdefpath & ("\NewCr.xlsx")
        '.Send
        .display
       
                      
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
       
    
End Sub
 








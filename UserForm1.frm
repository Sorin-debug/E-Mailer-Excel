VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "E-mailer @ 2021 - Sorin Petcu "
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9645.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
Send_Email
End Sub

Private Sub UserForm_Initialize()
Dim lTop As Long, lLeft As Long
Dim lRow As Long, lCol As Long

    With ActiveWindow.VisibleRange
        lRow = .Rows.count
        lCol = .Columns.count
    End With

    With Cells(lRow, lCol)
        lTop = .Top
        lLeft = .Left
    End With
    
    With Me
       .Top = lTop
       .Left = lLeft
    End With
    
'Dim zone
  
  Dim DB As Worksheet
  Set DB = Worksheets("Db")
  
  
'Labeling
Labelpop
    
'Codes for activating ComboBoxes
ComboPop



End Sub

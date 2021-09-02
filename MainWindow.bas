Attribute VB_Name = "MainWindow"
Option Explicit

Sub Message()

With UserForm1
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
 .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
 .Show
End With


End Sub


Sub Location()
MsgBox prompt:="I am here:   " & Application.ActiveWorkbook.Path
End Sub

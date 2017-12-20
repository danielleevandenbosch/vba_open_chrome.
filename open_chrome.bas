Attribute VB_Name = "Module1"
Option Compare Text
Option Explicit
Public Sub open_chrome(Optional url As String = "about:blank")
Call run_batch_script("@echo off" & vbCrLf & _
"""C:\Program Files (x86)\Google\Application\chrome.exe"" """ & url & """" & vbCrLf & _
"""C:\Documents and Settings\%username%\Local Settings\Application Data\Google\Chrome.exe"" """ & url & """" & vbCrLf & _
"C:\Users\%UserName%\AppDataLocal\Google\Chrome.exe  """ & url & """" & vbCrLf & _
"""C:\Program Files\Google\Chrome\Application\chrome.exe"" """ & url & """" & vbCrLf & _
"""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" """ & url & """")
If Err.number <> 0 Then Lmsgbox Error$ & " there was an error on public sub open_chrome"
End Sub


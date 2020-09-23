Attribute VB_Name = "modVirtualFunctions"
Option Explicit
Global readytostart As Boolean
Global ready_next As String

Function PathGetPath(PATH As String) As String
Dim bak As String
Dim i As Integer

For i = Len(PATH) To 1 Step -1
If (Mid(PATH, i, 1) = "\") Or (Mid(PATH, i, 1) = "/") Then
bak = Mid(PATH, 1, i - 1)
Exit For
End If
Next i

PathGetPath = bak
End Function

Function ProcessURL(sURL As String, client As WebBrowser) As Boolean
Dim sBuffer As String
ProcessURL = True
If Left(sURL, Len("install::")) = "install::" Then
    sBuffer = Replace(sURL, "install::", "")
    Select Case sBuffer
    Case "quit"
        Dim res As VbMsgBoxResult
        
        res = MsgBox("Are you sure you want to quit?", vbInformation Or vbYesNo, client.Document.Title)
        If res = vbYes Then
            client.Stop
            End
        Else
            
        End If
    Case "quit2"
        End
    End Select
    ProcessURL = False
End If


End Function

Function SetProgressBarPos(x As Long, y As Long)
frmMain.progress.Move x, y
End Function

Function SetBrowsePos(x As Long, y As Long)
frmMain.txtBrowse.Move x, y

frmMain.cmdBrowse.Move x + 10 + frmMain.txtBrowse.Width, y
frmMain.cmdBrowse.Visible = True
frmMain.txtBrowse.Visible = True
End Function

Function CopyDB(sFile As String)
On Error Resume Next
Kill App.PATH & "\list.txt"
DownloadFromWeb PATH & "/" & sFile, App.PATH & "\list.txt"
End Function

Function DoDB()
Dim sLine As String
Dim sData() As String
Dim count As String

Open App.PATH & "\list.txt" For Input As #1
Line Input #1, count

Do While Not EOF(1)
    
    Line Input #1, sLine
    If sLine = "" Then GoTo skip
    sData = Split(sLine, vbTab)
    Select Case sData(0)
    Case "user"
        DownloadFromWeb PATH & "/" & sData(1), Replace(frmMain.txtBrowse.Text & "\" & sData(1), "\\", "\")
    Case "windows"
        DownloadFromWeb PATH & "/" & sData(1), "c:\windows\" & sData(1)
    Case "system"
        DownloadFromWeb PATH & "/" & sData(1), "c:\windows\system\" & sData(1)
    Case "system32"
        DownloadFromWeb PATH & "/" & sData(1), "c:\windows\system32\" & sData(1)
    Case Else
        'MsgBox "invalid destination: sData(0)
        DownloadFromWeb PATH & "/" & sData(1), sData(2)
    End Select
    frmMain.progress.Value = frmMain.progress.Value + (100 / CInt(count))

skip:
Loop
Close #1

frmMain.client.Navigate PATH & "/" & ready_next
End Function
Function ProcessStatus(sText As String)
Dim sBuffer As String
Dim sData() As String
If Left(sText, Len("install::")) = "install::" Then
    sBuffer = Replace(sText, "install::", "")
    sData = Split(sBuffer, ",")
    If sData(2) = "browse" Then
        SetBrowsePos CLng(sData(0)), CLng(sData(1))
    ElseIf sData(2) = "progress" Then
        SetProgressBarPos CLng(sData(0)), CLng(sData(1))
    ElseIf sData(2) = "copy" Then
        frmMain.client.Stop
        ready_next = sData(1)
        readytostart = True
        CopyDB sData(0)
        
        
    End If
End If
End Function

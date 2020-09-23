Attribute VB_Name = "modDownload"
Option Explicit
Global PATH As String

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Function DownloadFromWeb(ByVal strURL As String, ByVal SaveFilePathName As String) As Long
    On Error Resume Next
    DownloadFromWeb = URLDownloadToFile(0, strURL, SaveFilePathName, 0, 0)
End Function





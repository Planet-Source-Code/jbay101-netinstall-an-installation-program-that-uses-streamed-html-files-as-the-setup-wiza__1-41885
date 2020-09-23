VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "netInstall - Loading installation"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1455
      Top             =   3105
   End
   Begin VB.PictureBox picLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   2340
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   1
      Top             =   2685
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   825
         Left            =   -15
         Top             =   -15
         Width           =   3720
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "frmMain.frx":058A
         Stretch         =   -1  'True
         Top             =   165
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading installation...please wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   2
         Top             =   285
         Width           =   2520
      End
   End
   Begin MSComctlLib.ProgressBar progress 
      Height          =   540
      Left            =   510
      TabIndex        =   5
      Top             =   4785
      Visible         =   0   'False
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   953
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtBrowse 
      Height          =   315
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2145
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3660
      TabIndex        =   3
      Top             =   2130
      Visible         =   0   'False
      Width           =   1305
   End
   Begin SHDocVwCtl.WebBrowser client 
      CausesValidation=   0   'False
      Height          =   225
      Left            =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1155
      Width           =   240
      ExtentX         =   423
      ExtentY         =   397
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub client_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
tmrCheck.Enabled = False
If ProcessURL(CStr(URL), client) Then
'    Me.Caption = "netInstall - Loading installation"
'    client.Visible = False
'    picLoading.Visible = True
    Me.MousePointer = 11
Else
    Cancel = True
End If



End Sub

Private Sub client_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
cmdBrowse.Visible = False
txtBrowse.Visible = False
progress.Visible = False
Me.Caption = "netInstall - " & client.Document.Title
Me.MousePointer = 0
'picLoading.Visible = False
client.Visible = True
PATH = PathGetPath(client.LocationURL)

tmrCheck.Enabled = True
'On Error Resume Next
'ProcessStatus client.StatusText

End Sub

Private Sub client_StatusTextChange(ByVal Text As String)
ProcessStatus Text
End Sub

Private Sub cmdBrowse_Click()
txtBrowse.Text = SelectFolder(Me)
End Sub

Private Sub Form_Load()
Dim res As VbMsgBoxResult
client.Navigate "about:blank"
MsgBox "This sample will stream an installation package from my web site. It will deliver a small package."
res = MsgBox("[YES] Online test (streamed)" & vbCrLf & "[NO] Offline test (local)" & vbCrLf & vbCrLf & "if online doesn't work, try offline!", vbYesNo)

If res = vbYes Then
    client.Navigate "http://flux3d.port5.com/netinstall/default.htm" ' App.PATH & "\sample\default.htm"
Else
    client.Navigate App.PATH & "\sample\default.htm"
End If
End Sub

Private Sub Form_Resize()
client.Move -1, -1, Me.ScaleWidth + 30, Me.ScaleHeight + 2 '- 10 '+ 30
picLoading.Move Me.ScaleWidth / 2 - picLoading.Width / 2, Me.ScaleHeight / 2 - picLoading.Height / 2
End Sub

Private Sub tmrCheck_Timer()
If readytostart = True Then
tmrCheck.Enabled = False
progress.Visible = True
DoDB
readytostart = False
End If

End Sub

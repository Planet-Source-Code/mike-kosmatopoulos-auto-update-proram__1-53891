VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Updater"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   30
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   30
      ExtentX         =   53
      ExtentY         =   53
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   30
      ExtentX         =   53
      ExtentY         =   53
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check for updates!"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtNew 
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmUpdate.frx":0000
      Top             =   2040
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   5
   End
   Begin VB.Label lblNewest 
      Caption         =   "Newest Version: No information"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current version: "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Waiting for user..."
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''Program Auto Updater by Mike Kosmatopoulos'''
'Edit the information below to make the auto updater work!!
'For more information, read Readme.txt that came with this source
Const InfoSite As String = "http://www.freepgs.com/mike2/update.html" 'Website to check if a new version is available, check what's new etc
Const DownloadSite As String = "http://www.freepgs.com/mike2/AutoUpdater.zip" 'Website to download new version if there is
Const CurrentVersion As String = "1.0" 'Current version to compare with the one the program "downloaded" from the website

Dim bChecking As Boolean
Dim NewVersion As String
'Don't forget to change the Message Box texts! Press Ctrl + F and search for "MsgBox"

Private Sub cmdCheck_Click()
If bChecking = True Then
    MsgBox "Already checking for updates.", vbInformation, "Error"
        Exit Sub
End If
bChecking = True

Bar.Visible = True
Bar.Value = Bar.Min
lblStatus.Caption = "Status: Checking for the newest version..."
Bar.Value = Bar.Value + 1
GetNewVersionNumber
lblStatus.Caption = "Status: Retrieving if the lastest version is beta..."
Bar.Value = Bar.Value + 1
CheckIfBeta
Bar.Value = Bar.Value + 1



If lblNewest.Caption = "Newest Version: " & CurrentVersion Then
lblStatus.Caption = "No new version available"
Bar.Value = Bar.Max
Bar.Visible = False
Else
Bar.Value = Bar.Value + 1
lblStatus.Caption = "Status: New version available! Retrieving what's new in this version..."
GetWhatsNew
lblStatus.Caption = "Done. New version available for download."
Dim Response As String
Response = MsgBox("A new version of Auto Updater is available!Your current version is " & CurrentVersion & " and the newest version is " & NewVersion & "." & vbCrLf & vbCrLf & vbCrLf & txtNew.Text & vbCrLf & vbCrLf & "Would you like to update now?", vbYesNo + vbQuestion + vbSystemModal, "New version available!")

Select Case Response

    Case vbYes
 Bar.Value = Bar.Value + 1
    Browser.Visible = True
Browser.Navigate DownloadSite
MsgBox "Please wait a few seconds for the download screen to load." & vbCrLf & vbCrLf & "When done downloading, click ""OK"" to terminate Auto Updater to install the new version.", vbInformation, "Auto Updater"
Dim X As Form
For Each X In Forms
Unload X
Next X
bChecking = False
End
Case vbNo
Bar.Value = Bar.Value + 1
MsgBox "Run the Auto-Updater to update at later time.", vbInformation + vbSystemModal, "Auto Updater"
bChecking = False
End Select


'Bar.Value = Bar.Value + 1
Bar.Visible = False
End If




End Sub
Private Sub GetNewVersionNumber()
Dim strSource As String
Dim strSplit() As String

Do While strSource = ""
strSource = Inet1.OpenURL(InfoSite)
Loop

strSource = Replace(strSource, "</b>", "<b>", , , vbTextCompare)
strSplit = Split(strSource, "<b>", , vbTextCompare)

strSplit = Split(strSplit(1), " ", , vbTextCompare)

lblNewest.Caption = "Newest Version: " & strSplit(0)
NewVersion = strSplit(0)
'If strSplit(0) = "1.1 Beta" Then
'lblStatus.Caption = "No new version available"
'End If
End Sub
Private Sub CheckIfBeta()
Dim strSource As String
Dim strSplit() As String

Do While strSource = ""
strSource = Inet1.OpenURL(InfoSite)
Loop

strSource = Replace(strSource, "</b>", "<b>", , , vbTextCompare)
strSplit = Split(strSource, "<b>", , vbTextCompare)

strSplit = Split(strSplit(5), " ", , vbTextCompare)  'Change strSplit(x)

If strSplit(0) = "Yes" Then
NewVersion = NewVersion & " Beta"
lblNewest.Caption = lblNewest.Caption & " Beta"
End If
End Sub

Private Sub GetWhatsNew()
Dim strSource As String
Dim strSplit() As String
Do While strSource = ""
strSource = Inet1.OpenURL(InfoSite)
Loop
strSource = Replace(strSource, "</b>", "<b>", , , vbTextCompare)
strSplit = Split(strSource, "<b>", , vbTextCompare)

strSplit(3) = Replace(strSplit(3), "%n", vbCrLf)

txtNew.Text = "What's new?" & vbCrLf & vbCrLf & strSplit(3)

End Sub




Private Sub Form_Load()
lblCurrent.Caption = "Cuurent Version: " & CurrentVersion
End Sub

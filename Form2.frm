VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FWD File Encoding Converstion to ANSI"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   105
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Batch Convert to ANSI"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select FWD Folder"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'declare variables
Dim FWDFolder As String

'select folder code
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type


'select folder in treeview control listing computer directories
Private Sub Command1_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo

    szTitle = "Select the directory containing the FWD files."
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        'MsgBox sBuffer
        FWDFolder = sBuffer
        Text1.Text = sBuffer
    End If

End Sub


'batch convert encoding
Private Sub Command2_Click()
If Text1.Text <> "" Then
    On Error Resume Next
    'disble the buttons until completed
    Command1.Enabled = False
    Command2.Enabled = False

    Dim myfile As String
    Dim path, spath As String
        
    'create the output folder
    path = Text1.Text & "\"
    Shell "cmd /a/c" & Chr$(34) & "md " & Chr$(34) & path & "ansi" & Chr$(34), vbHide
    
    'count the txt files in the current directory
    Dim e As String
    Dim f, g As Integer
    f = 0
    g = 0
    e = Dir(path & "*.fwd")
    Do While e <> ""
    f = f + 1
    e = Dir()
    Loop
        
    'initialize progress bar
    ProgressBar1.Max = f
    ProgressBar1.Value = 0
    
    myfile = Dir(path & "*.fwd")
    'MsgBox myfile, vbInformation, "Message"
    
    If myfile = "" Then
        GoTo Missing
    
    Else
        Do While myfile <> ""
            Dim ansiconvert As String
            'double quoutes would be """ or Chr$(34)
            ansiconvert = "TYPE " & Chr$(34) & path & myfile & Chr$(34) & " > " & Chr$(34) & path & Chr$(34) & "ansi\" & Chr$(34) & myfile & Chr$(34)
            Shell "cmd /a/c" & ansiconvert, vbHide
            Label1.Caption = Fix(((ProgressBar1.Value + 1) / ProgressBar1.Max) * 100) & "%" & " completed."
            Label1.Refresh
            g = g + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            myfile = Dir()
        Loop
        MsgBox "FWD file successfully converted to ANSI encoding.", vbInformation, "Message"
    End If
    GoTo Term
    
Missing:
    MsgBox "No FWD files in the selected directory.", vbInformation, "Message"

Else
    MsgBox "No FWD file directory selected.", vbInformation, "Message"

End If

Term:
    Label1.Caption = ""
    ProgressBar1.Value = 0
    Command1.Enabled = True
    Command2.Enabled = True


End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoadmdbFile 
      Caption         =   "Load mdb File"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tables"
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2775
      Begin VB.ListBox List1 
         Height          =   4350
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7215
   End
   Begin VB.CommandButton cmdAddFWDData 
      Caption         =   "Add FWD Data"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoadtxtFile 
      Caption         =   "Load Text File"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   8760
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


Dim fileNameFWD As String
Dim filePathFWD As String
Dim fileContents As String
Dim lines() As String
Dim ff As Integer


'creating a new database
Private Sub create_mdb()
    Dim newdb As ADOX.Catalog
    Dim newtbl As ADOX.table
    
    Set newdb = CreateObject("ADOX.Catalog")
    Set newtbl = New ADOX.table
    
    newdb.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePathFWD & ";Jet OLEDB:Engine Type=5"
      ' Engine Type=5 = Access 2000 Database
      ' Engine Type=4 = Access 97 Database
    With newtbl
        .Name = fileNameFWD
        
        With .Columns
            .Append "DATA", adVarWChar, 1
            .Append "DISTANCE", adVarWChar, 5
            .Append "LOAD", adVarWChar, 25
            .Append "D0", adVarWChar, 25
            .Append "D1", adVarWChar, 25
            .Append "D2", adVarWChar, 25
            .Append "D3", adVarWChar, 25
            .Append "D4", adVarWChar, 25
            .Append "D5", adVarWChar, 25
            .Append "D6", adVarWChar, 25
            .Append "AIR_TEMP", adVarWChar, 4
            .Append "PAVE_TEMP", adVarWChar, 4
            .Append "EMOD", adVarWChar, 10
            .Append "LAT", adVarWChar, 25
            .Append "LON", adVarWChar, 25
            .Append "CHAINAGE", adVarWChar, 5
            .Append "TIME", adVarWChar, 10
            .Append "SURVEY_ID", adVarWChar, 255
        End With
    End With
    
    newdb.Tables.Append newtbl
        
    Set newdb = Nothing
    Set newtbl = Nothing

End Sub


'open database connection
Private Sub open_mdb()
    Set con = New ADODB.Connection
    Set recs = New ADODB.Recordset
    With con
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & filePathFWD & ";Persist Security Info=False"
    End With
    
    With recs
        .CursorLocation = adUseClient
        .ActiveConnection = con
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
End Sub


'close database connection
Private Sub close_mdb()
    Set con = Nothing
End Sub


'display tables
Private Sub display_tbl()
table:
    Dim X As Integer
    X = 0
    List1.Clear
    Call open_mdb
    Set recs = con.OpenSchema(adSchemaTables)
    With recs
    Do While Not .EOF
        X = X + 1
        Dim Y As String
        Y = recs("TABLE_NAME")
        If Not (Y Like "MSys*") Then
        'filter tables with MSys and F_D names
        'If Not (y Like "MSys*" Or y Like "F_D*") Then
            List1.AddItem Y
        End If
        'display all tables
        'List1.AddItem recs("TABLE_NAME")
        .MoveNext
    Loop
    End With
    Call close_mdb
End Sub


Private Sub AnalyzeFile(fileName As String)
    Dim i As Integer
    Dim sqlstatement As String
    Dim fields As String
    Dim field_num As Integer
    Dim delimeter As String
    
    fileContents = LoadFile(fileName)
    lines = Split(fileContents, vbCrLf)
    delimeter = " "
    
    Call open_mdb
    With con
    
    For i = 25 To UBound(lines)
        If InStr(1, lines(i), "D  ", vbTextCompare) Then
            'MsgBox "Line #" & i & " contains the string 'D'" + vbCrLf + vbCrLf + lines(i)
            sqlstatement = "INSERT INTO " & fileNameFWD & " VALUES "
            
            fields = Split(Line(i), delimeter)
            For field_num = LBound(fields) To UBound(fields)
            Next field_num
            
            .Execute sql_statement
            
        End If
    Next
    
    End With
    Call close_mdb
End Sub


'loading the text file
Private Function LoadFile(dFile As String) As String
    On Error Resume Next
    ff = FreeFile
    Open dFile For Binary As #ff
        LoadFile = Space(LOF(ff))
        Get #ff, , LoadFile
    Close #ff
End Function


'button to add fwd data to the mdb file
Private Sub cmdAddFWDData_Click()
    AnalyzeFile cDlg.fileName
End Sub


'load fwd file and create mdb
Private Sub cmdLoadtxtFile_Click()
    cDlg.DefaultExt = "txt"
    cDlg.Filter = "Text Files|*.txt;*.log"
    cDlg.ShowOpen
    
    fileNameFWD = Left$(cDlg.FileTitle, Len(cDlg.FileTitle) - 4)
    filePathFWD = App.Path & "\" & fileNameFWD & ".mdb"
    
    On Error Resume Next
    Text1.Text = filePathFWD
    Call create_mdb
    
    'Call display_tbl
    'MsgBox fileNameFWD
    'MsgBox filePathFWD
    
    'If cDlg.fileName <> "" Then AnalyzeFile cDlg.fileName
End Sub

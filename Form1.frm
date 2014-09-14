VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KUAB FWD Processing Software"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   14430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConvertANSI 
      Caption         =   "Convert to &ANSI"
      Height          =   375
      Left            =   12360
      TabIndex        =   16
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Match Coordinates"
      Height          =   375
      Left            =   12360
      TabIndex        =   15
      Top             =   3240
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   105
      Left            =   120
      TabIndex        =   13
      Top             =   6480
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdBatchExport 
      Caption         =   "&Batch Export"
      Height          =   375
      Left            =   12360
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuickExport 
      Caption         =   "&Quick Export"
      Height          =   375
      Left            =   12360
      TabIndex        =   11
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportCSV 
      Caption         =   "&Export to CSV"
      Height          =   375
      Left            =   12360
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   12015
   End
   Begin VB.CommandButton cmdConvertCoordinates 
      Caption         =   "&Convert Coordinates"
      Height          =   375
      Left            =   12360
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Records"
      Height          =   4695
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   9135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4350
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7673
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdLoadmdbFile 
      Caption         =   "Load &mdb File"
      Height          =   375
      Left            =   12360
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tables"
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1200
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
      Width           =   12015
   End
   Begin VB.CommandButton cmdAddFWDData 
      Caption         =   "&Import FWD Data"
      Height          =   375
      Left            =   12360
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoadtxtFile 
      Caption         =   "Load &FWD File"
      Height          =   375
      Left            =   12360
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   14040
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   13320
      Picture         =   "Form1.frx":08CA
      Top             =   6000
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   12015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declare the variables
Dim fileNameFWD As String
Dim filePathFWD As String
Dim con As ADODB.Connection
Dim recs As ADODB.Recordset
Dim num_records As Long


'send mail to xfuentes@gmail.com
Private Const IDC_HAND = 32649&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5


'create new database
Public Sub create_mdb()
    Dim newdb As ADOX.Catalog
    Dim newtbl As ADOX.table
    
    Set newdb = CreateObject("ADOX.Catalog")
    Set newtbl = New ADOX.table
    
    newdb.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePathFWD & ";Jet OLEDB:Engine Type=5"
      ' Engine Type=5 = Access 2000 Database
      ' Engine Type=4 = Access 97 Database
    With newtbl
        .Name = "Data"
        
        With .Columns
            .Append "D", adVarWChar, 255
            .Append "DISTANCE", adVarWChar, 255
            .Append "IMPACT_NO", adVarWChar, 255
            .Append "LOAD", adVarWChar, 255
            .Append "D0", adVarWChar, 255
            .Append "D1", adVarWChar, 255
            .Append "D2", adVarWChar, 255
            .Append "D3", adVarWChar, 255
            .Append "D4", adVarWChar, 255
            .Append "D5", adVarWChar, 255
            .Append "D6", adVarWChar, 255
            .Append "AIR_TEMP", adVarWChar, 255
            .Append "PAVE_TEMP", adVarWChar, 255
            .Append "EMOD", adVarWChar, 255
            .Append "LAT", adVarWChar, 255
            .Append "LON", adVarWChar, 255
            .Append "CHAINAGE", adVarWChar, 255
            .Append "TIME", adVarWChar, 255
            .Append "SURVEY_ID", adVarWChar, 255
        End With
    End With
    
    newdb.Tables.Append newtbl
        
    Set newdb = Nothing
    Set newtbl = Nothing

End Sub


'open database connection
Public Sub open_mdb()
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
Public Sub close_mdb()
    Set con = Nothing

End Sub


'refresh datagrid on click
Public Sub refresh_dgrid()
    Call open_mdb
    With recs
        .Open "Select * from " & List1.Text & ""
            If .RecordCount <> 0 Then
                Set DataGrid1.DataSource = recs
            Else
                Set DataGrid1.DataSource = Nothing
                'replaced with autoclose
                MsgBox "No records to display.", vbInformation, "Message"
                'Call ACmsgbox(3, "No records to display.", vbInformation, "Message")
            End If
    End With
    Call close_mdb

End Sub


'display tables
Public Sub display_tbl()
table:
    Dim x As Integer
    x = 0
    List1.Clear
    Call open_mdb
    Set recs = con.OpenSchema(adSchemaTables)
        
    With recs
    Do While Not .EOF
        x = x + 1
        Dim y As String
        y = recs("TABLE_NAME")
        If Not (y Like "MSys*") Then
            List1.AddItem y
        End If
        .MoveNext
    Loop
    End With
    
    Call close_mdb

End Sub


'import fwd data option 1
Public Sub AnalyzeFile(fileName As String)
    Dim i As Integer
    Dim sqlstatement As String
    Dim fields As String
    Dim field_num As Integer
    Dim delimeter As String
    Dim lines() As String
    Dim fileContents As String
    
    fileContents = LoadFile(fileName)
    lines = Split(fileContents, vbCrLf)
    delimeter = " "
    
    Call open_mdb
    With con
    
    For i = 25 To UBound(lines)
        If InStr(1, lines(i), "D  ", vbTextCompare) Then
            MsgBox "Line #" & i & " contains the string 'D'" + vbCrLf + vbCrLf + lines(i)
            'sqlstatement = "INSERT INTO " & fileNameFWD & " VALUES "
            
            'fields = Split(lines(i), " ")
            'MsgBox fields, vbInformation, "Message"
            'For field_num = LBound(fields) To UBound(fields)
            'Next field_num
            
            '.Execute sql_statement
            
        End If
    Next
    
    End With
    Call close_mdb

End Sub


'import fwd data option 2
Public Sub import_data()
    Dim delimiter As String
    Dim contents As String
    Dim lines() As String
    Dim fields() As String
    Dim fnum As Integer
    Dim line_num As Integer
    Dim field_num As Integer
    Dim sql_statement As String

        'delimiter = cboDelimiter.Text
        'If delimiter = "<space>" Then delimiter = " "
        'If delimiter = "<tab>" Then delimiter = vbTab
    
        'hard coding the variables
        delimiter = " "
    
    'Grab the file's contents.
    fnum = FreeFile
    
    'label error handler
    On Error GoTo Error
    
    Open cDlg.fileName For Input As fnum
    contents = Input$(LOF(fnum), #fnum)
    Close #fnum

    'Split the contents into lines.
    lines = Split(contents, vbCrLf)

    Call open_mdb
    With recs
        .Open "Select * from Data"
        If .RecordCount = 0 Then
            'Process the lines and create records.
            For line_num = 25 To UBound(lines)
                'find only lines with letter d
                'If InStr(1, lines(line_num), "D  ", vbTextCompare) Then
                'MsgBox InStr(1, lines(line_num), "D  ", vbTextCompare)
                    
                'find lines that start with d
                'MsgBox Mid$(lines(line_num), 1, 1)
                If Mid$(lines(line_num), 1, 1) = "D" Then
                    'Read a text line.
                    If Len(lines(line_num)) > 0 Then
                        'Build an INSERT statement.
                        sql_statement = "INSERT INTO " & _
                        "Data" & " VALUES ("

                    fields = Split(lines(line_num), delimiter)
                    
                    For field_num = LBound(fields) To UBound(fields)
                        'Add the field to the statement.
                        sql_statement = sql_statement & _
                        "'" & fields(field_num) & "', "
                    Next field_num
            
                    'Remove the last comma.
                    'sql_statement = Left$(sql_statement, Len(sql_statement) - 2) & ")"
                
                    'add survey_id field
                    sql_statement = sql_statement & "'" & fileNameFWD & "')"
                
                    'remove other fields
                    sql_statement = Replace(sql_statement, "'',", "")
            
                    'check is sql statement is correct
                    'MsgBox sql_statement, vbInformation, "Message"
            
                    With con
                        .Execute sql_statement
                    End With
            
                    num_records = num_records + 1
                    End If
                End If
                Next line_num
        Else
            MsgBox "Appending data into the database is now allowed.", vbInformation, "Message"
        End If
    
    End With
    Call close_mdb
    'List1.Text = "Data"
    GoTo Term

Error:
    MsgBox "Error in importing data.", vbInformation, "Message"

Term:

End Sub


'convert coordinates into decimal degrees and match coordinates
Public Sub convert_coordinates()
    Dim coor As String
    Dim deg As String
    Dim min As String
    Dim ddeg As Double
    Dim lat As String
    Dim lon As String
    
    
    Call open_mdb
    With recs
.Open "Select * from Data"
If .RecordCount <> 0 Then
        .MoveFirst
        
        If Mid$(.fields("LAT"), 3, 1) <> "." Then
            While Not .EOF
                coor = .fields("LAT")
                'hardcoding the degrees value
                '2 means first 2 digits is the degrees
                deg = Mid$(coor, 1, 2)
                min = Right$(coor, Len(coor) - 2)
                ddeg = Val(deg) + (Val(min) / 60)
                .fields("LAT") = ddeg
            
                coor = .fields("LON")
                'hardcoding the degrees value
                '2 means first 2 digits is the degrees
                deg = Mid$(coor, 1, 3)
                min = Right$(coor, Len(coor) - 3)
                ddeg = Val(deg) + (Val(min) / 60)
                .fields("LON") = ddeg
                .MoveNext
                DoEvents
            Wend
            MsgBox "Coordinates successfully converted to 'Decimal Degrees'.", vbInformation, "Message"
        Else
            MsgBox "Please check if coordinates are already in 'Decimal Degrees'.", vbInformation, "Message"
        
        End If
    
    
        'match coordinates for the succeeding drops
        On Error Resume Next
        If Check1.Value = 1 Then
            Dim i As Integer
            i = 1
            .MoveFirst
            .MoveNext
                For i = 1 To .RecordCount
                'While Not .EOF
                    lat = .fields("LAT")
                    lon = .fields("LON")
                    .MovePrevious
                    .fields("LAT") = lat
                    .fields("LON") = lon
                    .MoveNext
                    .MoveNext
                    .MoveNext
                'Wend
                Next
            MsgBox "Coordinates successfully matched.", vbInformation, "Message"
        End If
    
Else
    MsgBox "Please check if data has been imported into the database.", vbInformation, "Message"
    GoTo Term
End If
    
    End With
    Call close_mdb

Term:

End Sub


'export data table to csv
Private Function DBExport() As Long
    On Error Resume Next
    Dim x, y As String
    x = "Data"
    'delete invalid characters in arcgis
    y = fileNameFWD
    y = Replace$(y, "-", "_")
    y = Replace$(y, " ", "_")
    y = Replace$(y, ".", "_")
    y = Replace$(y, ",", "_")
    y = Replace$(y, "&", "_")
    y = Replace$(y, "(", "_")
    y = Replace$(y, ")", "")
    
    'delete extra underscores
    y = Replace$(y, "_____", "_")
    y = Replace$(y, "____", "_")
    y = Replace$(y, "___", "_")
    y = Replace$(y, "__", "_")
    
    'append the .csv extension
    y = y & ".csv"
    Kill App.path & "\" & y
    
    Call open_mdb
    With con
        .Execute "SELECT * INTO [Text;Database=" & App.path & ";HDR=Yes;FMT=Delimited].[" & y & "] FROM [" & x & "]", DBExport, adCmdText Or adExecuteNoRecords
    End With
    Call close_mdb
    
    Kill App.path & "\schema.ini"

End Function


'load fwd text file
Private Function LoadFile(dFile As String) As String
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open dFile For Binary As #ff
        LoadFile = Space(LOF(ff))
        Get #ff, , LoadFile
    Close #ff

End Function


'button to add fwd data to the mdb file
Private Sub cmdAddFWDData_Click()
    Dim msg As String
    'get data using analyze fiel
    'AnalyzeFile cDlg.fileName
    
    Call import_data
    List1.Text = "Data"
    msg = Format$(num_records) & " records inserted into the database."
    MsgBox msg, vbInformation, "Message"
    num_records = 0
    
End Sub


'batch export text files to csv
Private Sub cmdBatchExport_Click()
    On Error Resume Next
    Kill App.path & "\" & "*.mdb"
    Kill App.path & "\" & "*.csv"
    
    'create the log file
    Dim LogFile As Integer
    LogFile = FreeFile
        Open App.path & "\" & "Export_Errors.log" For Output As #LogFile
        Print #LogFile, "************************************************************"
        Print #LogFile, "              FWD Export to CSV errors/warnings.            "
        Print #LogFile, "                  File logger by xfuentes.                  "
        Print #LogFile, "************************************************************"
        Print #LogFile, ""
        Print #LogFile, ""
        Print #LogFile, "Files with errors/warnings:"
        Close #LogFile
        
    'count the txt files in the current directory
    Dim e As String
    Dim f, g As Integer
    f = 0
    g = 0
    e = Dir(CurDir() & "\" & "*.fwd")
    Do While e <> ""
    f = f + 1
    e = Dir()
    Loop
    
    'show the txt file count
    'MsgBox f & " TXT files in current directory.", vbInformation, "Message"
    
    
    'doing the progress bar
    ProgressBar1.Max = f
    ProgressBar1.Value = 0
    
    'disbale buttons until exporting is finished
    cmdLoadtxtFile.Enabled = False
    cmdLoadmdbFile.Enabled = False
    cmdAddFWDData.Enabled = False
    cmdConvertCoordinates.Enabled = False
    cmdExportCSV.Enabled = False
    cmdQuickExport.Enabled = False
    cmdBatchExport.Enabled = False
    Check1.Enabled = False
    cmdConvertANSI.Enabled = False
    
        'batch process txt and export to csv
        Dim myfile As String
        myfile = Dir(CurDir() & "\" & "*.fwd")
        
        'check if txt files are present in current directory
        If myfile = "" Then
            GoTo Missing
        Else
            Do While myfile <> ""
            
            cDlg.fileName = myfile
            fileNameFWD = Left$(myfile, Len(myfile) - 4)
            'MsgBox fileNameFWD
            filePathFWD = App.path & "\" & fileNameFWD & ".mdb"
            
            Text1.Text = fileNameFWD
            Text2.Text = filePathFWD
            
            Call create_mdb
            Call display_tbl
            Call import_data
                
                
            'converting the coordinates
            Dim coor As String
            Dim deg As String
            Dim min As String
            Dim ddeg As Double
            Dim lat As String
            Dim lon As String
    
            Call open_mdb
            With recs
                .Open "Select * from Data"
                
                'continue only if there are records
                If .RecordCount <> 0 Then
                    .MoveFirst
                    While Not .EOF
                            coor = .fields("LAT")
                            'hardcoding the degrees value
                            '2 means first 2 digits is the degrees
                            deg = Mid$(coor, 1, 2)
                            min = Right$(coor, Len(coor) - 2)
                            ddeg = Val(deg) + (Val(min) / 60)
                            .fields("LAT") = ddeg
                            
                            coor = .fields("LON")
                            'hardcoding the degrees value
                            '2 means first 2 digits is the degrees
                            deg = Mid$(coor, 1, 3)
                            min = Right$(coor, Len(coor) - 3)
                            ddeg = Val(deg) + (Val(min) / 60)
                            .fields("LON") = ddeg
                            .MoveNext
                            DoEvents
                    Wend
            
                    'match coordinates for the succeeding drops
                    Dim i As Integer
                    i = 1
                    .MoveFirst
                    .MoveNext
                    
                    On Error Resume Next
                    If Check1.Value = 1 Then
                        'MsgBox "You have chosen to match the coordinates.", vbInformation, "Message"
                        For i = 1 To .RecordCount
                        'While Not .EOF
                            lat = .fields("LAT")
                            lon = .fields("LON")
                            .MovePrevious
                            .fields("LAT") = lat
                            .fields("LON") = lon
                            .MoveNext
                            .MoveNext
                            .MoveNext
                        'Wend
                        'List1.Text = "Data"
                        'MsgBox "Coordinates successfully matched.", vbInformation, "Message"
                        Next
                    End If
                    
                    List1.Text = "Data"
                    
                    Call DBExport
                    Label1.Caption = Fix(((ProgressBar1.Value + 1) / ProgressBar1.Max) * 100) & "%" & " completed."
                    Label1.Refresh
                    g = g + 1
            Else
                GoTo Error
                'Open App.path & "\" & "Export_Errors.log" For Append As #LogFile
                'Print #LogFile, fileNameFWD & ".fwd"
                'Close #LogFile
                'GoTo Continue
            End If

            End With
            Call close_mdb
            
            'export only mdb with data
            'Call open_mdb
            'With recs
            '    .Open "Select * from Data"
            '        If .RecordCount <> 0 Then
            '            Call DBExport
            '            Label1.Caption = Fix(((ProgressBar1.Value + 1) / ProgressBar1.Max) * 100) & "%" & " completed."
            '            Label1.Refresh
            '            g = g + 1
            '            'activate only when delete mdb file will work
            '            'GoTo Continue
            '        End If
            'End With
            'Call close_mdb
            '
            'code not working
            'delete mdb file
            'List1.Clear
            'Set DataGrid1.DataSource = Nothing
            'Kill filePathFWD
Continue:
            ProgressBar1.Value = ProgressBar1.Value + 1
            myfile = Dir()
            DoEvents
            Loop

        End If
        GoTo Export
        
Error:
            Close #LogFile
            Open App.path & "\" & "Export_Errors.log" For Append As #LogFile
            Print #LogFile, fileNameFWD & ".fwd"
            Close #LogFile
            GoTo Continue

Export:
'    Label2.Caption = "Batch Export to CSV completed successfully."
'    'replaced with autoclose
    MsgBox "Batch Export to CSV completed successfully.", vbInformation, "Message"
    ProgressBar1.Value = 0
    Label1.Caption = ""
    
    cmdAddFWDData.Enabled = True
    cmdConvertCoordinates.Enabled = True
    cmdExportCSV.Enabled = True
    GoTo Term
    
Missing:
    MsgBox "No FWD files found in the current directory.", vbInformation, "Message"
    GoTo Term2
    
Term:
    
    'putting the summary
    Open App.path & "\" & "Export_Errors.log" For Append As #LogFile
    Print #LogFile, ""
    Print #LogFile, ""
    Print #LogFile, "------------------------------------------------------------"
    Print #LogFile, "Export Summary:"
    Print #LogFile, "   Errors/Warnings:        " & f - g
    Print #LogFile, "   Exported to CSV:        " & g
    Print #LogFile, "   FWD Files Processed:    " & f
    Print #LogFile, "------------------------------------------------------------"
    Close #LogFile
    
    'open the log file on finish
    Shell "notepad " & App.path & "\" & "Export_Errors.log", vbNormalFocus

Term2:
    cmdLoadtxtFile.Enabled = True
    cmdQuickExport.Enabled = True
    cmdLoadmdbFile.Enabled = True
    cmdBatchExport.Enabled = True
    Check1.Enabled = True
    cmdConvertANSI.Enabled = True

End Sub


'form2 for ansi converter
Private Sub cmdConvertANSI_Click()
    Form2.Show vbModal
End Sub


'convert the coordinates into decimal degrees
Private Sub cmdConvertCoordinates_Click()
    Call convert_coordinates
    List1.Text = "Data"
    
End Sub


'export table to csv
Private Sub cmdExportCSV_Click()
    On Error GoTo Error
    Call open_mdb
    With recs
        .Open "Select * from Data"
        If .RecordCount <> 0 Then
            If Mid$(.fields("LAT"), 3, 1) = "." Then
                MsgBox CStr(DBExport()) & " records exported to csv.", vbInformation, "Message"
            Else
                MsgBox "Coordinates are possibly not in 'Decimal Degrees' format.", vbInformation, "Message"
            End If
        End If
    End With
    Call close_mdb
    GoTo Term

Error:
    MsgBox "Data table has no record to export.", vbCritical, "Message"

Term:

End Sub


'load fwd mdb file
Private Sub cmdLoadmdbFile_Click()
    On Error GoTo Error
    CommonDialog1.Filter = "mdb Files | *.mdb"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen

    fileNameFWD = Left$(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
    filePathFWD = CommonDialog1.fileName
    
    Text1.Text = ""
    Text2.Text = filePathFWD
    cDlg.fileName = ""
    
    Call display_tbl
    Call open_mdb
    With recs
        .Open "Select * from Data"
            If .RecordCount <> 0 Then
                Set DataGrid1.DataSource = recs
                List1.Text = "Data"
                'MsgBox "Records found in table.", vbInformation, "Message"
            Else
                List1.Text = "Data"
            End If
    End With
    Call close_mdb
    
    cmdAddFWDData.Enabled = True
    cmdConvertCoordinates.Enabled = True
    cmdExportCSV.Enabled = True

Error:

End Sub

'load fwd txt file and create mdb
Private Sub cmdLoadtxtFile_Click()
    On Error GoTo Error
    cDlg.DefaultExt = "fwd"
    cDlg.Filter = "FWD Files|*.fwd"
    cDlg.CancelError = True
    cDlg.ShowOpen
    
    fileNameFWD = Left$(cDlg.FileTitle, Len(cDlg.FileTitle) - 4)
    filePathFWD = App.path & "\" & fileNameFWD & ".mdb"
    
    On Error Resume Next
    Kill filePathFWD
    Text1.Text = cDlg.fileName
    Text2.Text = filePathFWD
    Call create_mdb
    Call display_tbl
    Set DataGrid1.DataSource = Nothing
    
    cmdAddFWDData.Enabled = True
    cmdConvertCoordinates.Enabled = True
    cmdExportCSV.Enabled = True
    
    GoTo Term
    'analyze text file on load
    'If cDlg.fileName <> "" Then AnalyzeFile cDlg.fileName

Error:
    MsgBox "Error loading FWD file.", vbCritical, "Error"
    
Term:
    
End Sub


'automatically export loaded text file
Private Sub cmdQuickExport_Click()
    On Error GoTo Error
    cDlg.DefaultExt = "fwd"
    cDlg.Filter = "FWD Files|*.fwd"
    cDlg.CancelError = True
    cDlg.ShowOpen
    
    fileNameFWD = Left$(cDlg.FileTitle, Len(cDlg.FileTitle) - 4)
    filePathFWD = App.path & "\" & fileNameFWD & ".mdb"
    
    On Error Resume Next
    Kill filePathFWD
    Text1.Text = cDlg.fileName
    Text2.Text = filePathFWD
    Call create_mdb
    Call display_tbl
    Call refresh_dgrid
    
    cmdAddFWDData.Enabled = True
    cmdConvertCoordinates.Enabled = True
    cmdExportCSV.Enabled = True
    
    Call cmdAddFWDData_Click
    Call cmdConvertCoordinates_Click
    Call cmdExportCSV_Click
    
    GoTo Term
    'analyze text file on load
    'If cDlg.fileName <> "" Then AnalyzeFile cDlg.fileName

Error:
    MsgBox "Error loading FWD file.", vbCritical, "Error"
    
Term:

End Sub


'disable other buttons
Private Sub Form_Load()
    cmdAddFWDData.Enabled = False
    cmdConvertCoordinates.Enabled = False
    cmdExportCSV.Enabled = False

End Sub


'refresh datagrid
Private Sub List1_Click()
    Call refresh_dgrid

End Sub


'mail to aex.gisco@gmail.com
Private Sub Image1_Click()
    ShellExecute hWnd, "open", "mailto: aex.gisco@gmail.com" & vbNullString & vbNullString & vbNullString & vbNullString, vbNullString, vbNullString, SW_SHOW
End Sub


'change pointer on mouse over
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub


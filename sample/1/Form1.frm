VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   4800
   Begin VB.ComboBox cboDelimiter 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1320
      List            =   "Form1.frx":0013
      TabIndex        =   8
      Text            =   "*"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtTable 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "DataValues"
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtDatabaseFile 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "C:\Temp\test.mdb"
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtTextFile 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "C:\Temp\test.txt"
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Delimiter"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Table"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Database File"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Text File"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()
Dim delimiter As String
Dim contents As String
Dim lines() As String
Dim fields() As String
Dim wks As Workspace
Dim db As Database
Dim fnum As Integer
Dim line_num As Integer
Dim field_num As Integer
Dim sql_statement As String
Dim num_records As Long

    delimiter = cboDelimiter.Text
    If delimiter = "<space>" Then delimiter = " "
    If delimiter = "<tab>" Then delimiter = vbTab

    ' Grab the file's contents.
    fnum = FreeFile
    On Error GoTo NoTextFile
    Open txtTextFile.Text For Input As fnum
    contents = Input$(LOF(fnum), #fnum)
    Close #fnum

    ' Split the contents into lines.
    lines = Split(contents, vbCrLf)

    ' Open the database.
    On Error GoTo NoDatabase
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(txtDatabaseFile.Text)
    On Error GoTo 0

    ' Process the lines and create records.
    For line_num = LBound(lines) To UBound(lines)
        ' Read a text line.
        If Len(lines(line_num)) > 0 Then
            ' Build an INSERT statement.
            sql_statement = "INSERT INTO " & _
                txtTable.Text & " VALUES ("

            fields = Split(lines(line_num), delimiter)
            For field_num = LBound(fields) To UBound(fields)
                ' Add the field to the statement.
                sql_statement = sql_statement & _
                    "'" & fields(field_num) & "', "
            Next field_num

            ' Remove the last comma.
            sql_statement = Left$(sql_statement, Len(sql_statement) - 2) & ")"

            MsgBox sql_statement
            
            ' Insert the record.
            On Error GoTo SQLError
            db.Execute sql_statement
            On Error GoTo 0
            num_records = num_records + 1
        End If
    Next line_num

    ' Close the database.
    db.Close
    wks.Close
    MsgBox "Inserted " & Format$(num_records) & " records"
    Exit Sub

NoTextFile:
    MsgBox "Error opening text file."
    Exit Sub

NoDatabase:
    MsgBox "Error opening database."
    Close fnum
    Exit Sub

SQLError:
    MsgBox "Error executing SQL statement '" & _
        sql_statement & "'"
    Close fnum
    db.Close
    wks.Close
    Exit Sub
End Sub
Private Sub Form_Load()
    ' Enter default file and database names.
    txtTextFile.Text = App.Path & "\testdata.txt"
    txtDatabaseFile.Text = App.Path & "\testdata.mdb"
End Sub

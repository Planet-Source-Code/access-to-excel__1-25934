VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CopyToExcel 
   Caption         =   "Copy Database Tables to Excel"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CopyToExcel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Click on a Table to Copy to Excel"
      ForeColor       =   &H00C00000&
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Reselect Database"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   6000
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4875
         TabIndex        =   2
         Top             =   5640
         Width           =   4935
      End
      Begin VB.ListBox List1 
         Height          =   5325
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.mdb"
      Filter          =   "Access Files (*.mdb)"
      FilterIndex     =   1
   End
End
Attribute VB_Name = "CopyToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim dbSR As Database
    Dim rs As Recordset
    Dim strcaption, sn
    Dim Td As TableDef
    Dim i As Single
    Dim Recs As Integer, Counter As Integer
    Dim Barstring As String, MdbFile As String
    Dim Junk As String
    
    Private Type ExlCell
        row As Long
        col As Long
    End Type

Private Sub Form_Load()
    'set blue bar colour
    Picture1.ForeColor = RGB(0, 0, 255)
    
    'open commondialag control
    On Error GoTo errhandler
    CommonDialog1.Filter = "Access Files (*.mdb)"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.FileName = "*.mdb"
    CommonDialog1.ShowOpen
    MdbFile = (CommonDialog1.FileName)

    'set mdb file
    Set dbSR = OpenDatabase(MdbFile)
    
    'populate list with recordsets from selected .mdb
    List1.Clear
    For Each Td In dbSR.TableDefs
        Junk = Td.Name
        Junk = UCase(Junk)
        If Left(Junk, 4) <> "MSYS" Then
            List1.AddItem Td.Name
        End If
    Next
    Frame1.Visible = True

    Exit Sub
errhandler:
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    dbSR.Close
    Set dbSR = Nothing
    End
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
    'open commondialag control
    On Error GoTo errhandler
    CommonDialog1.Filter = "Access Files (*.mdb)"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.FileName = "*.mdb"
    CommonDialog1.ShowOpen
    MdbFile = (CommonDialog1.FileName)
    'set mdb file
    Set dbSR = OpenDatabase(MdbFile)
    'populate list with recordsets from selected .mdb
    List1.Clear
    For Each Td In dbSR.TableDefs
        List1.AddItem Td.Name
    Next
errhandler:
    Frame1.Visible = True
End Sub

Private Sub List1_Click()
    On Error GoTo errortrapper
    Screen.MousePointer = vbHourglass

    Junk = List1.Text
    Set rs = dbSR.OpenRecordset(Junk, dbOpenDynaset)
    Call ToExcel(rs, "C:\wk.xls")
    GoTo skiperrortrapper
errortrapper:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox "This is a system file" & Chr(10) & "and is not accessible."
skiperrortrapper:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CopyRecords(rs As Recordset, ws As Worksheet, _
    StartingCell As ExlCell)
    Dim SomeArray() As Variant
    Dim row As Long, col As Long
    Dim fd As Field
    'Check if rs is not empty
    If rs.EOF And rs.BOF Then Exit Sub
    rs.MoveLast
    ReDim SomeArray(rs.RecordCount + 1, rs.Fields.Count)
    ' Copy column headers to array
    col = 0

    For Each fd In rs.Fields
        SomeArray(0, col) = fd.Name
        col = col + 1
    Next
    ' Copy rs to some array
    rs.MoveFirst
    Recs = rs.RecordCount
    Counter = 0
    
    For row = 1 To rs.RecordCount - 1
        Counter = Counter + 1
        If Counter <= Recs Then i = (Counter / Recs) * 100
        UpdateProgress Picture1, i
        For col = 0 To rs.Fields.Count - 1
            SomeArray(row, col) = rs.Fields(col).Value
            If IsNull(SomeArray(row, col)) Then _
            SomeArray(row, col) = ""
        Next
        rs.MoveNext
    Next
    ' The range should have the same number
    '     of
    ' rows and cols as in the recordset
    ws.Range(ws.Cells(StartingCell.row, StartingCell.col), _
    ws.Cells(StartingCell.row + rs.RecordCount + 1, _
    StartingCell.col + rs.Fields.Count)).Value = SomeArray
End Sub

Private Sub ToExcel(sn As Recordset, strcaption As String)
    Dim oExcel As Object
    Dim objExlSht As Object ' OLE automation object
    Dim stCell As ExlCell

    DoEvents
        On Error Resume Next
        Set oExcel = GetObject(, "Excel.Application")
        ' If Excel is not launched start it
        If Err = 429 Then
            Err = 0
            Set oExcel = CreateObject("Excel.Application")
            ' Can't create object
            If Err = 429 Then
                MsgBox Err & ": " & Error, vbExclamation + vbOKOnly
                Exit Sub
            End If
        End If
        oExcel.Workbooks.Add
        oExcel.Worksheets("sheet1").Name = strcaption
        Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
        stCell.row = 1
        stCell.col = 1
        ' Place the fields across the top of the
        '     spreadsheet:
        CopyRecords sn, objExlSht, stCell
        ' Give the user control
        oExcel.Visible = True
        oExcel.Interactive = True
        ' Clean up:
        Set objExlSht = Nothing ' Remove object variable.
        Set oExcel = Nothing ' Remove object variable.
        Set sn = Nothing ' Remove snapshot object.
    End Sub

Sub UpdateProgress(PB As Control, ByVal percent)
    Dim num$        'use percent
    If Not PB.AutoRedraw Then      'picture in memory ?
        PB.AutoRedraw = -1          'no, make one
    End If
    PB.Cls                      'clear picture in memory
    PB.ScaleWidth = 100         'new sclaemodus
    PB.DrawMode = 10            'not XOR Pen Modus
    num$ = Barstring & Format$(percent, "###") + "%"
    PB.CurrentX = 50 - PB.TextWidth(num$) / 2
    PB.CurrentY = (PB.ScaleHeight - PB.TextHeight(num$)) / 2
    PB.Print num$               'print percent
    PB.Line (0, 0)-(percent, PB.ScaleHeight), , BF
    PB.Refresh          'show difference
End Sub


VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransMaint 
   Caption         =   "Translation Mapping Maintenance"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabMaint 
      Height          =   2505
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   4419
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tables"
      TabPicture(0)   =   "frmTransMaint.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Filters"
      TabPicture(1)   =   "frmTransMaint.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdFilterData"
      Tab(1).Control(1)=   "cmdClearAll"
      Tab(1).Control(2)=   "cmdClearFilter"
      Tab(1).Control(3)=   "cmdAddFilter"
      Tab(1).Control(4)=   "txtFilterValue"
      Tab(1).Control(5)=   "cboFilterCol"
      Tab(1).Control(6)=   "grdFilter"
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label6"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Excel Maintenance"
      TabPicture(2)   =   "frmTransMaint.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdExcelPath"
      Tab(2).Control(1)=   "txtFileImport"
      Tab(2).Control(2)=   "cmdUpload"
      Tab(2).Control(3)=   "cmdTemplate"
      Tab(2).Control(4)=   "Label8"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdExcelPath 
         Caption         =   "..."
         Height          =   405
         Left            =   -69450
         TabIndex        =   28
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txtFileImport 
         Height          =   405
         Left            =   -74790
         TabIndex        =   26
         Top             =   840
         Width           =   5265
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Height          =   405
         Left            =   -68910
         TabIndex        =   25
         Top             =   840
         Width           =   1275
      End
      Begin VB.CommandButton cmdTemplate 
         Caption         =   "Get Template"
         Height          =   405
         Left            =   -68910
         TabIndex        =   24
         Top             =   1470
         Width           =   1275
      End
      Begin VB.CommandButton cmdFilterData 
         Caption         =   "&Filter Data"
         Height          =   465
         Left            =   -71610
         TabIndex        =   18
         Top             =   1830
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clea&r All"
         Height          =   465
         Left            =   -74700
         TabIndex        =   19
         Top             =   1830
         Width           =   945
      End
      Begin VB.CommandButton cmdClearFilter 
         Caption         =   "&Clear Filter"
         Enabled         =   0   'False
         Height          =   465
         Left            =   -73710
         TabIndex        =   20
         Top             =   1830
         Width           =   945
      End
      Begin VB.CommandButton cmdAddFilter 
         Caption         =   "Add F&ilter"
         Height          =   465
         Left            =   -72600
         TabIndex        =   17
         Top             =   1830
         Width           =   945
      End
      Begin VB.TextBox txtFilterValue 
         Height          =   315
         Left            =   -73290
         TabIndex        =   16
         Top             =   1020
         Width           =   2775
      End
      Begin VB.ComboBox cboFilterCol 
         Height          =   315
         Left            =   -73290
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   540
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select Translation to Maintain..."
         Height          =   1965
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7035
         Begin VB.ComboBox cboTable 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   570
            Width           =   4965
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Mapping Table:"
            Height          =   225
            Left            =   150
            TabIndex        =   13
            Top             =   615
            Width           =   1185
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
            Height          =   225
            Left            =   150
            TabIndex        =   12
            Top             =   1140
            Width           =   1185
         End
         Begin VB.Label lblDesc 
            Height          =   435
            Left            =   1470
            TabIndex        =   11
            Top             =   1140
            Width           =   4965
         End
      End
      Begin FPSpread.vaSpread grdFilter 
         Height          =   1935
         Left            =   -70380
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   450
         Width           =   3120
         _Version        =   131077
         _ExtentX        =   5503
         _ExtentY        =   3413
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   8
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   2
         ShadowColor     =   8421440
         ShadowText      =   -2147483634
         SpreadDesigner  =   "frmTransMaint.frx":0054
         StartingColNumber=   0
         UserResize      =   0
         VisibleCols     =   1
      End
      Begin VB.Label Label8 
         Caption         =   "Select Excel Maintenance File:"
         Height          =   285
         Left            =   -74760
         TabIndex        =   27
         Top             =   570
         Width           =   2325
      End
      Begin VB.Label Label5 
         Caption         =   "Valid operators: =, <, >, <=, >=, <>, *"
         Height          =   225
         Left            =   -74460
         TabIndex        =   23
         Top             =   1410
         Width           =   2595
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Where values are:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   21
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Filter data in column:"
         Height          =   225
         Left            =   -74940
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Translations..."
      Height          =   5475
      Left            =   60
      TabIndex        =   4
      Top             =   2670
      Width           =   8715
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4950
         TabIndex        =   7
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7740
         TabIndex        =   2
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5940
         TabIndex        =   1
         Top             =   5040
         Width           =   855
      End
      Begin FPSpread.vaSpread grdTbl 
         Height          =   4725
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   8460
         _Version        =   131077
         _ExtentX        =   14922
         _ExtentY        =   8334
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   0
         MaxRows         =   0
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   2
         ShadowColor     =   8421440
         ShadowText      =   -2147483634
         SpreadDesigner  =   "frmTransMaint.frx":0221
         StartingColNumber=   0
         VisibleCols     =   4
      End
      Begin VB.Label Label4 
         Caption         =   "*Click column headers to sort"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmTransMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'AJN, 10/8/09 removed this combo
'Dim cCboInterface As New cComboBox
Public cCboTable As New cComboBox
Dim iInterfaceID As Integer
Dim iLastRow As Single
Dim iNewRow As Single
Public gsTransMaintMode As String
Public gsTransMaintNewItem
Dim iTMRowCount As Single
Dim sSort As String
Dim iFilterRow As Long 'AJN, 6/13/05 changed to long
Dim iOperStart As Integer 'AJN, 6/1/05
Dim iOperLen As Integer 'AJN, 6/1/05
Dim sOperator As String 'AJN, 6/1/05
Dim sWhere As String 'AJN, 6/1/05

Public Enum TransMaintCols
    TM_COLS_TBL_NAME = 1
    TM_COLS_COL_NAME
    TM_COLS_COL_TYPE
    TM_COLS_COL_LENGTH
    TM_COLS_ALLOW_NULL
    TM_COLS_PRIM_KEY 'AJN 3/23/05
End Enum

Public Enum TE_TRANS_TBL_COLS
    TT_COLS_TBL_NAME = 1
    TT_COLS_TBL_DESC
    TT_COLS_PECOS_SEGMENT
    TT_COLS_PECOS_SEG_CODE
    TT_COLS_PECOS_SEG_DESC
    TT_COLS_PECOS_PARENT_TBL
End Enum

Private Const TM_COL_DEF_COUNT = 6

'AJN, 10/9/09 PECOS Valueset automation
Public sPecosSegmentCode As String
Public sPecosParentTbl As String
Public sPecosSegmentCol As String
Public sPecosDescCol As String
Public sPecosSegmentDesc As String
Public sPecosUpdateTbl As String
Public sPecosParentCode As String
Public sPecosPostableFlag As String
Public sPecosBudgetableFlag As String
Public sInterCompanyFlag As String 'XLVASAMS
Dim iPostableCol As Integer
Dim iBudgetableCol As Integer

'SINC1119485 : Citrix save to share path
'Private Const TEMP_PATH As String = "c:\$r6$\"
'SINC1119485: Changes ends here

Private Const TEMP_FILE_EXT As String = ".xls"

'AJN, 10/8/09 removed this combo
'Private Sub cboInterface_Click()
'
'    iInterfaceID = cCboInterface.ColText(2, cboInterface.ListIndex)
'    Call LoadCombos(cboTable)
'    'AJN, 6/1/05
'    Call ClearFilterTab
'    tabMaint.TabEnabled(1) = False
'
'End Sub

Private Sub cboTable_Click()

    lblDesc.Caption = cCboTable.ColText(2, cboTable.ListIndex)
    sPecosUpdateTbl = cboTable.Text
    iLastRow = 0
    'AJN, 12/4/09
    iPostableCol = 0
    iBudgetableCol = 0
    sSort = "1"
    'AJN, 6/1/05
    Call ClearFilterTab
    Call LoadGrid
    Call GetTableDefinition
    tabMaint.TabEnabled(1) = True
    tabMaint.TabEnabled(2) = True
    txtFileImport.Text = ""
    Call LoadFilterCombo

End Sub

Private Sub cmdAddFilter_Click()
'AJN, 6/1/05 added

    If CheckValidOperator = True Then
        If CheckValidFilter = True Then
            Call AddFilter
            txtFilterValue.SetFocus
        End If
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdClearAll_Click()
'AJN, 6/1/05 added

    grdFilter.MaxRows = 0
    iFilterRow = 0
    sWhere = ""
    cmdClearFilter.Enabled = False

End Sub

Private Sub cmdClearFilter_Click()
'AJN, 6/1/05 added

    Call Grd_DeleteRow(grdFilter, iFilterRow)
    cmdClearFilter.Enabled = False
    iFilterRow = 0

End Sub

Private Sub cmdDelete_Click()

    Dim sWhere As String
    Dim i As Integer
    
    If iLastRow <> 0 And _
    MsgBox("Are you sure you want to delete this translation mapping?", vbQuestion + vbYesNo, "CONFIRM DELETE") = vbYes Then
        'Generate Where clause
        sWhere = ""
        With grdTbl
            For i = 1 To .MaxCols - 3 'exlude change tracking columns
                .row = 0
                .col = i
                sWhere = sWhere & " " & .Text & " = "
                .row = iLastRow
                sWhere = sWhere & "'" & .Text & "' and"
            Next
        End With
        sWhere = Left(sWhere, Len(sWhere) - 4)
        If cboTable.Text = "FIN_CONTRACT_TO_DEPT_V" Or cboTable.Text = "FIN_CLEAN_LICENSEE_V" Then 'N XLVASAMS
            Call objServAcc.Svr_Delete(cboTable.Text, sWhere, "", True)
        Else
            Call objServAcc.Svr_Delete(cboTable.Text, sWhere)
        End If
        'AJN, 10/9/09
        If sPecosSegmentCol <> "" Then
            Call Cmn_SavePecosValueset(sPecosUpdateTbl, "DELETE", sPecosSegmentCode, sPecosSegmentDesc, sPecosParentCode, "N", "N")
        End If
        Call Grd_DeleteRow(grdTbl, iLastRow)
        iLastRow = 0
        Call EnableButtons
    End If

End Sub

Private Sub cmdExcelPath_Click()

    Dim sFile As String
    Dim bError As Boolean
    
    On Error GoTo ErrHandler

    With MDIfrmMain.cdgFile
        .CancelError = True
        .DialogTitle = "Select Excel file to upload"
        '.Filter = "XLS Files (*.xls)|*.xls"
        'CRQ000000226637- Need to allow both xls and xlsx
        '.Filter = "XLSX file (*.xlsx)|*.xlsx"
        .Filter = "XLSX file (*.xlsx)|*.xlsx|XLS Files (*.xls)|*.xls"
        .FilterIndex = 1
        .ShowOpen
        If bError = False Then
            If .FileName <> "" Then
                txtFileImport.Text = .FileName
            End If
            If txtFileImport.Text <> "" Then
                txtFileImport.ToolTipText = txtFileImport.Text
            End If
        End If
    End With

Exit Sub
ErrHandler:
    If err.Number = "32755" Then 'Cancel button selected
        bError = True
        Resume Next
    End If
    
End Sub

Private Sub cmdFilterData_Click()
'AJN, 6/1/05 added

    Dim i As Long 'AJN, 6/13/05 changed to long
    
    sWhere = ""
    If grdFilter.MaxRows > 0 Then
        sWhere = " Where " & Grd_GetGridText(grdFilter, 1, 1)
        For i = 2 To grdFilter.MaxRows
            sWhere = sWhere & " and " & Grd_GetGridText(grdFilter, i, 1)
        Next
        iLastRow = 0
        sWhere = Replace(sWhere, "*", "%")
    End If
    Call LoadGrid

End Sub

Private Sub cmdNew_Click()

    Call Grd_HiLiteRow(grdTbl, iLastRow, False)
    iLastRow = 0
    frmTransMaint.sPecosParentCode = "" ' CRQ000000226637 - Clear the parent code value
    gsTransMaintMode = "Insert"
    frmTransData.Show vbModal
    If gsTransMaintMode <> "Cancel" Then
        Call LoadGrid
        grdTbl.row = iNewRow
        grdTbl.col = 1
        grdTbl.Action = 1
    End If
    Call EnableButtons

End Sub

Private Sub cmdPrint_Click()

    Dim table()
    Dim i As Long
    Dim j As Long

    Screen.MousePointer = vbHourglass
    
    'AJN, 6/13/05 transpose to array to speed up printing to Excel
    ReDim table(1 To grdTbl.MaxRows, 1 To grdTbl.MaxCols)
    For i = 1 To grdTbl.MaxRows
        For j = 1 To grdTbl.MaxCols
            table(i, j) = Grd_GetGridText(grdTbl, i, j)
        Next
    Next
    Call ExcelPrintGrid(cboTable.Text, grdTbl, table())
'    Call ExcelPrintGrid(cboTable.Text, grdTbl)
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdTemplate_Click()

    Dim i As Long
    Dim oRange As Range
    Dim sValue As String

    Screen.MousePointer = vbHourglass

    If oExcelApp Is Nothing Then
        Set oExcelApp = New Excel.Application
        Set oWorkbook = oExcelApp.Workbooks.Add
    Else
        Set oExcelApp = GetObject(, "Excel.Application")
        Set oWorkbook = oExcelApp.Workbooks.Add
    End If
    For i = 1 To GetGridColCount(True)
        sValue = CStr(Grd_GetGridText(grdTbl, 0, i))
        sValue = UCase(sValue)
        Call Exr_SetCellText(1, i, sValue)
    Next
    Set oWorksheet = oWorkbook.ActiveSheet
    oWorksheet.Columns("A:" & Exr_ColLtr(GetGridColCount(True))).NumberFormat = "@"
    Set oRange = oWorksheet.Range("A1:" & Exr_ColLtr(GetGridColCount(True)) & "1")
    oRange.Font.Bold = True
    oWorksheet.Columns.AutoFit
    oExcelApp.Visible = True
    Set oRange = Nothing
    Set oWorkbook = Nothing
    Set oExcelApp = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdUpdate_Click()

    gsTransMaintMode = "Update"
    Call RefreshPecosVariables '' CRQ000000226637 -PecosParentCode is not getting refreshed while clicking update all the time
    frmTransData.Show vbModal
    If gsTransMaintMode <> "Cancel" Then
        Call LoadGrid
        'AJN, 5/11/05
        iLastRow = iNewRow
        grdTbl.row = iLastRow
        grdTbl.col = 1
        grdTbl.Action = 1
        grdTbl.Action = 0
        Call Grd_HiLiteRow(grdTbl, iLastRow, True)
        Call EnableButtons
    End If

End Sub

Private Sub AddFilter()
'AJN, 6/1/05 added
    
    Dim sValue As String
    Dim sFilter As String
    
    sValue = Trim$(Right(txtFilterValue.Text, Len(txtFilterValue.Text) - iOperLen))
    sValue = "'" & sValue & "'"
    If sOperator <> "*" Then
        sFilter = cboFilterCol.Text & " " & sOperator & " " & sValue
    Else
        sFilter = cboFilterCol.Text & " Like " & " " & sValue
    End If
    Call Grd_QuickNR(grdFilter)
    Call Grd_SetGridText(grdFilter, grdFilter.MaxRows, 1, sFilter)

End Sub

Private Function CheckValidOperator() As Boolean
'AJN, 6/1/05 added
    
    Dim sValue As String
    
    CheckValidOperator = False
    txtFilterValue.Text = Trim$(txtFilterValue.Text)
    sValue = txtFilterValue.Text
    iOperStart = 1
    If InStr(1, sValue, "<>") > 0 Then
        iOperLen = 2
        sOperator = "<>"
        CheckValidOperator = True
    ElseIf InStr(1, sValue, ">=") > 0 Then
        iOperLen = 2
        sOperator = ">="
        CheckValidOperator = True
    ElseIf InStr(1, sValue, "<=") > 0 Then
        iOperLen = 2
        sOperator = "<="
        CheckValidOperator = True
    ElseIf InStr(1, sValue, "<") > 0 Then
        iOperLen = 1
        sOperator = "<"
        CheckValidOperator = True
    ElseIf InStr(1, sValue, ">") > 0 Then
        iOperLen = 1
        sOperator = ">"
        CheckValidOperator = True
    ElseIf InStr(1, sValue, "=") > 0 Then
        iOperLen = 1
        sOperator = "="
        CheckValidOperator = True
    ElseIf InStr(1, sValue, "*") > 0 Then
        iOperLen = 0
        sOperator = "*"
        CheckValidOperator = True
    Else
        MsgBox "Please enter a valid operator before adding the filter.", vbInformation + vbOKOnly, "INVALID OPERATOR"
    End If

End Function

Private Function CheckValidFilter() As Boolean

    CheckValidFilter = True
    If cboFilterCol.ListIndex = -1 Then
        MsgBox "Please choose a column to filter on.", vbInformation + vbOKOnly, "NO FILTER COLUMN"
        CheckValidFilter = False
    ElseIf Right(Trim(txtFilterValue.Text), Len(txtFilterValue.Text) - iOperLen) = "" Then
        MsgBox "Please enter a value to filter for after the operator.", vbInformation + vbOKOnly, "INVALID FILTER VALUE"
        CheckValidFilter = False
    End If

End Function

Private Sub ClearFilterTab()
'AJN, 6/1/05 added

    sWhere = ""
    grdFilter.MaxRows = 0
    txtFilterValue.Text = ""
    cboFilterCol.Clear
    cmdClearFilter.Enabled = False

End Sub

Private Sub EnableButtons()

    If iLastRow <> 0 Then
        cmdUpdate.Enabled = True
        If iPostableCol = 0 And iBudgetableCol = 0 Then
            cmdDelete.Enabled = True
        End If
    Else
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
    End If

End Sub

Private Sub cmdUpload_Click()

    If Trim(txtFileImport.Text) = "" Then
        MsgBox "Please select a file to upload.", vbOKOnly, "IMPORT FILE"
    Else
        Call ImportExcelFile
    End If

End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    Call Cmn_Form_Center(Me)
    'AJN, 6/1/05
    tabMaint.TabEnabled(1) = False
    tabMaint.TabEnabled(2) = False
    cmdCancel.Picture = GetIcon(8)
    'AJN, 10/8/09 removed this combo
    'Call cCboInterface.Init(cboInterface, 2)
    Call cCboTable.Init(cboTable, 6)
    Call LoadCombos(cboTable)
    ReDim gTransMaintCols(1 To TM_COL_DEF_COUNT, 1 To 1)
    iTMRowCount = 0
    Call PurgeOldFiles
    Call Sec_FormAccess(Me, False)
    Screen.MousePointer = vbDefault

End Sub

Private Sub PurgeOldFiles()
'AJN, 10/15/09 added

    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    
    'SINC1119485: Citrix save to share path
    Dim TEMP_PATH As String
    TEMP_PATH = GetSetting("LMS32", "Settings", "LogPath")
    If TEMP_PATH = "" Then
    TEMP_PATH = "C:\$r6$\"
    End If
    'SINC1119485: Changes ends here


   
    'Get rid of old PECOS_VIP error files in c:\%r6%
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO.FolderExists(TEMP_PATH) Then
        Set oFolder = oFSO.GetFolder(TEMP_PATH)
        For Each oFile In oFolder.Files
            'If Left(oFile.Name, 9) = "PECOS_VIP" And DateDiff("h", oFile.DateCreated, Now()) > 72 Then ''CRQ000000226637 - To resolve the runtime error which raised due to Datecreated function
            If Left(oFile.Name, 9) = "PECOS_VIP" And DateDiff("h", oFile.DateLastModified, Now()) > 72 Then
                Call oFile.DELETE(True)
            End If
        Next
    End If
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oFSO = Nothing

End Sub

Private Function GetGridColCount(bExcludeAuditCols As Boolean) As Integer

    Dim i As Integer
    Dim iCount As Integer
    
    iCount = 0
    For i = 1 To grdTbl.MaxCols
        If bExcludeAuditCols = True Then
            Select Case UCase(Grd_GetGridText(grdTbl, 0, i))
                Case "DATE_LAST_MOD", "TIME_LAST_MOD", "USER_LAST_MOD"
                    'dont increment col count
                Case Else
                    iCount = iCount + 1
            End Select
        Else
            iCount = iCount + 1
        End If
    Next
    GetGridColCount = iCount

End Function

Private Sub GetTableDefinition()

    Dim sSQL As String
    Dim r()
    Dim bGetData As Boolean
    Dim i As Integer
    Dim j As Integer
        
    bGetData = True
    For i = 1 To UBound(gTransMaintCols, 2)
        If gTransMaintCols(TM_COLS_TBL_NAME, i) = cboTable.Text Then
            bGetData = False 'table definition already loaded
            Exit For
        End If
    Next

    If bGetData Then
        sSQL = "spGetColAttributes '" & cboTable.Text & "'"
        If objServAcc.Svr_SnapShot(sSQL, r()) Then
            ReDim Preserve gTransMaintCols(1 To TM_COL_DEF_COUNT, 1 To iTMRowCount + UBound(r, 1))
            For i = 1 To UBound(r, 1)
                iTMRowCount = iTMRowCount + 1
                gTransMaintCols(TM_COLS_TBL_NAME, iTMRowCount) = r(i, TM_COLS_TBL_NAME)
                gTransMaintCols(TM_COLS_COL_NAME, iTMRowCount) = r(i, TM_COLS_COL_NAME)
                gTransMaintCols(TM_COLS_COL_TYPE, iTMRowCount) = r(i, TM_COLS_COL_TYPE)
                gTransMaintCols(TM_COLS_COL_LENGTH, iTMRowCount) = r(i, TM_COLS_COL_LENGTH)
                gTransMaintCols(TM_COLS_ALLOW_NULL, iTMRowCount) = r(i, TM_COLS_ALLOW_NULL)
                gTransMaintCols(TM_COLS_PRIM_KEY, iTMRowCount) = r(i, TM_COLS_PRIM_KEY)
            Next
        End If
    End If
    'AJN, 10/9/09 PECOS Valueset Automation
    sPecosParentTbl = cCboTable.ColText(TT_COLS_PECOS_PARENT_TBL, frmTransMaint.cboTable.ListIndex)
    If Not IsNull(cCboTable.ColText(TT_COLS_PECOS_SEG_CODE, frmTransMaint.cboTable.ListIndex)) Then
        sPecosSegmentCol = cCboTable.ColText(TT_COLS_PECOS_SEG_CODE, frmTransMaint.cboTable.ListIndex)
        sPecosDescCol = cCboTable.ColText(TT_COLS_PECOS_SEG_DESC, frmTransMaint.cboTable.ListIndex)
    Else
        sPecosSegmentCol = ""
        sPecosDescCol = ""
    End If
    
End Sub

Private Sub ImportExcelFile()
'AJN, 10/12/09 added Excel maintenance function
    
    Dim sFilePath As String
    Dim iTblColCount As Integer
    Dim iExcelColCount As Integer
    Dim lLastRow As Long
    Dim sLastCell As String
    'Dim oExcelApp As Excel.Application
    Dim oSheet As Worksheet
    Dim sName As String
    Dim sTmpFile As String
    Dim iRowsIngested As Long
    Dim sSQL As String
    Dim Results()
    
    Screen.MousePointer = vbHourglass
        
    'SINC1119485: Citrix save to share path
    Dim TEMP_PATH As String
    TEMP_PATH = GetSetting("LMS32", "Settings", "LogPath")
    If TEMP_PATH = "" Then
    TEMP_PATH = "C:\$r6$\"
    End If
    'SINC1119485: Changes ends here
        
    sFilePath = Trim(txtFileImport.Text)
    'Validate path to the file
    If Dir(sFilePath) <> "" Then
        iTblColCount = GetGridColCount(True) 'subtract out the user_last_mod cols (always last 3); set those in code
        'Open the Excel file
        Set oExcelApp = GetExcelApp(sFilePath)
        Set oSheet = oExcelApp.Workbooks(1).ActiveSheet
        iExcelColCount = Exr_GetLastNonEmptyCol(oSheet)
        'Make sure we have the correct number of columns
        If iTblColCount <> iExcelColCount Then
            MsgBox "The upload file selected does not contain the correct number of columns.  Click the 'Get Template' button to get a valid upload template.", vbOKOnly, "INVALID FILE FORMAT"
        Else
            'Add a single-quote in front of every cell to prevent Excel from implicitly converting to numeric
            lLastRow = Exr_GetLastNonEmptyRowIndex(oSheet)
            sLastCell = Exr_ColLtr(iExcelColCount) & lLastRow
            Call Exr_AddSingleQuote(oSheet, "A2", sLastCell) 'assume data range starts at A2
            sName = "PECOS_VIP_" & LoggedInUser.sUserID & "_" & Format(Now(), "yyyymmddhhmmss") & "_" & sPecosUpdateTbl
            sTmpFile = TEMP_PATH & sName & TEMP_FILE_EXT
            'Have to write it bak to disc b/c the Microsoft.Jet provider wants a file path, not an open object
            'Modifed by Manamohana to take only xls file evenif user tries to upload .xlsx file
            'oExcelApp.Workbooks(1).SaveAs FileName:=sTmpFile, FileFormat:=xlExcel9795
            'CRQ000000226637
            If val(oExcelApp.VERSION) < 12 Then
                oExcelApp.Workbooks(1).SaveAs FileName:=sTmpFile, FileFormat:=xlExcel9795
            Else
                oExcelApp.Workbooks(1).SaveAs FileName:=sTmpFile, FileFormat:=56
            End If
            
            Call Exr_CloseExcel(oSheet, oExcelApp)
            iRowsIngested = objServAcc.Svr_CreateSQLtblFromExcel(sTmpFile, sName, "A1", sLastCell, sPecosUpdateTbl)'XLVASAMS Added parameter sPecosUpdateTbl
            If iRowsIngested > 0 Then 'If true, then we had a successful import from Excel
                'Now run the PECOS Valueset Interface Procedure
                'Success will return one column and row with "Success"
                'Failure will return the entire table with an errors column
                'All rows must pass validations in SQL for import to succeed, else generate an excel error file
                sSQL = "spPecosValueset '" & sPecosUpdateTbl & "','UPDATE','','','','','','" & sName & "','" & LoggedInUser.sUserID & "','Y'"
                If objServAcc.Svr_SnapShot(sSQL, Results()) Then
				'XLVASAMS
                  If sPecosUpdateTbl = "FIN_CONTRACT_TO_DEPT_V" Or sPecosUpdateTbl = "FIN_CLEAN_LICENSEE_V" Then
                    objServAcc.SaveDataToProdCopy (sSQL)
                    End If
                     'XLVASAMS
                    Select Case Results(1, 1)
                        Case "0:Success"
                            MsgBox "The Excel file was uploaded successfully.", vbOKOnly, "UPLOAD SUCCESS"
                            Call LoadGrid
                        Case "1:Success"
                            MsgBox "The Excel file was uploaded successfully.  All updates have been submitted to the PECOS Valueset interface.", vbOKOnly, "UPLOAD SUCCESS"
                            Call LoadGrid
                        Case Else
                            MsgBox "The Excel upload failed.  Please correct all errors in the following Excel file and try again.", vbOKOnly, "UPLOAD FAILURE"
                            Call ShowExcelImportErrors(Results())
                    End Select
                End If
            Else
                MsgBox "No rows were imported from Excel.  Please contact MIS.", vbOKOnly, "EXCEL IMPORT ERROR"
            End If
        End If
    Else
        MsgBox "Invalid file name or directory.  Please select a valid file path before uploading.", vbOKOnly, "INVALID FILE"
    End If

    Set oExcelApp = Nothing
    Set oSheet = Nothing
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub ShowExcelImportErrors(Results())

    Dim i As Long
    Dim j As Long
    Dim oRange As Range
    Dim sValue As String

    Screen.MousePointer = vbHourglass
    If oExcelApp Is Nothing Then
        Set oExcelApp = New Excel.Application
        Set oWorkbook = oExcelApp.Workbooks.Add
    Else
        Set oExcelApp = GetObject(, "Excel.Application")
        Set oWorkbook = oExcelApp.Workbooks.Add
    End If
    For i = 1 To GetGridColCount(True)
        sValue = CStr(Grd_GetGridText(grdTbl, 0, i))
        sValue = UCase(sValue)
        Call Exr_SetCellText(1, i, sValue)
    Next
    Set oWorksheet = oWorkbook.ActiveSheet
    Set oRange = oWorksheet.Range("A1:" & Exr_ColLtr(GetGridColCount(True)) & "1")
    oRange.Font.Bold = True
    For i = 1 To UBound(Results, 1)
        For j = 1 To UBound(Results, 2)
            sValue = Results(i, j)
            Call Exr_SetCellText(i + 1, j, sValue, True)
        Next
    Next
    oWorksheet.Columns.AutoFit
    oExcelApp.Visible = True
    Set oRange = Nothing
    Set oWorkbook = Nothing
    Set oExcelApp = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Function GetExcelApp(sFilePath As String) As Excel.Application

    Dim oExcel As Excel.Application

    If oExcel Is Nothing Then
        Set oExcel = New Excel.Application
    Else
        Set oExcel = GetObject(, "Excel.Application")
    End If
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Open txtFileImport
    Set GetExcelApp = oExcel
    Set oExcel = Nothing

End Function

Private Sub LoadCombos(cbo As ComboBox)

    Dim i As Integer
    Dim r()
    Dim sSQL As String
    
    Screen.MousePointer = vbHourglass
    
    Select Case cbo
'AJN, 10/8/09 removed this combo
'        Case cboInterface
'            ssql = "Select interface_desc, interface_id "
'            ssql = ssql & "from TE_INTERFACE_V "
'            ssql = ssql & "order by interface_desc"
        Case cboTable
            sSQL = "Select tbl_name, tbl_desc, PecosSegment, PecosCodeCol, PecosDescCol, PecosParentTbl "
            sSQL = sSQL & "from TE_TRANSLATION_TBL_V "
            'AJN, 10/8/09
            'ssql = ssql & "where interface_id = " & iInterfaceID
            sSQL = sSQL & "where enabled = 'Y' "
            sSQL = sSQL & "order by tbl_name"
            Call cCboTable.Clear
    End Select
    
    If objServAcc.Svr_SnapShot(sSQL, r()) Then
        For i = 1 To UBound(r, 1)
            Select Case cbo
'AJN, 10/8/09 removed this combo
'                Case cboInterface
'                    Call cCboInterface.AddItem(1, r(i, 1), True)
'                    Call cCboInterface.AddItem(2, r(i, 2), False)
'                    If UBound(r, 1) = 1 Then 'only 1 item, select it by default
'                        cboInterface.ListIndex = 0
'                    End If
                Case cboTable
                    Call cCboTable.AddItem(1, r(i, 1), True)
                    Call cCboTable.AddItem(2, r(i, 2), False)
                    Call cCboTable.AddItem(3, r(i, 3), False)
                    Call cCboTable.AddItem(4, r(i, 4), False)
                    Call cCboTable.AddItem(5, r(i, 5), False)
                    Call cCboTable.AddItem(6, r(i, 6), False)
            End Select
        Next
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub LoadFilterCombo()

    Dim i As Integer
    
    cboFilterCol.Clear
    grdTbl.row = 0
    For i = 1 To grdTbl.MaxCols
        grdTbl.col = i
        cboFilterCol.AddItem grdTbl.Value
    Next

End Sub

Private Sub LoadGrid()

    Dim sSQL As String
    Dim rst As ADODB.Recordset
    Dim row()
    Dim i As Single
    Dim j As Single
        
    Screen.MousePointer = vbHourglass
    grdTbl.MaxRows = 0
    grdTbl.ReDraw = False
    Sta_Caption ("Retrieving Translation Table Data...")
    'AJN, 6/1/05 changed
    sSQL = "Select * from " & cboTable.Text
    If sWhere <> "" Then
        sSQL = sSQL & " " & sWhere
    End If
    sSQL = sSQL & " order by " & sSort
    Set rst = objServAcc.Svr_Recordset(sSQL)
    With grdTbl
        .MaxCols = rst.Fields.Count
        'set column headers to table's column names
        For i = 1 To rst.Fields.Count
            .col = i
            .row = 0
            .Text = rst.Fields(i - 1).Name
            .ColWidth(i) = Len(.Text)
        Next
        'now add translation rows
        j = 0
        Do Until rst.EOF
            j = j + 1
            grdTbl.MaxRows = grdTbl.MaxRows + 1
            .row = .MaxRows
            For i = 1 To rst.Fields.Count
                .col = i
                'AJN, 5/3/05
                If IsNull(rst.Fields(i - 1).Value) Then
                    .Text = ""
                Else
                    .Text = rst.Fields(i - 1).Value
                End If
                'AJN, 5/11/05
'                If i = 1 And .Text = gsTransMaintNewItem Then
'                    iNewRow = j
'                End If
            Next
            If gsTransMaintNewItem <> "" Then
                If CheckFindValue(rst) = True Then
                    iNewRow = j
                End If
            End If
            rst.MoveNext
        Loop
        If .MaxRows > 0 Then
            cmdPrint.Enabled = True
        Else
            cmdPrint.Enabled = False
        End If
    End With
    Call EnableButtons
    Sta_Caption ("")
    cmdNew.Enabled = True
    grdTbl.ReDraw = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Function CheckFindValue(rst As ADODB.Recordset) As Boolean

    Dim iFields As Integer
    Dim iPos As Integer
    Dim i As Integer
    Dim sCheckValue As String
    Dim values()
    
    CheckFindValue = False
    sCheckValue = gsTransMaintNewItem
    If InStr(1, sCheckValue, "|") = 0 Then
        If rst.Fields(0) = sCheckValue Then
            CheckFindValue = True
        End If
    Else
        iPos = InStr(1, sCheckValue, "|")
        ReDim values(1 To 1)
        values(1) = Left(sCheckValue, iPos - 1)
        sCheckValue = Right(sCheckValue, Len(sCheckValue) - iPos)
        iFields = 1
        Do Until iPos = 0
            iPos = InStr(iPos, sCheckValue, "|")
            iFields = iFields + 1
            ReDim Preserve values(1 To iFields)
            If iPos = 0 Then 'last field
                values(iFields) = sCheckValue
            Else
                values(iFields) = Left(sCheckValue, iPos - 1)
                sCheckValue = Right(sCheckValue, Len(sCheckValue) - iPos - 1)
            End If
        Loop
        For i = 1 To iFields
            If rst.Fields(i - 1) = values(i) Then
                CheckFindValue = True
            Else
                CheckFindValue = False
                Exit Function
            End If
        Next
        If CheckFindValue = True Then
            Exit Function
        End If
    End If

End Function

Private Sub grdFilter_Click(ByVal col As Long, ByVal row As Long)

    If Grd_IsLit(grdFilter, row) Then
        Call Grd_HiLiteRow(grdFilter, row, False)
        iFilterRow = 0
    ElseIf iFilterRow <> 0 Then
        Call Grd_HiLiteRow(grdFilter, iFilterRow, False)
        iFilterRow = row
        Call Grd_HiLiteRow(grdFilter, row, True)
    Else
        iFilterRow = row
        Call Grd_HiLiteRow(grdFilter, row, True)
    End If
    If iFilterRow <> 0 Then
        iFilterRow = row
        cmdClearFilter.Enabled = True
    Else
        cmdClearFilter.Enabled = False
    End If

End Sub

Private Sub RefreshPecosVariables()
    
    'AJN, 10/8/09 check if there's a PECOS parent
    If sPecosParentTbl <> "" Then
        sPecosSegmentCode = GetPecosSegmentCode(sPecosSegmentCol)
        sPecosSegmentDesc = GetPecosSegmentDesc(sPecosDescCol)
        sPecosParentCode = GetPecosParentCode
    End If
    Call RefreshPecosSegmentFlags
	Call RefreshInterCityFlag 'XLVASAMS
End Sub

Private Sub RefreshPecosSegmentFlags()

    Dim i As Integer
    Dim sGridTxt As String
   
    If iPostableCol = 0 Or iBudgetableCol = 0 Then
        For i = 1 To grdTbl.MaxCols
            sGridTxt = Grd_GetGridText(grdTbl, 0, i)
            If UCase("Pecos_Allow_Posting") = UCase(sGridTxt) Then
                sPecosPostableFlag = Grd_GetGridText(grdTbl, iLastRow, i)
                iPostableCol = i
                cmdDelete.Enabled = False
            ElseIf UCase("Pecos_Allow_Budgeting") = UCase(sGridTxt) Then
                sPecosBudgetableFlag = Grd_GetGridText(grdTbl, iLastRow, i)
                iBudgetableCol = i
                cmdDelete.Enabled = False
            End If
        Next
    Else
        sPecosPostableFlag = Grd_GetGridText(grdTbl, iLastRow, iPostableCol)
        sPecosBudgetableFlag = Grd_GetGridText(grdTbl, iLastRow, iBudgetableCol)
    End If

End Sub

Public Function GetPecosSegmentCode(sPecosSegmentCol As String) As String
   
    Dim i As Integer
    Dim sGridTxt As String
   
    For i = 1 To grdTbl.MaxCols
        sGridTxt = Grd_GetGridText(grdTbl, 0, i)
        If UCase(sPecosSegmentCol) = UCase(sGridTxt) Then
            GetPecosSegmentCode = Grd_GetGridText(grdTbl, iLastRow, i)
            Exit For
        End If
    Next

End Function

Public Function GetPecosSegmentDesc(sPecosDescCol As String) As String

    Dim oCtrl As Control
    
    Dim i As Integer
    Dim sGridTxt As String
   
    For i = 1 To grdTbl.MaxCols
        sGridTxt = Grd_GetGridText(grdTbl, 0, i)
        If UCase(sPecosDescCol) = UCase(sGridTxt) Then
            GetPecosSegmentDesc = Grd_GetGridText(grdTbl, iLastRow, i)
            Exit For
        End If
    Next

End Function

Public Function GetPecosParentCode() As String

    Dim oCtrl As Control
    
    Dim i As Integer
    Dim sGridTxt As String
   
    For i = 1 To grdTbl.MaxCols
        sGridTxt = Grd_GetGridText(grdTbl, 0, i)
        If UCase("Pecos_Parent_Code") = UCase(sGridTxt) Then
            GetPecosParentCode = Grd_GetGridText(grdTbl, iLastRow, i)
            Exit For
        End If
    Next

End Function

Private Sub grdTbl_Click(ByVal col As Long, ByVal row As Long)

    Dim i As Long 'AJN, 6/13/05 changed to long

    Screen.MousePointer = vbHourglass
    If row = 0 And grdTbl.MaxRows <> 0 Then
        Me.Enabled = False
        With grdTbl
            'AJN, 5/11/05
            'Call Grd_HiLiteRow(grdTbl, iLastRow, False)
            'iLastRow = 0
            .col = 1
            .Col2 = .MaxCols
            .row = 1
            .Row2 = .MaxRows
            .SortBy = 0
            .SortKey(1) = col
            .SortKeyOrder(1) = IIf(.SortKeyOrder(1) = 1, 2, 1)
            .Action = 25
            sSort = CStr(col)
            If .SortKeyOrder(1) = 2 Then
                sSort = sSort & " desc"
            End If
            If col <> 1 Then
                .SortKey(2) = 1
                .SortKeyOrder(2) = .SortKeyOrder(1)
                sSort = sSort & ", 1 "
            End If
            'AJN, 5/11/05
            For i = 1 To .MaxRows
                If Grd_IsLit(grdTbl, i) Then
                    iLastRow = i
                    .row = iLastRow
                    .col = 1
                    .Action = 1
                    .Action = 0
                    Call EnableButtons
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next
        End With
    End If
    If Grd_IsLit(grdTbl, row) Then
        Call Grd_HiLiteRow(grdTbl, row, False)
        iLastRow = 0
    ElseIf iLastRow <> 0 Then
        Call Grd_HiLiteRow(grdTbl, iLastRow, False)
        iLastRow = row
        Call Grd_HiLiteRow(grdTbl, row, True)
    Else
        iLastRow = row
        Call Grd_HiLiteRow(grdTbl, row, True)
    End If
    If iLastRow <> 0 Then
        iLastRow = row
    End If
    'AJN, 10/9/09
    Call RefreshPecosVariables
    Call EnableButtons
    Me.Enabled = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub grdTbl_DblClick(ByVal col As Long, ByVal row As Long)
    
    If row <> 0 And grdTbl.MaxRows <> 0 Then
        If iLastRow <> 0 Then
            Call Grd_HiLiteRow(grdTbl, iLastRow, False)
            iLastRow = row
            Call Grd_HiLiteRow(grdTbl, row, True)
        Else
            iLastRow = row
            Call Grd_HiLiteRow(grdTbl, row, True)
        End If
        If iLastRow <> 0 Then
            iLastRow = row
        End If
        'AJN, 10/9/09
        Call RefreshPecosVariables
        Call cmdUpdate_Click
    End If

End Sub

Private Sub txtFilterValue_GotFocus()

    txtFilterValue.SelStart = 0
    txtFilterValue.SelLength = Len(txtFilterValue)
    
End Sub

Private Sub txtFilterValue_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdAddFilter_Click
    End If

End Sub
'XLVASAMS
Private Sub RefreshInterCityFlag()

    Dim i As Integer
    Dim sGridTxt As String
    
     For i = 1 To grdTbl.MaxCols
            sGridTxt = Grd_GetGridText(grdTbl, 0, i)
            If UCase("INTER_COMPANY") = UCase(sGridTxt) Then
                sInterCompanyFlag = Grd_GetGridText(grdTbl, iLastRow, i)
             End If
              Next
   'sPecosPostableFlag = Grd_GetGridText(grdTbl, iLastRow, 0)
   
End Sub
'XLVASAMS
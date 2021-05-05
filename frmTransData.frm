VERSION 5.00
Begin VB.Form frmTransData 
   Caption         =   "Translation Data"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkInterCmpny 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   735
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save Changes"
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame fraControls 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   330
      Width           =   9615
      Begin VB.CheckBox chkBudgetable 
         Height          =   345
         Left            =   2760
         TabIndex        =   10
         Top             =   1860
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox chkPostable 
         Height          =   345
         Left            =   2760
         TabIndex        =   8
         Top             =   1530
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox cboPecosParent 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   690
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.TextBox txt_1_ 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   270
         Width           =   6615
      End
      Begin VB.Label lbl_1_ 
         Alignment       =   1  'Right Justify
         Caption         =   "Column1"
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   2565
      End
   End
   Begin VB.Label lblTblName 
      Caption         =   "Label2"
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Top             =   120
      Width           =   4785
   End
   Begin VB.Label lblTbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Table Name:"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmTransData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const LEFT_OFFSET = 100
Const TOP_OFFSET = 100
Const DEFAULT_HEIGHT = 315
Const DEFAULT_WIDTH = 6600
Const LABEL_WIDTH = 1545

Dim sMode As String
Dim iLastCtlTop As Integer
Dim iCtlCount As Integer
Dim ColDefinition()
Dim sSQLWhereKey As String
Private Const TM_COL_DEF_COUNT = 6
Dim cCboParent As New cComboBox
Dim sPecosParentTextBoxName As String
Dim sPecosPostableCheckBoxName As String
Dim sPecosBudgetableCheckBoxName As String
Dim sInterCompanyCheckBoxName As String 'XLVASAMS

Private Sub cboPecosParent_Click()

    Dim oCtrl As Control
    
    Set oCtrl = Controls(sPecosParentTextBoxName)
    oCtrl.Text = cCboParent.ColText(2, cboPecosParent.ListIndex)

End Sub

Private Sub chkBudgetable_Click()

    Dim oCtrl As Control
    
    Set oCtrl = Controls(sPecosBudgetableCheckBoxName)
    oCtrl.Text = IIf(chkBudgetable.Value = vbChecked, "Y", "N")

End Sub

Private Sub chkPostable_Click()

    Dim oCtrl As Control
    
    Set oCtrl = Controls(sPecosPostableCheckBoxName)
    oCtrl.Text = IIf(chkPostable.Value = vbChecked, "Y", "N")

End Sub
'XLVASAMS
Private Sub chkInterCmpny_Click()
      Dim oCtrl As Control
        Set oCtrl = Controls(sInterCompanyCheckBoxName)
        oCtrl.Text = IIf(chkInterCmpny.Value = vbChecked, "Y", "N")
End Sub
'XLVASAMS

Private Sub cmdCancel_Click()

    frmTransMaint.gsTransMaintMode = "Cancel"
    Unload Me

End Sub

Private Sub cmdOK_Click()

    If ValidateData = True Then
        If SaveData = True Then
            Unload Me
            DoEvents
        Else
            frmTransMaint.gsTransMaintMode = "Cancel"
            Unload Me
            DoEvents
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdOK_Click
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo frmTDerr

    Screen.MousePointer = vbHourglass
    Call Cmn_Form_Center(Me)
    iLastCtlTop = 0
    iCtlCount = 1
    sSQLWhereKey = ""
    sMode = frmTransMaint.gsTransMaintMode
    lblTblName.Caption = frmTransMaint.cboTable.Text
    Call LoadColDefs
    Call ConfigureForm
    Call Sec_FormAccess(Me, False)
    Screen.MousePointer = vbDefault
    
Exit Sub
frmTDerr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "Form_Load")

End Sub

Private Sub LoadColDefs()
    
    Dim i As Integer
    Dim iRowCount As Integer
    
    ReDim ColDefinition(1 To frmTransMaint.grdTbl.MaxCols - 3, 1 To TM_COL_DEF_COUNT) 'exclude change tracking columns
    For i = 1 To UBound(gTransMaintCols, 2)
        If gTransMaintCols(TM_COLS_TBL_NAME, i) = lblTblName.Caption Then
            iRowCount = iRowCount + 1
            ColDefinition(iRowCount, TM_COLS_TBL_NAME) = gTransMaintCols(TM_COLS_TBL_NAME, i)
            ColDefinition(iRowCount, TM_COLS_COL_NAME) = gTransMaintCols(TM_COLS_COL_NAME, i)
            ColDefinition(iRowCount, TM_COLS_COL_TYPE) = gTransMaintCols(TM_COLS_COL_TYPE, i)
            ColDefinition(iRowCount, TM_COLS_COL_LENGTH) = gTransMaintCols(TM_COLS_COL_LENGTH, i)
            ColDefinition(iRowCount, TM_COLS_ALLOW_NULL) = gTransMaintCols(TM_COLS_ALLOW_NULL, i)
            ColDefinition(iRowCount, TM_COLS_PRIM_KEY) = gTransMaintCols(TM_COLS_PRIM_KEY, i)
        End If
    Next

End Sub

Private Sub ConfigureForm()

    Dim i As Integer
    Dim sLblCaption As String
    Dim oCtrl As Control
    
    On Error GoTo ConfigFormErr
    
    cmdOK.Picture = GetIcon(4)
    cmdCancel.Picture = GetIcon(8)
    fraControls.Height = 600
    iLastCtlTop = txt_1_.Top
    txt_1_.MaxLength = ColDefinition(1, TM_COLS_COL_LENGTH)
    With frmTransMaint.grdTbl
        For i = 1 To .MaxCols
            .row = 0
            .col = i
            sLblCaption = .Text
            If i = 1 Then
                lbl_1_.Caption = sLblCaption
            End If
             If Not (sLblCaption = "date_last_mod" Or sLblCaption = "time_last_mod" Or sLblCaption = "user_last_mod" _
                    Or sLblCaption = "Pecos_Parent_Code" Or sLblCaption = "Pecos_Allow_Posting" Or _
                    sLblCaption = "Pecos_Allow_Budgeting" Or sLblCaption = "INTER_COMPANY") Then
                If sMode = "Update" Then
                    .row = .ActiveRow
                    If i > 1 Then
                        Call AddControl(i, "VB.TextBox", sLblCaption, True, .Text, frmTransMaint.sPecosSegmentCol)
                    Else
                        txt_1_.Text = .Text
                        txt_1_.Width = DEFAULT_WIDTH
                        'AJN, 10/16/09 disable primary key field(s) if in UPDATE mode
                        If sMode = "UPDATE" And ColDefinition(i, TM_COLS_PRIM_KEY) = True Then
                            txt_1_.Enabled = False
                            txt_1_.BackColor = vbButtonFace
                        End If
                        'AJN, 10/9/09 need to flag the text box for PECOS segment code if it exists
                        'The PECOS segment code column will be the same as the text box's label caption
                        'since those are the column names of the table being maintained.
                        If UCase(frmTransMaint.sPecosSegmentCol) = UCase(sLblCaption) Then
                            txt_1_.Tag = "PECOS_CODE|" & frmTransMaint.sPecosSegmentCol
                        End If
                    End If
                Else
                    If i > 1 Then
                        Call AddControl(i, "VB.TextBox", sLblCaption, True, , frmTransMaint.sPecosSegmentCol)
                    ElseIf UCase(frmTransMaint.sPecosSegmentCol) = UCase(sLblCaption) Then
                        txt_1_.Tag = "PECOS_CODE|" & frmTransMaint.sPecosSegmentCol
                    End If
                End If
                'AJN, 5/3/05
                fraControls.Height = fraControls.Height + DEFAULT_HEIGHT + TOP_OFFSET - 50
            ElseIf sLblCaption = "Pecos_Parent_Code" Then
                Call AddControl(i, "VB.TextBox", sLblCaption, False)
                'AJN, 10/8/09 Show a combo if there is a PECOS parent
                Call SetupPecosParent(i, frmTransMaint.sPecosParentTbl, frmTransMaint.sPecosSegmentCode)
            'AJN, 12/4/09 Show PECOS Postable and Budgetable flags if a PECOS segment
            ElseIf sLblCaption = "Pecos_Allow_Posting" Or sLblCaption = "Pecos_Allow_Budgeting" Then
                Call AddControl(i, "VB.TextBox", sLblCaption, False)
                Call SetupPecosFlags(i, sLblCaption)
                    'XLVASAMS
             ElseIf sLblCaption = "INTER_COMPANY" Then
               Call AddControl(i, "VB.TextBox", sLblCaption, False)
                Call SetupChkBox(i, sLblCaption)
                     'XLVASAMS
            End If
        Next
    End With
    Me.Height = lblTbl.Height + fraControls.Height + 800
    'AJN, 5/3/05
    Call Me.Move(Me.Left, Me.Top - (iCtlCount * 100))
    cmdOK.TabIndex = iCtlCount + 1
    cmdCancel.TabIndex = cmdOK.TabIndex + 1
    If sMode = "UPDATE" Then
        Call BuildSQLWhere(sSQLWhereKey)
    End If

Exit Sub
ConfigFormErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "ConfigureForm()")

End Sub

Private Function GetPecosSegmentCode(sPecosSegmentCol As String) As String

    Dim oCtrl As Control

    GetPecosSegmentCode = ""
    For Each oCtrl In Me.Controls
    Debug.Print "Control Name: " & oCtrl.Name & " Tag: " & oCtrl.Tag
        If oCtrl.Tag = "PECOS_CODE|" & sPecosSegmentCol Then
            GetPecosSegmentCode = oCtrl.Text
            Exit For
        End If
    Next
    Set oCtrl = Nothing

End Function

Private Function GetPecosSegmentDesc(sPecosDescCol As String) As String

    Dim oCtrl As Control

    GetPecosSegmentDesc = ""
    For Each oCtrl In Me.Controls
        If oCtrl.Tag = "PECOS_DESC|" & sPecosDescCol Then
            GetPecosSegmentDesc = oCtrl.Text
            Exit For
        End If
    Next
    Set oCtrl = Nothing

End Function

Private Sub SetupPecosParent(iCount As Integer, sParentTbl As String, sSegmentCode As String)
'AJN, 10/8/09 PECOS Valueset automation
   
    Dim i As Integer
    Dim sSQL As String
    Dim r()
    Dim oCtrl As Control
    Dim sCtlName As String
    
    sSQL = "spGetPecosParents '" & sParentTbl & "'"
    If objServAcc.Svr_SnapShot(sSQL, r()) Then
        cCboParent.Init cboPecosParent, 2
        For i = 1 To UBound(r, 1)
            Call cCboParent.AddItem(1, r(i, 3), True)
            Call cCboParent.AddItem(2, r(i, 1), False)
        Next
        sCtlName = "lbl" & "_" & CStr(iCount) & "_"
        With cboPecosParent
            .Visible = True
            .Width = DEFAULT_WIDTH
            .Left = txt_1_.Left
            .Top = fraControls.Top + iLastCtlTop + TOP_OFFSET - 50
            .TabIndex = iCtlCount
            iLastCtlTop = .Top
        End With
        fraControls.Height = fraControls.Height + DEFAULT_HEIGHT + TOP_OFFSET - 50
        sPecosParentTextBoxName = "txt_" & iCount & "_"
        If frmTransMaint.sPecosParentCode <> "" Then
            Call cCboParent.FindItem(2, frmTransMaint.sPecosParentCode)
            Set oCtrl = Controls(sPecosParentTextBoxName)
            oCtrl.Text = frmTransMaint.sPecosParentCode
        End If
    Else
        MsgBox "Could not retrieve PECOS parent values for this table.  Please contact MIS.", vbOKOnly, "PECOS PARENT"
    End If
    Set oCtrl = Nothing

End Sub

Private Sub SetupPecosFlags(iCount As Integer, sColName As String)
'AJN, 12/4/09 PECOS Valueset automation
   
    Dim i As Integer
    Dim sSQL As String
    Dim r()
    Dim oCtrl As Control
    Dim sCtlName As String
    
    sCtlName = "lbl" & "_" & CStr(iCount) & "_"
    Select Case sColName
        Case "Pecos_Allow_Posting"
            sPecosPostableCheckBoxName = "txt_" & iCount & "_"
            Set oCtrl = Controls(sPecosPostableCheckBoxName)
            oCtrl.Text = frmTransMaint.sPecosPostableFlag
            Set oCtrl = chkPostable
            oCtrl.Value = IIf(frmTransMaint.sPecosPostableFlag = "Y", vbChecked, vbUnchecked)
        Case "Pecos_Allow_Budgeting"
            sPecosBudgetableCheckBoxName = "txt_" & iCount & "_"
            Set oCtrl = Controls(sPecosBudgetableCheckBoxName)
            oCtrl.Text = frmTransMaint.sPecosBudgetableFlag
            Set oCtrl = chkBudgetable
            oCtrl.Value = IIf(frmTransMaint.sPecosBudgetableFlag = "Y", vbChecked, vbUnchecked)
    End Select
    With oCtrl
        .Visible = True
        .Left = txt_1_.Left
        .Top = fraControls.Top + iLastCtlTop + TOP_OFFSET - 50
        .TabIndex = iCtlCount
        iLastCtlTop = .Top
    End With
    fraControls.Height = fraControls.Height + DEFAULT_HEIGHT + TOP_OFFSET - 50

    Set oCtrl = Nothing

End Sub

Private Sub AddControl(iCount As Integer, sCtlType As String, sLblCaption As String, bVisible As Boolean, Optional sValue As String, Optional sPecosSegmentCol As String)
'Adds a control of sCtlType and a corresponding label to the left of it in fraControls

    Dim sCtlName As String
    Dim oCtrl As Control
    Dim iLblRight As Integer

    On Error GoTo AddControlErr

    iCtlCount = iCtlCount + 1
    'Add the label
    'AJN, 5/6/05
    sCtlName = "lbl" & "_" & CStr(iCount) & "_"
    Set oCtrl = Controls.Add("VB.Label", sCtlName, fraControls)
    With oCtrl
        .Visible = True
        .Height = DEFAULT_HEIGHT
        .Width = lbl_1_.Width 'LABEL_WIDTH
        .Caption = sLblCaption
        .ToolTipText = .Caption
        .Alignment = vbRightJustify
        .Left = lbl_1_.Left 'fraControls.Left + LEFT_OFFSET
        .Top = fraControls.Top + iLastCtlTop + TOP_OFFSET
        .AutoSize = True
        iLblRight = .Left + .Width
    End With
    'Add the control
    sCtlName = "txt_" & CStr(iCount) & "_"
    Set oCtrl = Controls.Add(sCtlType, sCtlName, fraControls)
    With oCtrl
        .Visible = bVisible
        .Height = DEFAULT_HEIGHT
        .Width = DEFAULT_WIDTH
        .Left = txt_1_.Left
        .Top = fraControls.Top + iLastCtlTop + TOP_OFFSET - 50
        .TabIndex = iCtlCount
        If bVisible = True Then
            iLastCtlTop = .Top
        End If
        .MaxLength = ColDefinition(iCount, TM_COLS_COL_LENGTH)
        If sValue <> "" Then
            .Text = Trim(sValue)
        End If
        'AJN, 10/16/09 disable primary key field(s) if in UPDATE mode
        If sMode = "UPDATE" And ColDefinition(iCount, TM_COLS_PRIM_KEY) = True Then
            .Enabled = False
            .BackColor = vbButtonFace
        End If
        'AJN, 10/9/09 need to flag the text box for PECOS segment code if it exists
        'The PECOS segment code column will be the same as the text box's label caption
        'since those are the column names of the table being maintained.
        If UCase(frmTransMaint.sPecosSegmentCol) = UCase(sLblCaption) Then
            .Tag = "PECOS_CODE|" & frmTransMaint.sPecosSegmentCol
        End If
        If UCase(frmTransMaint.sPecosDescCol) = UCase(sLblCaption) Then
            .Tag = "PECOS_DESC|" & frmTransMaint.sPecosDescCol
        End If
    End With
    Set oCtrl = Nothing
    
Exit Sub
AddControlErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "AddControl()")

End Sub

Private Function GetSQLset() As String
    
    Dim oCtrl As Control
    Dim i As Integer
        
    On Error GoTo SQLSetErr
    
    GetSQLset = ColDefinition(1, TM_COLS_COL_NAME) & " = '" & Trim(Replace(txt_1_.Text, "'", "")) & "', "
    For Each oCtrl In Me.Controls
        If Left(oCtrl.Name, 3) <> "cmd" And Left(oCtrl.Name, 3) <> "fra" And Left(oCtrl.Name, 3) <> "lbl" Then
            For i = 2 To iCtlCount
                'AJN, 5/6/05
                If InStr(1, oCtrl.Name, "_" & CStr(i) & "_") > 0 Then
                    GetSQLset = GetSQLset & ColDefinition(i, TM_COLS_COL_NAME) & " = '" & Trim(Replace(oCtrl.Text, "'", "")) & "', " 'assumes varchars
'                ElseIf oCtrl.Name = "cboPecosParent" And oCtrl.Visible = True Then
'                    GetSQLset = GetSQLset & " Pecos_Parent_Code = '" & cCboParent.ColText(2, cboPecosParent.ListIndex) & "', "
'                    Exit For
                End If
            Next
        End If
    Next
    GetSQLset = Left(GetSQLset, Len(GetSQLset) - 2)

Exit Function
SQLSetErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "GetSQLSet()")

End Function

Private Sub BuildSQLWhere(ByRef sSQLWhere As String)
'AJN 3/23/05 - This builds a WHERE clause based on the original
'primary key values brought in for updates
    
    Dim oCtrl As Control
    Dim sCtrlName As String
    Dim i As Integer
    Dim j As Integer
        
    On Error GoTo SQLWhereErr
    
    sSQLWhere = ""
    For j = 1 To iCtlCount
        If ColDefinition(j, TM_COLS_PRIM_KEY) = True Then
            sCtrlName = "txt" & "_" & CStr(j) & "_"
            Call SetControl(oCtrl, sCtrlName)
            'AJN, 6/13/05 parse out apostrophies
            sSQLWhere = sSQLWhere & ColDefinition(j, TM_COLS_COL_NAME) & " = '" & Trim(Replace(oCtrl.Text, "'", "")) & "' and "
        End If
    Next
    If sSQLWhere <> "" Then
        sSQLWhere = Left(sSQLWhere, Len(sSQLWhere) - 4)
    End If

Exit Sub
SQLWhereErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "BuildSQLWhere()")

End Sub
Private Function Getparntnersegment() As String ''CRQ000000226637 - To get SAP_TRADING_PARTNET_V segment value

    Dim oCtrl As Control
    Dim sCtrlName As String
    Dim i As Integer
    Dim j As Integer
        
    On Error GoTo GetparntnersegmentErr
    
    Getparntnersegment = ""
    For j = 1 To iCtlCount
        If ColDefinition(j, TM_COLS_COL_NAME) = "Partner_Cons_Unit" Then
            sCtrlName = "txt" & "_" & CStr(j) & "_"
            Call SetControl(oCtrl, sCtrlName)
            Getparntnersegment = Trim(Replace(oCtrl.Text, "'", ""))
           Exit Function
        End If
    Next
    

Exit Function
GetparntnersegmentErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "Getparntnersegment()")

End Function
Private Sub PopulateInsert(ByRef Insert())

    Dim oCtrl As Control
    Dim i As Integer
        
    On Error GoTo PopInsertErr
        
    ReDim Insert(1 To iCtlCount)
    Insert(1) = Trim(Replace(txt_1_.Text, "'", ""))
    For Each oCtrl In Me.Controls
    'Debug.Print oCtrl.Name
        If Left(oCtrl.Name, 3) <> "cmd" And Left(oCtrl.Name, 3) <> "fra" And Left(oCtrl.Name, 3) <> "lbl" Then
            For i = 2 To iCtlCount
                If InStr(1, oCtrl.Name, "_" & CStr(i) & "_") > 0 Then
                    Insert(i) = Trim(Replace(oCtrl.Text, "'", ""))
'                ElseIf oCtrl.Name = "cboPecosParent" And oCtrl.Visible = True Then
'                    Insert(i) = cCboParent.ColText(2, cboPecosParent.ListIndex)
'                    Exit For
                    Debug.Print "Insert(" & i & "): " & Insert(i)
                End If
            Next
        End If
    Next
    
Exit Sub
PopInsertErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "PopulateInsert()")

End Sub

Private Sub SetControl(ByRef oCtrl As Control, sControlName As String)

    Dim oControl As Control
    
    For Each oControl In Me.Controls
        If oControl.Name = sControlName Then
            Set oCtrl = oControl
        End If
    Next

End Sub

Private Function ValidateData() As Boolean
'AJN, 3/24/05 reworked for primary keys

    Dim oCtrl As Control
    Dim sSQL As String
    Dim sTable As String
    Dim sSQLSet As String
    Dim Results()
    Dim bOK As Boolean
    
    On Error GoTo ValidateDataErr
        
    ValidateData = True
    
    If CheckBlankFields = False Then
        ValidateData = False
        Exit Function 'CRQ000000226637
    ElseIf CheckDuplicateKey = False Then
        ValidateData = False
        Exit Function 'CRQ000000226637
    End If
    
    sTable = lblTblName.Caption
    
     'CRQ000000226637 - Validate Segment length
    If sTable = "SAP_Trading_Partner_V" Then
        sSQL = " EXEC spValidateTransSegment  '" & sTable & "' , '" & ColDefinition(3, TM_COLS_COL_NAME) & "' ,  '" & Getparntnersegment & "'"
    Else
        sSQL = " EXEC spValidateTransSegment  '" & sTable & "' , '" & ColDefinition(1, TM_COLS_COL_NAME) & "' ,  '" & Trim(Replace(txt_1_.Text, "'", "''")) & "'"
    End If
     bOK = objServAcc.Tst_SnapShot(sSQL, Results())
    If bOK Then
      If Results(1) = "1" Then
            MsgBox Results(2), vbCritical + vbOKOnly, "VALIDATION ERROR"
            ValidateData = False
            Exit Function
      End If
    End If
    'Change Ends
    
    sSQLSet = GetSQLset
    
    If sSQLWhereKey = "" Then Call BuildSQLWhere(sSQLWhereKey)  'CRQ000000226637 -  SQLWhereKey is not passed everytime
    
    sSQL = "exec spValidateTableData '" & sTable & "', '" & Replace(sSQLSet, "'", "''") & "', '" & Replace(sSQLWhereKey, "'", "''") & " ' "

    bOK = objServAcc.Tst_SnapShot(sSQL, Results())
    
    If bOK Then
      If Results(1) = "N" Then
            MsgBox Results(2), vbCritical + vbOKOnly, "VALIDATION ERROR"
      End If
    End If
    
Exit Function
ValidateDataErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "ValidateData()")

End Function

Private Function CheckBlankFields() As Boolean
    
    Dim i As Integer
    Dim oCtrl As Control
    Dim sCtlName As String
    
    CheckBlankFields = True
    For i = 1 To iCtlCount
        sCtlName = "txt" & "_" & CStr(i) & "_"
        Call SetControl(oCtrl, sCtlName)
        If oCtrl.Text = "" And ColDefinition(i, TM_COLS_ALLOW_NULL) = False Then
            MsgBox "Please enter a value for " & ColDefinition(i, TM_COLS_COL_NAME) & ".  This field cannot be blank.", vbCritical + vbOKOnly, "VALIDATION ERROR"
            CheckBlankFields = False
            Exit Function
        ElseIf oCtrl.Text = "" And ColDefinition(i, TM_COLS_COL_NAME) = "Pecos_Parent_Code" Then 'CRQ000000226637 - PECOS_Parent_Code is a mandatory for PECOS table
            MsgBox "Please enter a value for " & ColDefinition(i, TM_COLS_COL_NAME) & ".  This field cannot be blank.", vbCritical + vbOKOnly, "VALIDATION ERROR"
            CheckBlankFields = False
            Exit Function
        End If
    Next

End Function

Private Function CheckDuplicateKey() As Boolean
'AJN 3/23/05 reworked

    Dim sSQL As String
    Dim sSQLWhere As String
    Dim r()
    Dim bCheckKey As Boolean
    
    CheckDuplicateKey = True
    Call BuildSQLWhere(sSQLWhere)
    If (sMode = "UPDATE" And sSQLWhere <> sSQLWhereKey) Or sMode = "INSERT" Then 'primary keys have changed, need to check db
        sSQL = "Select * from " & lblTblName.Caption & " where " & sSQLWhere
        If objServAcc.Svr_SnapShot(sSQL, r()) Then
            MsgBox "The value for the field(s) " & Replace(sSQLWhere, "'", "") & "must be unique, and already exist(s) in this table.  Please enter unique values.", vbCritical + vbOKOnly, "DUPLICATE RECORD KEY"
            CheckDuplicateKey = False
        End If
    End If

End Function

Private Function SaveData() As Boolean

    Dim Insert()
    Dim sSQL As String
    Dim oCtrl As Control
    Dim sMPM As String
    Dim sHypExt As String
    Dim sMsg As String
    Dim sTable As String
    Dim sSQLSet As String
    Dim r()
    
    On Error GoTo frmTDSaveDataErr
    
    Screen.MousePointer = vbHourglass
    SaveData = False
    sTable = lblTblName.Caption
    If sMode = "UPDATE" Then
        sSQLSet = GetSQLset
        Sta_Caption ("Updating Database...")
        If sTable = "FIN_CONTRACT_TO_DEPT_V" Or sTable = "FIN_CLEAN_LICENSEE_V" Then 'N XLVASAMS
            Call objServAcc.Svr_Update(sTable, sSQLSet, sSQLWhereKey, True, "", True) 'N XLVASAMS
        Else
            Call objServAcc.Svr_Update(sTable, sSQLSet, sSQLWhereKey)
        End If
    Else
        Call PopulateInsert(Insert())
        Sta_Caption ("Updating Database...")
        If sTable = "FIN_CONTRACT_TO_DEPT_V" Or sTable = "FIN_CLEAN_LICENSEE_V" Then 'N XLVASAMS
            Call objServAcc.Svr_Insert(sTable, Insert(), True, "", True) 'N XLVASAMS
        Else
            Call objServAcc.Svr_Insert(sTable, Insert())
        End If
    End If
    'AJN, 10/9/09 PECOS Valueset automation
    If frmTransMaint.sPecosSegmentCol <> "" Then
        Call SavePecosValueset
    End If
    Call SetFindValues
    Set oCtrl = Nothing
    Sta_Caption ("")
    SaveData = True
    Screen.MousePointer = vbDefault
    
Exit Function
frmTDSaveDataErr:
    Call Err_Error(err.Number, err.Description, "frmTransData", "SaveData()")

End Function

Private Function SavePecosValueset()
'AJN, 10/9/09 added

    Dim sParentCode As String
    Dim sPostable As String
    Dim sBudgetable As String
    
    sParentCode = cCboParent.ColText(2, cboPecosParent.ListIndex)
    sPostable = IIf(chkPostable.Value = vbChecked, "Y", "N")
    sBudgetable = IIf(chkBudgetable.Value = vbChecked, "Y", "N")
    Call Cmn_SavePecosValueset(frmTransMaint.sPecosUpdateTbl, "UPDATE", GetPecosSegmentCode(frmTransMaint.sPecosSegmentCol), GetPecosSegmentDesc(frmTransMaint.sPecosDescCol), sParentCode, sPostable, sBudgetable)

End Function

Private Sub SetFindValues()

    Dim oCtrl As Control
    Dim sCtrlName As String
    Dim j As Integer
    Dim sString As String
        
    For j = 1 To iCtlCount
        If ColDefinition(j, TM_COLS_PRIM_KEY) = True Then
            sCtrlName = "txt" & "_" & CStr(j) & "_"
            Call SetControl(oCtrl, sCtrlName)
            sString = sString & Trim(oCtrl.Text) & "|"
        End If
    Next
    sString = Left(sString, Len(sString) - 1)
    frmTransMaint.gsTransMaintNewItem = sString

End Sub

'XLVASAMS
Private Sub SetupChkBox(iCount As Integer, sColName As String)
'AJN, 12/4/09 PECOS Valueset automation
   
    Dim i As Integer
    Dim sSQL As String
    Dim r()
    Dim oCtrl As Control
    Dim sCtlName As String
    
    sCtlName = "lbl" & "_" & CStr(iCount) & "_"
    'Select Case sColName
       ' Case "Pecos_Allow_Posting"
       If sColName = "INTER_COMPANY" Then
            sInterCompanyCheckBoxName = "txt_" & iCount & "_"
            Set oCtrl = Controls(sInterCompanyCheckBoxName)
            oCtrl.Text = frmTransMaint.sInterCompanyFlag
            Set oCtrl = chkInterCmpny
            oCtrl.Value = IIf(frmTransMaint.sInterCompanyFlag = "Y", vbChecked, vbUnchecked)
            
           '     Case "INTER_COMPANY"
          '  sInterCompanyCheckBoxName = "txt_" & iCount & "_"
           ' Set oCtrl = Controls(sInterCompanyCheckBoxName)
            'oCtrl.Text = frmTransMaint.sInterCompanyFlag
           ' Set oCtrl = chkBudgetable
            'oCtrl.Value = IIf(frmTransMaint.sInterCompanyFlag = "Y", vbChecked, vbUnchecked)
      End If
     
    With oCtrl
        .Visible = True
        .Left = 2890 'txt_1_.Left
        .Top = 2443 'fraControls.Top + iLastCtlTop + TOP_OFFSET - 50
        .TabIndex = iCtlCount
        iLastCtlTop = .Top
    End With
    fraControls.Height = fraControls.Height + DEFAULT_HEIGHT + TOP_OFFSET - 50

    Set oCtrl = Nothing

End Sub
'XLVASAMS
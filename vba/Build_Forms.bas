' =============================================================================
' OCSS Employee Productivity Log
' VBA Module: Build_Forms
'
' PURPOSE
'   Creates all administrative maintenance forms for the database.
'   Run CreateAllForms() once after Build_Schema has completed.
'
' FORMS CREATED
'   frm_Units              - Datasheet: edit unit codes and names
'   frm_TaskTypes          - Datasheet: edit task codes, names, mail flag
'   frm_CertMailTypes      - Datasheet: edit certified mail type codes
'   frm_CertMailStatuses   - Datasheet: edit certified mail delivery statuses
'   frm_ReviewStatuses     - Datasheet: edit supervisor review outcome codes
'   frm_BacklogStatuses    - Datasheet: edit backlog lifecycle status codes
'   frm_Employees          - Continuous form: add/edit employees with combos
'   frm_Admin              - Switchboard: buttons to open all forms above
'
' HOW TO USE
'   1. In the VBA editor (Alt+F11), paste this file into a new module.
'   2. Press F5 with the cursor inside CreateAllForms.
'   3. Close the VBA editor and open frm_Admin from the Navigation Pane.
' =============================================================================

Option Compare Database
Option Explicit

' Layout constants (twips -- 1440 twips = 1 inch)
Private Const FORM_W    As Long = 9360
Private Const COL1_L    As Long = 360
Private Const COL1_W    As Long = 2000
Private Const COL2_L    As Long = 2520
Private Const COL2_W    As Long = 5400
Private Const ROW_H     As Long = 360
Private Const ROW_GAP   As Long = 120
Private Const ROW_START As Long = 360

' =============================================================================
' ENTRY POINT
' =============================================================================

Public Sub CreateAllForms()

    Debug.Print "=== Building Forms: " & Now() & " ==="

    DeleteFormIfExists "frm_Admin"
    DeleteFormIfExists "frm_Employees"
    DeleteFormIfExists "frm_Units"
    DeleteFormIfExists "frm_TaskTypes"
    DeleteFormIfExists "frm_CertMailTypes"
    DeleteFormIfExists "frm_CertMailStatuses"
    DeleteFormIfExists "frm_ReviewStatuses"
    DeleteFormIfExists "frm_BacklogStatuses"

    Build_frm_Units
    Build_frm_TaskTypes
    Build_frm_CertMailTypes
    Build_frm_CertMailStatuses
    Build_frm_ReviewStatuses
    Build_frm_BacklogStatuses
    Build_frm_Employees
    Build_frm_Admin

    Debug.Print "=== Forms Complete: " & Now() & " ==="
    MsgBox "All forms created successfully!" & vbCrLf & vbCrLf & _
           "Open frm_Admin from the Navigation Pane to manage your data.", _
           vbInformation, "Forms Ready"

End Sub

' =============================================================================
' HELPER: delete a form if it already exists (silently)
' =============================================================================

Private Sub DeleteFormIfExists(formName As String)
    On Error Resume Next
    DoCmd.DeleteObject acForm, formName
    Err.Clear
    On Error GoTo 0
End Sub

' =============================================================================
' HELPER: add a label + text box pair to an open form (design view).
' Returns the text box so the caller can set extra properties.
' =============================================================================

Private Function AddLabeledTextBox(formName As String, _
                                   fieldName As String, _
                                   labelCaption As String, _
                                   rowIndex As Integer) As Control
    Dim lbl     As Control
    Dim txt     As Control
    Dim topPos  As Long

    topPos = ROW_START + rowIndex * (ROW_H + ROW_GAP)

    Set lbl = CreateControl(formName, acLabel, acDetail, "", "", _
                            COL1_L, topPos, COL1_W, ROW_H)
    lbl.Caption = labelCaption

    Set txt = CreateControl(formName, acTextBox, acDetail, "", fieldName, _
                            COL2_L, topPos, COL2_W, ROW_H)
    txt.Name = "txt_" & fieldName

    Set AddLabeledTextBox = txt
End Function

' =============================================================================
' HELPER: add a label + check box pair
' =============================================================================

Private Sub AddLabeledCheckBox(formName As String, _
                               fieldName As String, _
                               labelCaption As String, _
                               rowIndex As Integer)
    Dim lbl     As Control
    Dim chk     As Control
    Dim topPos  As Long

    topPos = ROW_START + rowIndex * (ROW_H + ROW_GAP)

    Set lbl = CreateControl(formName, acLabel, acDetail, "", "", _
                            COL1_L, topPos, COL1_W, ROW_H)
    lbl.Caption = labelCaption

    Set chk = CreateControl(formName, acCheckBox, acDetail, "", fieldName, _
                            COL2_L, topPos, 300, ROW_H)
    chk.Name = "chk_" & fieldName
End Sub

' =============================================================================
' HELPER: add a label + combo box pair.
' Returns the combo so the caller can set RowSource etc.
' =============================================================================

Private Function AddLabeledCombo(formName As String, _
                                 fieldName As String, _
                                 labelCaption As String, _
                                 rowIndex As Integer) As Control
    Dim lbl     As Control
    Dim cbo     As Control
    Dim topPos  As Long

    topPos = ROW_START + rowIndex * (ROW_H + ROW_GAP)

    Set lbl = CreateControl(formName, acLabel, acDetail, "", "", _
                            COL1_L, topPos, COL1_W, ROW_H)
    lbl.Caption = labelCaption

    Set cbo = CreateControl(formName, acComboBox, acDetail, "", fieldName, _
                            COL2_L, topPos, COL2_W, ROW_H)
    cbo.Name = "cbo_" & fieldName

    Set AddLabeledCombo = cbo
End Function

' =============================================================================
' LOOKUP FORMS (Datasheet view)
' =============================================================================

Private Sub Build_frm_Units()
    Dim frm As Form
    Dim txt As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "Units"
        .DefaultView       = 3
        .Caption           = "Units"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "UnitCode", 0, 0, 1440, 300)
    txt.Name = "txt_UnitCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "UnitName", 0, 0, 2880, 300)
    txt.Name = "txt_UnitName"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsActive", 0, 0, 300, 300)
    chk.Name = "chk_IsActive"

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_Units"
    Debug.Print "  OK: frm_Units"
End Sub


Private Sub Build_frm_TaskTypes()
    Dim frm As Form
    Dim txt As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "TaskTypes"
        .DefaultView       = 3
        .Caption           = "Task Types"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "TaskCode", 0, 0, 1440, 300)
    txt.Name = "txt_TaskCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "TaskName", 0, 0, 3600, 300)
    txt.Name = "txt_TaskName"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsCertifiedMailTask", 0, 0, 300, 300)
    chk.Name = "chk_IsCertifiedMailTask"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsActive", 0, 0, 300, 300)
    chk.Name = "chk_IsActive"

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_TaskTypes"
    Debug.Print "  OK: frm_TaskTypes"
End Sub


Private Sub Build_frm_CertMailTypes()
    Dim frm As Form
    Dim txt As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "CertifiedMailTypes"
        .DefaultView       = 3
        .Caption           = "Certified Mail Types"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "MailTypeCode", 0, 0, 1440, 300)
    txt.Name = "txt_MailTypeCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "MailTypeName", 0, 0, 3600, 300)
    txt.Name = "txt_MailTypeName"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsActive", 0, 0, 300, 300)
    chk.Name = "chk_IsActive"

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_CertMailTypes"
    Debug.Print "  OK: frm_CertMailTypes"
End Sub


Private Sub Build_frm_CertMailStatuses()
    Dim frm As Form
    Dim txt As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "CertifiedMailStatuses"
        .DefaultView       = 3
        .Caption           = "Certified Mail Statuses"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "StatusCode", 0, 0, 1440, 300)
    txt.Name = "txt_StatusCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "StatusName", 0, 0, 3600, 300)
    txt.Name = "txt_StatusName"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsFinalStatus", 0, 0, 300, 300)
    chk.Name = "chk_IsFinalStatus"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsActive", 0, 0, 300, 300)
    chk.Name = "chk_IsActive"

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_CertMailStatuses"
    Debug.Print "  OK: frm_CertMailStatuses"
End Sub


Private Sub Build_frm_ReviewStatuses()
    Dim frm As Form
    Dim txt As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "ReviewStatuses"
        .DefaultView       = 3
        .Caption           = "Review Statuses"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "StatusCode", 0, 0, 1440, 300)
    txt.Name = "txt_StatusCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "StatusName", 0, 0, 3600, 300)
    txt.Name = "txt_StatusName"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsActive", 0, 0, 300, 300)
    chk.Name = "chk_IsActive"

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_ReviewStatuses"
    Debug.Print "  OK: frm_ReviewStatuses"
End Sub


Private Sub Build_frm_BacklogStatuses()
    Dim frm As Form
    Dim txt As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "BacklogStatuses"
        .DefaultView       = 3
        .Caption           = "Backlog Statuses"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "StatusCode", 0, 0, 1440, 300)
    txt.Name = "txt_StatusCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "StatusName", 0, 0, 3600, 300)
    txt.Name = "txt_StatusName"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsClosedStatus", 0, 0, 300, 300)
    chk.Name = "chk_IsClosedStatus"

    Set chk = CreateControl(frm.Name, acCheckBox, acDetail, "", "IsActive", 0, 0, 300, 300)
    chk.Name = "chk_IsActive"

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_BacklogStatuses"
    Debug.Print "  OK: frm_BacklogStatuses"
End Sub

' =============================================================================
' EMPLOYEES FORM (Continuous form with combo boxes)
' =============================================================================

Private Sub Build_frm_Employees()
    Dim frm    As Form
    Dim txt    As Control
    Dim cbo    As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "Employees"
        .DefaultView       = 1
        .Caption           = "Employees"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    Set txt = AddLabeledTextBox(frm.Name, "EmployeeNumber", "Employee No.:", 0)
    Set txt = AddLabeledTextBox(frm.Name, "LastName", "Last Name:", 1)
    Set txt = AddLabeledTextBox(frm.Name, "FirstName", "First Name:", 2)
    Set txt = AddLabeledTextBox(frm.Name, "EmailAddress", "Email:", 3)

    Set cbo = AddLabeledCombo(frm.Name, "UnitID", "Unit:", 4)
    With cbo
        .RowSourceType = "Table/Query"
        .RowSource     = "SELECT UnitID, UnitName FROM Units ORDER BY UnitName;"
        .ColumnCount   = 2
        .ColumnWidths  = "0;2880"
        .BoundColumn   = 1
        .LimitToList   = True
    End With

    Set cbo = AddLabeledCombo(frm.Name, "SupervisorEmployeeID", "Supervisor:", 5)
    With cbo
        .RowSourceType = "Table/Query"
        .RowSource     = "SELECT EmployeeID, LastName & ', ' & FirstName" & _
                         " FROM Employees ORDER BY LastName, FirstName;"
        .ColumnCount   = 2
        .ColumnWidths  = "0;2880"
        .BoundColumn   = 1
        .LimitToList   = False
    End With

    Set txt = AddLabeledTextBox(frm.Name, "HireDate", "Hire Date:", 6)
    txt.Format = "Short Date"

    AddLabeledCheckBox frm.Name, "IsActive", "Active:", 7

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_Employees"
    Debug.Print "  OK: frm_Employees"
End Sub

' =============================================================================
' ADMIN SWITCHBOARD
' =============================================================================

Private Sub Build_frm_Admin()
    Dim frm      As Form
    Dim lbl      As Control
    Dim btn      As Control
    Dim mdl      As Object
    Dim code     As String
    Dim btnTop   As Long
    Dim i        As Integer

    Dim captions(7) As String
    Dim targets(7)  As String
    Dim btnNames(7) As String

    captions(0) = "Manage Units"
    targets(0)  = "frm_Units"
    btnNames(0) = "btnUnits"

    captions(1) = "Manage Task Types"
    targets(1)  = "frm_TaskTypes"
    btnNames(1) = "btnTaskTypes"

    captions(2) = "Manage Certified Mail Types"
    targets(2)  = "frm_CertMailTypes"
    btnNames(2) = "btnCertMailTypes"

    captions(3) = "Manage Certified Mail Statuses"
    targets(3)  = "frm_CertMailStatuses"
    btnNames(3) = "btnCertMailStatuses"

    captions(4) = "Manage Review Statuses"
    targets(4)  = "frm_ReviewStatuses"
    btnNames(4) = "btnReviewStatuses"

    captions(5) = "Manage Backlog Statuses"
    targets(5)  = "frm_BacklogStatuses"
    btnNames(5) = "btnBacklogStatuses"

    captions(6) = "Manage Employees"
    targets(6)  = "frm_Employees"
    btnNames(6) = "btnEmployees"

    captions(7) = "Close Admin Panel"
    targets(7)  = ""
    btnNames(7) = "btnClose"

    Set frm = CreateForm()
    With frm
        .Caption           = "OCSS Productivity - Admin"
        .DefaultView       = 0
        .ScrollBars        = 0
        .RecordSelectors   = False
        .NavigationButtons = False
        .HasModule         = True
        .Width             = 5400
    End With

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, "", "", 360, 240, 4680, 480)
    lbl.Caption  = "OCSS Productivity Database - Admin"
    lbl.FontSize = 14
    lbl.FontBold = True
    lbl.Name     = "lblTitle"

    btnTop = 1080
    For i = 0 To 7
        Set btn = CreateControl(frm.Name, acCommandButton, acDetail, "", "", _
                                360, btnTop, 4680, 480)
        btn.Name    = btnNames(i)
        btn.Caption = captions(i)
        btn.OnClick = "[Event Procedure]"
        btnTop = btnTop + 600
    Next i

    Set mdl = frm.Module

    On Error Resume Next
    If mdl.CountOfLines > 0 Then
        mdl.DeleteLines 1, mdl.CountOfLines
    End If
    On Error GoTo 0

    code = "Option Compare Database" & vbCrLf
    code = code & "Option Explicit" & vbCrLf & vbCrLf

    For i = 0 To 6
        code = code & "Private Sub " & btnNames(i) & "_Click()" & vbCrLf
        code = code & "    DoCmd.OpenForm """ & targets(i) & """" & vbCrLf
        code = code & "End Sub" & vbCrLf & vbCrLf
    Next i

    code = code & "Private Sub btnClose_Click()" & vbCrLf
    code = code & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    code = code & "End Sub" & vbCrLf

    mdl.AddFromString code

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_Admin"
    Debug.Print "  OK: frm_Admin"
End Sub

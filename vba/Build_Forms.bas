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
Private Const FORM_W    As Long = 9360    ' Form width
Private Const COL1_L    As Long = 360     ' Label left
Private Const COL1_W    As Long = 2000    ' Label width
Private Const COL2_L    As Long = 2520    ' Control left
Private Const COL2_W    As Long = 5400    ' Control width
Private Const ROW_H     As Long = 360     ' Row height
Private Const ROW_GAP   As Long = 120     ' Gap between rows
Private Const ROW_START As Long = 360     ' First row top

' =============================================================================
' ENTRY POINT
' =============================================================================

Public Sub CreateAllForms()

    Debug.Print "=== Building Forms: " & Now() & " ==="

    ' Remove any previous versions so the build is clean
    DeleteFormIfExists "frm_Admin"
    DeleteFormIfExists "frm_Employees"
    DeleteFormIfExists "frm_Units"
    DeleteFormIfExists "frm_TaskTypes"
    DeleteFormIfExists "frm_CertMailTypes"
    DeleteFormIfExists "frm_CertMailStatuses"
    DeleteFormIfExists "frm_ReviewStatuses"
    DeleteFormIfExists "frm_BacklogStatuses"

    ' Build lookup datasheet forms
    Build_frm_Units
    Build_frm_TaskTypes
    Build_frm_CertMailTypes
    Build_frm_CertMailStatuses
    Build_frm_ReviewStatuses
    Build_frm_BacklogStatuses

    ' Build the Employees entry form
    Build_frm_Employees

    ' Build the Admin switchboard last (it references the forms above)
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
' HELPER: add a label + text box pair to an open form (design view)
' Returns the text box control so the caller can adjust properties.
' =============================================================================

Private Function AddLabeledTextBox(formName As String, _
                                   fieldName As String, _
                                   labelCaption As String, _
                                   rowIndex As Integer) As Control
    Dim lbl As Control
    Dim txt As Control
    Dim topPos As Long

    topPos = ROW_START + rowIndex * (ROW_H + ROW_GAP)

    ' Label
    Set lbl = CreateControl(formName, acLabel, acDetail, "", "", _
                            COL1_L, topPos, COL1_W, ROW_H)
    lbl.Caption = labelCaption

    ' Bound text box
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
    Dim lbl As Control
    Dim chk As Control
    Dim topPos As Long

    topPos = ROW_START + rowIndex * (ROW_H + ROW_GAP)

    Set lbl = CreateControl(formName, acLabel, acDetail, "", "", _
                            COL1_L, topPos, COL1_W, ROW_H)
    lbl.Caption = labelCaption

    Set chk = CreateControl(formName, acCheckBox, acDetail, "", fieldName, _
                            COL2_L, topPos, 300, ROW_H)
    chk.Name = "chk_" & fieldName
End Sub

' =============================================================================
' HELPER: add a label + combo box pair
' =============================================================================

Private Function AddLabeledCombo(formName As String, _
                                 fieldName As String, _
                                 labelCaption As String, _
                                 rowIndex As Integer) As Control
    Dim lbl As Control
    Dim cbo As Control
    Dim topPos As Long

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
' LOOKUP FORMS (Datasheet view -- one per lookup table)
' =============================================================================

Private Sub Build_frm_Units()
    Dim frm As Form
    Dim txt As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "Units"
        .DefaultView       = 3        ' Datasheet
        .Caption           = "Units"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    ' Bound controls (position is irrelevant in datasheet view)
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "UnitCode", 0, 0, 1440, 300)
    txt.Name = "txt_UnitCode"

    Set txt = CreateControl(frm.Name, acTextBox, acDetail, "", "UnitName", 0, 0, 2880, 300)
    txt.Name = "txt_UnitName"

    Dim chk As Control
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
    Dim frm As Form
    Dim txt As Control
    Dim cbo As Control
    Dim chk As Control

    Set frm = CreateForm()
    With frm
        .RecordSource      = "Employees"
        .DefaultView       = 1        ' Continuous Forms
        .Caption           = "Employees"
        .AllowAdditions    = True
        .AllowEdits        = True
        .AllowDeletions    = True
        .NavigationButtons = True
        .Width             = FORM_W
    End With

    ' Row 0: Employee Number
    Set txt = AddLabeledTextBox(frm.Name, "EmployeeNumber", "Employee No.:", 0)

    ' Row 1: Last Name
    Set txt = AddLabeledTextBox(frm.Name, "LastName", "Last Name:", 1)

    ' Row 2: First Name
    Set txt = AddLabeledTextBox(frm.Name, "FirstName", "First Name:", 2)

    ' Row 3: Email
    Set txt = AddLabeledTextBox(frm.Name, "EmailAddress", "Email:", 3)

    ' Row 4: Unit (combo box)
    Set cbo = AddLabeledCombo(frm.Name, "UnitID", "Unit:", 4)
    With cbo
        .RowSourceType = "Table/Query"
        .RowSource     = "SELECT UnitID, UnitName FROM Units ORDER BY UnitName;"
        .ColumnCount   = 2
        .ColumnWidths  = "0cm;4cm"
        .BoundColumn   = 1
        .LimitToList   = True
    End With

    ' Row 5: Supervisor (combo box -- self-referencing Employees table)
    Set cbo = AddLabeledCombo(frm.Name, "SupervisorEmployeeID", "Supervisor:", 5)
    With cbo
        .RowSourceType = "Table/Query"
        .RowSource     = "SELECT EmployeeID, LastName & "", "" & FirstName " & _
                         "FROM Employees ORDER BY LastName, FirstName;"
        .ColumnCount   = 2
        .ColumnWidths  = "0cm;4cm"
        .BoundColumn   = 1
        .LimitToList   = False   ' Allow blank (no supervisor)
    End With

    ' Row 6: Hire Date
    Set txt = AddLabeledTextBox(frm.Name, "HireDate", "Hire Date:", 6)
    txt.Format = "Short Date"

    ' Row 7: Is Active
    AddLabeledCheckBox frm.Name, "IsActive", "Active:", 7

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_Employees"
    Debug.Print "  OK: frm_Employees"
End Sub

' =============================================================================
' ADMIN SWITCHBOARD
' Creates a navigation form with one button per maintenance form.
' =============================================================================

Private Sub Build_frm_Admin()
    Dim frm As Form
    Dim lbl As Control
    Dim btn As Control
    Dim mdl As Module
    Dim i As Integer

    ' --- Button definitions: Caption, form to open ---
    Dim captions(7)  As String
    Dim targets(7)   As String
    Dim btnNames(7)  As String

    captions(0) = "Manage Units"                     : targets(0) = "frm_Units"             : btnNames(0) = "btnUnits"
    captions(1) = "Manage Task Types"                : targets(1) = "frm_TaskTypes"          : btnNames(1) = "btnTaskTypes"
    captions(2) = "Manage Certified Mail Types"      : targets(2) = "frm_CertMailTypes"      : btnNames(2) = "btnCertMailTypes"
    captions(3) = "Manage Certified Mail Statuses"   : targets(3) = "frm_CertMailStatuses"   : btnNames(3) = "btnCertMailStatuses"
    captions(4) = "Manage Review Statuses"           : targets(4) = "frm_ReviewStatuses"     : btnNames(4) = "btnReviewStatuses"
    captions(5) = "Manage Backlog Statuses"          : targets(5) = "frm_BacklogStatuses"    : btnNames(5) = "btnBacklogStatuses"
    captions(6) = "Manage Employees"                 : targets(6) = "frm_Employees"          : btnNames(6) = "btnEmployees"
    captions(7) = "Close Admin Panel"               : targets(7) = ""                        : btnNames(7) = "btnClose"

    ' --- Create the form ---
    Set frm = CreateForm()
    With frm
        .Caption          = "OCSS Productivity - Admin"
        .DefaultView      = 0          ' Single Form
        .ScrollBars       = 0          ' Neither
        .RecordSelectors  = False
        .NavigationButtons = False
        .HasModule        = True       ' Enable VBA module for button events
        .Width            = 5400       ' ~3.75 inches wide
    End With

    ' --- Title label ---
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, "", "", _
                            360, 240, 4680, 480)
    lbl.Caption   = "OCSS Productivity Database - Admin"
    lbl.FontSize  = 14
    lbl.FontBold  = True
    lbl.Name      = "lblTitle"

    ' --- One button per entry ---
    Dim btnTop As Long
    btnTop = 1080
    For i = 0 To 7
        Set btn = CreateControl(frm.Name, acCommandButton, acDetail, "", "", _
                                360, btnTop, 4680, 480)
        btn.Name    = btnNames(i)
        btn.Caption = captions(i)
        btn.OnClick = "[Event Procedure]"
        btnTop = btnTop + 600
    Next i

    ' --- Resize the form detail section to fit all buttons ---
    frm.Section(acDetail).Height = btnTop + 360

    ' --- Write click-event procedures into the form module ---
    Set mdl = frm.Module

    ' Remove the default stub Access may have added
    On Error Resume Next
    mdl.DeleteLines 1, mdl.CountOfLines
    On Error GoTo 0

    Dim code As String
    code = "Option Compare Database" & vbCrLf & _
           "Option Explicit" & vbCrLf & vbCrLf

    ' Generate one click sub per form-opening button
    For i = 0 To 6
        code = code & _
               "Private Sub " & btnNames(i) & "_Click()" & vbCrLf & _
               "    DoCmd.OpenForm """ & targets(i) & """" & vbCrLf & _
               "End Sub" & vbCrLf & vbCrLf
    Next i

    ' Close button
    code = code & _
           "Private Sub btnClose_Click()" & vbCrLf & _
           "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
           "End Sub" & vbCrLf

    mdl.AddFromString code

    DoCmd.Close acForm, frm.Name, acSaveYes, "frm_Admin"
    Debug.Print "  OK: frm_Admin"
End Sub

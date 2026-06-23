' =============================================================================
' OCSS Employee Productivity Log
' VBA Module: Build_Forms
'
' HOW TO USE
'   1. Alt+F11 to open VBA Editor
'   2. Insert > Module  (creates a BRAND NEW empty module -- do not reuse
'      the module that contains Build_Schema)
'   3. Paste this entire file into that new module
'   4. Click inside CreateAllForms and press F5
'   5. Close VBA editor, then open frm_Admin from the Navigation Pane
'
' FORMS CREATED
'   frm_Units, frm_TaskTypes, frm_CertMailTypes, frm_CertMailStatuses,
'   frm_ReviewStatuses, frm_BacklogStatuses  -- Datasheet lookup editors
'   frm_Employees  -- Continuous form with Unit and Supervisor combos
'   frm_Admin      -- Switchboard with buttons for all forms above
' =============================================================================

Option Compare Database
Option Explicit

' Access AcControlType numeric values (avoids enum-resolution compile issues)
' 100=Label  104=CommandButton  106=CheckBox  109=TextBox  111=ComboBox
' Access AcObjectType: 2=Form
' Access AcCloseSave:  1=SaveYes

' =============================================================================
' PUBLIC CALLBACK FUNCTIONS
' frm_Admin buttons call these via  OnClick = "=Btn_XXX()"
' They must live in this standard module.
' =============================================================================

Public Function Btn_Units() As Integer
    DoCmd.OpenForm "frm_Units"
End Function

Public Function Btn_TaskTypes() As Integer
    DoCmd.OpenForm "frm_TaskTypes"
End Function

Public Function Btn_CertMailTypes() As Integer
    DoCmd.OpenForm "frm_CertMailTypes"
End Function

Public Function Btn_CertMailStatuses() As Integer
    DoCmd.OpenForm "frm_CertMailStatuses"
End Function

Public Function Btn_ReviewStatuses() As Integer
    DoCmd.OpenForm "frm_ReviewStatuses"
End Function

Public Function Btn_BacklogStatuses() As Integer
    DoCmd.OpenForm "frm_BacklogStatuses"
End Function

Public Function Btn_Employees() As Integer
    DoCmd.OpenForm "frm_Employees"
End Function

Public Function Btn_Close() As Integer
    DoCmd.Close 2, "frm_Admin"
End Function

' =============================================================================
' ENTRY POINT
' =============================================================================

Public Sub CreateAllForms()

    Debug.Print "=== Building Forms: " & Now() & " ==="

    DropForm "frm_Admin"
    DropForm "frm_Employees"
    DropForm "frm_Units"
    DropForm "frm_TaskTypes"
    DropForm "frm_CertMailTypes"
    DropForm "frm_CertMailStatuses"
    DropForm "frm_ReviewStatuses"
    DropForm "frm_BacklogStatuses"

    Make_frm_Units
    Make_frm_TaskTypes
    Make_frm_CertMailTypes
    Make_frm_CertMailStatuses
    Make_frm_ReviewStatuses
    Make_frm_BacklogStatuses
    Make_frm_Employees
    Make_frm_Admin

    Debug.Print "=== Done: " & Now() & " ==="
    MsgBox "All forms created!" & vbCrLf & vbCrLf & _
           "Open frm_Admin from the Navigation Pane.", _
           vbInformation, "Done"

End Sub

Private Sub DropForm(fName As String)
    On Error Resume Next
    DoCmd.DeleteObject 2, fName
    Err.Clear
    On Error GoTo 0
End Sub

' =============================================================================
' LOOKUP TABLE FORMS  (Datasheet view)
' CreateControl args: formName, controlType, section, parent, column,
'                     left, top, width, height
' =============================================================================

Private Sub Make_frm_Units()
    Dim frm As Object
    Dim ctl As Object
    Set frm = CreateForm()
    frm.RecordSource = "Units"
    frm.DefaultView = 3
    frm.Caption = "Units"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    Set ctl = CreateControl(frm.Name, 109, 0, "", "UnitCode", 0, 0, 1440, 300)
    ctl.Name = "txt_UnitCode"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "UnitName", 0, 0, 2880, 300)
    ctl.Name = "txt_UnitName"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 0, 0, 300, 300)
    ctl.Name = "chk_IsActive"
    DoCmd.Close 2, frm.Name, 1, "frm_Units"
    Debug.Print "  OK: frm_Units"
End Sub

Private Sub Make_frm_TaskTypes()
    Dim frm As Object
    Dim ctl As Object
    Set frm = CreateForm()
    frm.RecordSource = "TaskTypes"
    frm.DefaultView = 3
    frm.Caption = "Task Types"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    Set ctl = CreateControl(frm.Name, 109, 0, "", "TaskCode", 0, 0, 1440, 300)
    ctl.Name = "txt_TaskCode"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "TaskName", 0, 0, 3600, 300)
    ctl.Name = "txt_TaskName"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsCertifiedMailTask", 0, 0, 300, 300)
    ctl.Name = "chk_IsCertifiedMailTask"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 0, 0, 300, 300)
    ctl.Name = "chk_IsActive"
    DoCmd.Close 2, frm.Name, 1, "frm_TaskTypes"
    Debug.Print "  OK: frm_TaskTypes"
End Sub

Private Sub Make_frm_CertMailTypes()
    Dim frm As Object
    Dim ctl As Object
    Set frm = CreateForm()
    frm.RecordSource = "CertifiedMailTypes"
    frm.DefaultView = 3
    frm.Caption = "Certified Mail Types"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    Set ctl = CreateControl(frm.Name, 109, 0, "", "MailTypeCode", 0, 0, 1440, 300)
    ctl.Name = "txt_MailTypeCode"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "MailTypeName", 0, 0, 3600, 300)
    ctl.Name = "txt_MailTypeName"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 0, 0, 300, 300)
    ctl.Name = "chk_IsActive"
    DoCmd.Close 2, frm.Name, 1, "frm_CertMailTypes"
    Debug.Print "  OK: frm_CertMailTypes"
End Sub

Private Sub Make_frm_CertMailStatuses()
    Dim frm As Object
    Dim ctl As Object
    Set frm = CreateForm()
    frm.RecordSource = "CertifiedMailStatuses"
    frm.DefaultView = 3
    frm.Caption = "Certified Mail Statuses"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    Set ctl = CreateControl(frm.Name, 109, 0, "", "StatusCode", 0, 0, 1440, 300)
    ctl.Name = "txt_StatusCode"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "StatusName", 0, 0, 3600, 300)
    ctl.Name = "txt_StatusName"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsFinalStatus", 0, 0, 300, 300)
    ctl.Name = "chk_IsFinalStatus"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 0, 0, 300, 300)
    ctl.Name = "chk_IsActive"
    DoCmd.Close 2, frm.Name, 1, "frm_CertMailStatuses"
    Debug.Print "  OK: frm_CertMailStatuses"
End Sub

Private Sub Make_frm_ReviewStatuses()
    Dim frm As Object
    Dim ctl As Object
    Set frm = CreateForm()
    frm.RecordSource = "ReviewStatuses"
    frm.DefaultView = 3
    frm.Caption = "Review Statuses"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    Set ctl = CreateControl(frm.Name, 109, 0, "", "StatusCode", 0, 0, 1440, 300)
    ctl.Name = "txt_StatusCode"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "StatusName", 0, 0, 3600, 300)
    ctl.Name = "txt_StatusName"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 0, 0, 300, 300)
    ctl.Name = "chk_IsActive"
    DoCmd.Close 2, frm.Name, 1, "frm_ReviewStatuses"
    Debug.Print "  OK: frm_ReviewStatuses"
End Sub

Private Sub Make_frm_BacklogStatuses()
    Dim frm As Object
    Dim ctl As Object
    Set frm = CreateForm()
    frm.RecordSource = "BacklogStatuses"
    frm.DefaultView = 3
    frm.Caption = "Backlog Statuses"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    Set ctl = CreateControl(frm.Name, 109, 0, "", "StatusCode", 0, 0, 1440, 300)
    ctl.Name = "txt_StatusCode"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "StatusName", 0, 0, 3600, 300)
    ctl.Name = "txt_StatusName"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsClosedStatus", 0, 0, 300, 300)
    ctl.Name = "chk_IsClosedStatus"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 0, 0, 300, 300)
    ctl.Name = "chk_IsActive"
    DoCmd.Close 2, frm.Name, 1, "frm_BacklogStatuses"
    Debug.Print "  OK: frm_BacklogStatuses"
End Sub

' =============================================================================
' EMPLOYEES FORM  (Continuous, with combo boxes for Unit and Supervisor)
' Row layout: 480 twips per row (360 height + 120 gap), starting at top=360
' =============================================================================

Private Sub Make_frm_Employees()
    Dim frm As Object
    Dim ctl As Object
    Dim rowTop As Long

    Set frm = CreateForm()
    frm.RecordSource = "Employees"
    frm.DefaultView = 1
    frm.Caption = "Employees"
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    frm.NavigationButtons = True
    frm.Width = 9360

    rowTop = 360
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Employee No.:"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "EmployeeNumber", 2520, rowTop, 5400, 360)
    ctl.Name = "txt_EmployeeNumber"

    rowTop = 840
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Last Name:"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "LastName", 2520, rowTop, 5400, 360)
    ctl.Name = "txt_LastName"

    rowTop = 1320
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "First Name:"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "FirstName", 2520, rowTop, 5400, 360)
    ctl.Name = "txt_FirstName"

    rowTop = 1800
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Email:"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "EmailAddress", 2520, rowTop, 5400, 360)
    ctl.Name = "txt_EmailAddress"

    rowTop = 2280
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Unit:"
    Set ctl = CreateControl(frm.Name, 111, 0, "", "UnitID", 2520, rowTop, 5400, 360)
    ctl.Name = "cbo_UnitID"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT UnitID, UnitName FROM Units ORDER BY UnitName;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;2880"
    ctl.BoundColumn = 1
    ctl.LimitToList = True

    rowTop = 2760
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Supervisor:"
    Set ctl = CreateControl(frm.Name, 111, 0, "", "SupervisorEmployeeID", 2520, rowTop, 5400, 360)
    ctl.Name = "cbo_Supervisor"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT EmployeeID, LastName & Chr(44) & Chr(32) & FirstName" & _
                   " FROM Employees ORDER BY LastName, FirstName;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;2880"
    ctl.BoundColumn = 1
    ctl.LimitToList = False

    rowTop = 3240
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Hire Date:"
    Set ctl = CreateControl(frm.Name, 109, 0, "", "HireDate", 2520, rowTop, 5400, 360)
    ctl.Name = "txt_HireDate"
    ctl.Format = "Short Date"

    rowTop = 3720
    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, rowTop, 2000, 360)
    ctl.Caption = "Active:"
    Set ctl = CreateControl(frm.Name, 106, 0, "", "IsActive", 2520, rowTop, 300, 360)
    ctl.Name = "chk_IsActive"

    DoCmd.Close 2, frm.Name, 1, "frm_Employees"
    Debug.Print "  OK: frm_Employees"
End Sub

' =============================================================================
' ADMIN SWITCHBOARD
' Buttons use OnClick = "=Btn_XXX()" to call public functions in this module.
' No form-module code writing required.
' =============================================================================

Private Sub Make_frm_Admin()
    Dim frm   As Object
    Dim ctl   As Object
    Dim btnTop As Long
    Dim i     As Integer

    Dim caps(7)   As String
    Dim clicks(7) As String

    caps(0) = "Manage Units"                   : clicks(0) = "=Btn_Units()"
    caps(1) = "Manage Task Types"              : clicks(1) = "=Btn_TaskTypes()"
    caps(2) = "Manage Certified Mail Types"    : clicks(2) = "=Btn_CertMailTypes()"
    caps(3) = "Manage Certified Mail Statuses" : clicks(3) = "=Btn_CertMailStatuses()"
    caps(4) = "Manage Review Statuses"         : clicks(4) = "=Btn_ReviewStatuses()"
    caps(5) = "Manage Backlog Statuses"        : clicks(5) = "=Btn_BacklogStatuses()"
    caps(6) = "Manage Employees"               : clicks(6) = "=Btn_Employees()"
    caps(7) = "Close Admin Panel"              : clicks(7) = "=Btn_Close()"

    Set frm = CreateForm()
    frm.Caption = "OCSS Productivity - Admin"
    frm.DefaultView = 0
    frm.ScrollBars = 0
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.Width = 5400

    Set ctl = CreateControl(frm.Name, 100, 0, "", "", 360, 240, 4680, 480)
    ctl.Caption = "OCSS Productivity - Admin Panel"
    ctl.FontSize = 14
    ctl.FontBold = True
    ctl.Name = "lblTitle"

    btnTop = 1080
    For i = 0 To 7
        Set ctl = CreateControl(frm.Name, 104, 0, "", "", 360, btnTop, 4680, 480)
        ctl.Caption = caps(i)
        ctl.OnClick = clicks(i)
        btnTop = btnTop + 600
    Next i

    DoCmd.Close 2, frm.Name, 1, "frm_Admin"
    Debug.Print "  OK: frm_Admin"
End Sub


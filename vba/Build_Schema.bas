' =============================================================================
' OCSS Employee Productivity Log
' VBA Module: Build_Schema
'
' PURPOSE
'   Executes all DDL and seed-data SQL scripts in the correct order to build
'   the complete database schema from scratch.
'
' HOW TO USE
'   1. Open your blank Access database (.accdb).
'   2. Press Alt + F11 to open the VBA Editor.
'   3. Insert > Module, then paste the entire contents of this file.
'   4. Press F5 (or Run > Run Sub) while the cursor is inside RunAllScripts.
'   5. Check the Immediate Window (Ctrl + G) for success/error messages.
'
' NOTES
'   - Run this only ONCE on a fresh database.
'   - If a statement fails, the error and statement text are printed to the
'     Immediate Window so you can diagnose and re-run just that statement
'     manually in the Access Query Editor.
'   - All DDL statements are run with dbFailOnError so the procedure stops
'     on the first unhandled error.
' =============================================================================

Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' Entry point -- call this to build everything in order.
' -----------------------------------------------------------------------------
Public Sub RunAllScripts()

    Debug.Print "=== OCSS Productivity DB Build Started: " & Now() & " ==="

    Debug.Print "--- Step 1: Lookup Tables ---"
    Call Create_LookupTables

    Debug.Print "--- Step 2: Employees ---"
    Call Create_Employees

    Debug.Print "--- Step 3: Reporting Months ---"
    Call Create_ReportingMonths

    Debug.Print "--- Step 4: Core Transactional Tables ---"
    Call Create_CoreTables

    Debug.Print "--- Step 5: Deferred / Circular Constraints ---"
    Call Add_DeferredConstraints

    Debug.Print "--- Step 6: Indexes ---"
    Call Create_Indexes

    Debug.Print "--- Step 7: Seed Data ---"
    Call Insert_SeedData

    Debug.Print "=== Build Complete: " & Now() & " ==="
    MsgBox "Database schema and seed data created successfully!", vbInformation, "Build Complete"

End Sub


' -----------------------------------------------------------------------------
' RESET: drops all tables so RunAllScripts can be re-run on the same database.
' Call this if you get "Table already exists" errors.
' WARNING: this permanently deletes all data. Use only on a dev/test database.
' -----------------------------------------------------------------------------
Public Sub ResetDatabase()

    Dim answer As VbMsgBoxResult
    answer = MsgBox("This will DELETE all tables and data permanently." & vbCrLf & _
                    "Are you sure you want to reset the database?", _
                    vbYesNo + vbCritical, "Confirm Reset")
    If answer <> vbYes Then
        Debug.Print "Reset cancelled."
        Exit Sub
    End If

    Debug.Print "=== Reset Started: " & Now() & " ==="

    ' Step 1: Remove deferred/circular constraints so tables can be dropped cleanly.
    RunSQLSafe "ALTER TABLE BacklogItems DROP CONSTRAINT FK_BacklogItems_CreatedFromDetail;", _
               "Drop FK_BacklogItems_CreatedFromDetail"
    RunSQLSafe "ALTER TABLE Employees DROP CONSTRAINT FK_Employees_Supervisor;", _
               "Drop FK_Employees_Supervisor"

    ' Step 2: Drop tables in reverse-dependency order.
    RunSQLSafe "DROP TABLE SupervisorReviews;",      "Drop SupervisorReviews"
    RunSQLSafe "DROP TABLE ProductivityLogDetails;", "Drop ProductivityLogDetails"
    RunSQLSafe "DROP TABLE BacklogItems;",           "Drop BacklogItems"
    RunSQLSafe "DROP TABLE ProductivityLogs;",       "Drop ProductivityLogs"
    RunSQLSafe "DROP TABLE ReportingMonths;",        "Drop ReportingMonths"
    RunSQLSafe "DROP TABLE Employees;",              "Drop Employees"
    RunSQLSafe "DROP TABLE BacklogStatuses;",        "Drop BacklogStatuses"
    RunSQLSafe "DROP TABLE ReviewStatuses;",         "Drop ReviewStatuses"
    RunSQLSafe "DROP TABLE CertifiedMailStatuses;",  "Drop CertifiedMailStatuses"
    RunSQLSafe "DROP TABLE CertifiedMailTypes;",     "Drop CertifiedMailTypes"
    RunSQLSafe "DROP TABLE TaskTypes;",              "Drop TaskTypes"
    RunSQLSafe "DROP TABLE Units;",                  "Drop Units"

    Debug.Print "=== Reset Complete. Run RunAllScripts to rebuild. ==="
    MsgBox "All tables dropped. Run RunAllScripts to rebuild the database.", _
           vbInformation, "Reset Complete"

End Sub


' -----------------------------------------------------------------------------
' Helper: executes a single DDL or DML statement; prints result to Immediate.
' -----------------------------------------------------------------------------
Private Sub RunSQL(sql As String, description As String)
    On Error GoTo Err_Handler
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "  OK: " & description
    Exit Sub

Err_Handler:
    Debug.Print "  ERROR in [" & description & "]: " & Err.Description
    Debug.Print "  SQL: " & sql
    ' Re-raise so RunAllScripts stops -- remove this line to continue on errors.
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


' -----------------------------------------------------------------------------
' Helper: like RunSQL but swallows errors -- used by ResetDatabase so missing
' tables or constraints don't abort the drop sequence.
' -----------------------------------------------------------------------------
Private Sub RunSQLSafe(sql As String, description As String)
    On Error Resume Next
    CurrentDb.Execute sql, dbFailOnError
    If Err.Number <> 0 Then
        Debug.Print "  SKIP (not found): " & description
        Err.Clear
    Else
        Debug.Print "  OK: " & description
    End If
    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------------
' Step 1 -- Lookup / Reference Tables
' -----------------------------------------------------------------------------
Private Sub Create_LookupTables()

    RunSQL "CREATE TABLE Units (" & _
           "    UnitID    AUTOINCREMENT CONSTRAINT PK_Units PRIMARY KEY," & _
           "    UnitCode  TEXT(20)  NOT NULL," & _
           "    UnitName  TEXT(100) NOT NULL," & _
           "    IsActive  YESNO     NOT NULL," & _
           "    CONSTRAINT UQ_Units_UnitCode UNIQUE (UnitCode)" & _
           ");", "CREATE TABLE Units"

    RunSQL "CREATE TABLE TaskTypes (" & _
           "    TaskTypeID          AUTOINCREMENT CONSTRAINT PK_TaskTypes PRIMARY KEY," & _
           "    TaskCode            TEXT(20)  NOT NULL," & _
           "    TaskName            TEXT(100) NOT NULL," & _
           "    IsCertifiedMailTask YESNO     NOT NULL," & _
           "    IsActive            YESNO     NOT NULL," & _
           "    CONSTRAINT UQ_TaskTypes_TaskCode UNIQUE (TaskCode)" & _
           ");", "CREATE TABLE TaskTypes"

    RunSQL "CREATE TABLE CertifiedMailTypes (" & _
           "    CertifiedMailTypeID AUTOINCREMENT CONSTRAINT PK_CertifiedMailTypes PRIMARY KEY," & _
           "    MailTypeCode        TEXT(20)  NOT NULL," & _
           "    MailTypeName        TEXT(100) NOT NULL," & _
           "    IsActive            YESNO     NOT NULL," & _
           "    CONSTRAINT UQ_CertifiedMailTypes_Code UNIQUE (MailTypeCode)" & _
           ");", "CREATE TABLE CertifiedMailTypes"

    RunSQL "CREATE TABLE CertifiedMailStatuses (" & _
           "    CertifiedMailStatusID AUTOINCREMENT CONSTRAINT PK_CertifiedMailStatuses PRIMARY KEY," & _
           "    StatusCode            TEXT(20)  NOT NULL," & _
           "    StatusName            TEXT(100) NOT NULL," & _
           "    IsFinalStatus         YESNO     NOT NULL," & _
           "    IsActive              YESNO     NOT NULL," & _
           "    CONSTRAINT UQ_CertifiedMailStatuses_Code UNIQUE (StatusCode)" & _
           ");", "CREATE TABLE CertifiedMailStatuses"

    RunSQL "CREATE TABLE ReviewStatuses (" & _
           "    ReviewStatusID AUTOINCREMENT CONSTRAINT PK_ReviewStatuses PRIMARY KEY," & _
           "    StatusCode     TEXT(20)  NOT NULL," & _
           "    StatusName     TEXT(100) NOT NULL," & _
           "    IsActive       YESNO     NOT NULL," & _
           "    CONSTRAINT UQ_ReviewStatuses_Code UNIQUE (StatusCode)" & _
           ");", "CREATE TABLE ReviewStatuses"

    RunSQL "CREATE TABLE BacklogStatuses (" & _
           "    BacklogStatusID AUTOINCREMENT CONSTRAINT PK_BacklogStatuses PRIMARY KEY," & _
           "    StatusCode      TEXT(20)  NOT NULL," & _
           "    StatusName      TEXT(100) NOT NULL," & _
           "    IsClosedStatus  YESNO     NOT NULL," & _
           "    IsActive        YESNO     NOT NULL," & _
           "    CONSTRAINT UQ_BacklogStatuses_Code UNIQUE (StatusCode)" & _
           ");", "CREATE TABLE BacklogStatuses"

End Sub


' -----------------------------------------------------------------------------
' Step 2 -- Employees
' -----------------------------------------------------------------------------
Private Sub Create_Employees()

    RunSQL "CREATE TABLE Employees (" & _
           "    EmployeeID           AUTOINCREMENT CONSTRAINT PK_Employees PRIMARY KEY," & _
           "    EmployeeNumber       TEXT(20)  NOT NULL," & _
           "    FirstName            TEXT(50)  NOT NULL," & _
           "    LastName             TEXT(50)  NOT NULL," & _
           "    EmailAddress         TEXT(120)," & _
           "    UnitID               LONG      NOT NULL," & _
           "    SupervisorEmployeeID LONG," & _
           "    IsActive             YESNO     NOT NULL," & _
           "    HireDate             DATETIME," & _
           "    CONSTRAINT UQ_Employees_EmployeeNumber UNIQUE (EmployeeNumber)," & _
           "    CONSTRAINT FK_Employees_Units FOREIGN KEY (UnitID) REFERENCES Units (UnitID)" & _
           ");", "CREATE TABLE Employees"

End Sub


' -----------------------------------------------------------------------------
' Step 3 -- Reporting Months
' -----------------------------------------------------------------------------
Private Sub Create_ReportingMonths()

    RunSQL "CREATE TABLE ReportingMonths (" & _
           "    ReportingMonthID   AUTOINCREMENT CONSTRAINT PK_ReportingMonths PRIMARY KEY," & _
           "    MonthStart         DATETIME NOT NULL," & _
           "    MonthEnd           DATETIME NOT NULL," & _
           "    IsClosed           YESNO    NOT NULL," & _
           "    ClosedByEmployeeID LONG," & _
           "    ClosedOn           DATETIME," & _
           "    CONSTRAINT UQ_ReportingMonths_MonthStart UNIQUE (MonthStart)," & _
           "    CONSTRAINT FK_ReportingMonths_ClosedBy FOREIGN KEY (ClosedByEmployeeID)" & _
           "        REFERENCES Employees (EmployeeID)" & _
           ");", "CREATE TABLE ReportingMonths"

End Sub


' -----------------------------------------------------------------------------
' Step 4 -- Core Transactional Tables
' -----------------------------------------------------------------------------
Private Sub Create_CoreTables()

    RunSQL "CREATE TABLE ProductivityLogs (" & _
           "    ProductivityLogID AUTOINCREMENT CONSTRAINT PK_ProductivityLogs PRIMARY KEY," & _
           "    EmployeeID        LONG     NOT NULL," & _
           "    WorkDate          DATETIME NOT NULL," & _
           "    ReportingMonthID  LONG     NOT NULL," & _
           "    SubmittedOn       DATETIME," & _
           "    DailyComment      LONGTEXT," & _
           "    CONSTRAINT UQ_ProductivityLogs_EmployeeDate UNIQUE (EmployeeID, WorkDate)," & _
           "    CONSTRAINT FK_ProductivityLogs_Employee" & _
           "        FOREIGN KEY (EmployeeID) REFERENCES Employees (EmployeeID)," & _
           "    CONSTRAINT FK_ProductivityLogs_ReportingMonth" & _
           "        FOREIGN KEY (ReportingMonthID) REFERENCES ReportingMonths (ReportingMonthID)" & _
           ");", "CREATE TABLE ProductivityLogs"

    RunSQL "CREATE TABLE BacklogItems (" & _
           "    BacklogItemID          AUTOINCREMENT CONSTRAINT PK_BacklogItems PRIMARY KEY," & _
           "    UnitID                 LONG     NOT NULL," & _
           "    TaskTypeID             LONG     NOT NULL," & _
           "    AssignedEmployeeID     LONG," & _
           "    BacklogStatusID        LONG     NOT NULL," & _
           "    OpenedDate             DATETIME NOT NULL," & _
           "    TargetDate             DATETIME," & _
           "    ClosedDate             DATETIME," & _
           "    InitialQuantity        LONG     NOT NULL," & _
           "    RemainingQuantity      LONG     NOT NULL," & _
           "    CreatedFromLogDetailID LONG," & _
           "    BacklogNote            LONGTEXT," & _
           "    CONSTRAINT FK_BacklogItems_Unit" & _
           "        FOREIGN KEY (UnitID) REFERENCES Units (UnitID)," & _
           "    CONSTRAINT FK_BacklogItems_TaskType" & _
           "        FOREIGN KEY (TaskTypeID) REFERENCES TaskTypes (TaskTypeID)," & _
           "    CONSTRAINT FK_BacklogItems_AssignedEmployee" & _
           "        FOREIGN KEY (AssignedEmployeeID) REFERENCES Employees (EmployeeID)," & _
           "    CONSTRAINT FK_BacklogItems_Status" & _
           "        FOREIGN KEY (BacklogStatusID) REFERENCES BacklogStatuses (BacklogStatusID)" & _
           ");", "CREATE TABLE BacklogItems"

    RunSQL "CREATE TABLE ProductivityLogDetails (" & _
           "    ProductivityLogDetailID AUTOINCREMENT CONSTRAINT PK_ProductivityLogDetails PRIMARY KEY," & _
           "    ProductivityLogID       LONG NOT NULL," & _
           "    TaskTypeID              LONG NOT NULL," & _
           "    QuantityCompleted       LONG NOT NULL," & _
           "    CertifiedMailTypeID     LONG," & _
           "    CertifiedMailStatusID   LONG," & _
           "    BacklogItemID           LONG," & _
           "    DetailComment           LONGTEXT," & _
           "    CONSTRAINT FK_LogDetails_Log" & _
           "        FOREIGN KEY (ProductivityLogID) REFERENCES ProductivityLogs (ProductivityLogID)," & _
           "    CONSTRAINT FK_LogDetails_TaskType" & _
           "        FOREIGN KEY (TaskTypeID) REFERENCES TaskTypes (TaskTypeID)," & _
           "    CONSTRAINT FK_LogDetails_CertType" & _
           "        FOREIGN KEY (CertifiedMailTypeID) REFERENCES CertifiedMailTypes (CertifiedMailTypeID)," & _
           "    CONSTRAINT FK_LogDetails_CertStatus" & _
           "        FOREIGN KEY (CertifiedMailStatusID) REFERENCES CertifiedMailStatuses (CertifiedMailStatusID)," & _
           "    CONSTRAINT FK_LogDetails_BacklogItem" & _
           "        FOREIGN KEY (BacklogItemID) REFERENCES BacklogItems (BacklogItemID)" & _
           ");", "CREATE TABLE ProductivityLogDetails"

    RunSQL "CREATE TABLE SupervisorReviews (" & _
           "    SupervisorReviewID AUTOINCREMENT CONSTRAINT PK_SupervisorReviews PRIMARY KEY," & _
           "    ProductivityLogID  LONG     NOT NULL," & _
           "    ReviewerEmployeeID LONG     NOT NULL," & _
           "    ReviewStatusID     LONG     NOT NULL," & _
           "    ReviewDate         DATETIME NOT NULL," & _
           "    BacklogFlag        YESNO    NOT NULL," & _
           "    BacklogComment     LONGTEXT," & _
           "    CONSTRAINT UQ_SupervisorReviews_Log UNIQUE (ProductivityLogID)," & _
           "    CONSTRAINT FK_SupervisorReviews_Log" & _
           "        FOREIGN KEY (ProductivityLogID) REFERENCES ProductivityLogs (ProductivityLogID)," & _
           "    CONSTRAINT FK_SupervisorReviews_Reviewer" & _
           "        FOREIGN KEY (ReviewerEmployeeID) REFERENCES Employees (EmployeeID)," & _
           "    CONSTRAINT FK_SupervisorReviews_Status" & _
           "        FOREIGN KEY (ReviewStatusID) REFERENCES ReviewStatuses (ReviewStatusID)" & _
           ");", "CREATE TABLE SupervisorReviews"

End Sub


' -----------------------------------------------------------------------------
' Step 5 -- Deferred / Circular Constraints
' -----------------------------------------------------------------------------
Private Sub Add_DeferredConstraints()

    RunSQL "ALTER TABLE Employees " & _
           "ADD CONSTRAINT FK_Employees_Supervisor " & _
           "FOREIGN KEY (SupervisorEmployeeID) REFERENCES Employees (EmployeeID);", _
           "ALTER TABLE Employees ADD supervisor self-ref FK"

    RunSQL "ALTER TABLE BacklogItems " & _
           "ADD CONSTRAINT FK_BacklogItems_CreatedFromDetail " & _
           "FOREIGN KEY (CreatedFromLogDetailID) " & _
           "REFERENCES ProductivityLogDetails (ProductivityLogDetailID);", _
           "ALTER TABLE BacklogItems ADD circular FK"

End Sub


' -----------------------------------------------------------------------------
' Step 6 -- Indexes
' -----------------------------------------------------------------------------
Private Sub Create_Indexes()

    RunSQL "CREATE INDEX IX_Employees_UnitID ON Employees (UnitID);",                              "Index Employees.UnitID"
    RunSQL "CREATE INDEX IX_Employees_SupervisorID ON Employees (SupervisorEmployeeID);",          "Index Employees.SupervisorEmployeeID"
    RunSQL "CREATE INDEX IX_ProductivityLogs_EmployeeID ON ProductivityLogs (EmployeeID);",        "Index ProductivityLogs.EmployeeID"
    RunSQL "CREATE INDEX IX_ProductivityLogs_WorkDate ON ProductivityLogs (WorkDate);",            "Index ProductivityLogs.WorkDate"
    RunSQL "CREATE INDEX IX_ProductivityLogs_ReportingMonthID ON ProductivityLogs (ReportingMonthID);", "Index ProductivityLogs.ReportingMonthID"
    RunSQL "CREATE INDEX IX_LogDetails_ProductivityLogID ON ProductivityLogDetails (ProductivityLogID);",   "Index LogDetails.ProductivityLogID"
    RunSQL "CREATE INDEX IX_LogDetails_TaskTypeID ON ProductivityLogDetails (TaskTypeID);",         "Index LogDetails.TaskTypeID"
    RunSQL "CREATE INDEX IX_LogDetails_CertifiedMailTypeID ON ProductivityLogDetails (CertifiedMailTypeID);", "Index LogDetails.CertifiedMailTypeID"
    RunSQL "CREATE INDEX IX_LogDetails_CertifiedMailStatusID ON ProductivityLogDetails (CertifiedMailStatusID);", "Index LogDetails.CertifiedMailStatusID"
    RunSQL "CREATE INDEX IX_LogDetails_BacklogItemID ON ProductivityLogDetails (BacklogItemID);",   "Index LogDetails.BacklogItemID"
    RunSQL "CREATE INDEX IX_BacklogItems_UnitID ON BacklogItems (UnitID);",                        "Index BacklogItems.UnitID"
    RunSQL "CREATE INDEX IX_BacklogItems_TaskTypeID ON BacklogItems (TaskTypeID);",                "Index BacklogItems.TaskTypeID"
    RunSQL "CREATE INDEX IX_BacklogItems_AssignedEmployeeID ON BacklogItems (AssignedEmployeeID);", "Index BacklogItems.AssignedEmployeeID"
    RunSQL "CREATE INDEX IX_BacklogItems_BacklogStatusID ON BacklogItems (BacklogStatusID);",       "Index BacklogItems.BacklogStatusID"
    RunSQL "CREATE INDEX IX_BacklogItems_OpenedDate ON BacklogItems (OpenedDate);",                "Index BacklogItems.OpenedDate"
    RunSQL "CREATE INDEX IX_SupervisorReviews_ReviewerEmployeeID ON SupervisorReviews (ReviewerEmployeeID);", "Index SupervisorReviews.ReviewerEmployeeID"
    RunSQL "CREATE INDEX IX_SupervisorReviews_ReviewStatusID ON SupervisorReviews (ReviewStatusID);", "Index SupervisorReviews.ReviewStatusID"

End Sub


' -----------------------------------------------------------------------------
' Step 7 -- Seed Data (Lookup Tables)
' -----------------------------------------------------------------------------
Private Sub Insert_SeedData()

    ' Units
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('CS',  'Child Support',    True);", "Seed Units: CS"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('ENF', 'Enforcement',      True);", "Seed Units: ENF"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('EST', 'Establishment',    True);", "Seed Units: EST"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('FIN', 'Financial',        True);", "Seed Units: FIN"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('INT', 'Intake',           True);", "Seed Units: INT"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('LGL', 'Legal',            True);", "Seed Units: LGL"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('MED', 'Medical Support',  True);", "Seed Units: MED"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('PAT', 'Paternity',        True);", "Seed Units: PAT"
    RunSQL "INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('ADM', 'Administration',   True);", "Seed Units: ADM"

    ' TaskTypes
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CASE-REV',  'Case Review',                      False, True);", "Seed TaskTypes: CASE-REV"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('DATA-ENT',  'Data Entry',                       False, True);", "Seed TaskTypes: DATA-ENT"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('PHONE-CTR', 'Phone Contact',                    False, True);", "Seed TaskTypes: PHONE-CTR"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CORR-GEN',  'General Correspondence',           False, True);", "Seed TaskTypes: CORR-GEN"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CERT-MAIL', 'Certified Mail Processing',        True,  True);", "Seed TaskTypes: CERT-MAIL"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CERT-RTN',  'Certified Mail Return Processing', True,  True);", "Seed TaskTypes: CERT-RTN"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('LOCATE',    'Locate Activity',                  False, True);", "Seed TaskTypes: LOCATE"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('LEGAL-REV', 'Legal Document Review',            False, True);", "Seed TaskTypes: LEGAL-REV"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('HEARING',   'Hearing Preparation',              False, True);", "Seed TaskTypes: HEARING"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('COLLECTIONS','Collections Processing',         False, True);", "Seed TaskTypes: COLLECTIONS"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('PATERNITY', 'Paternity Establishment',          False, True);", "Seed TaskTypes: PATERNITY"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('MED-ENROLL','Medical Enrollment',              False, True);", "Seed TaskTypes: MED-ENROLL"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('INTERSTAT', 'Interstate Case Work',            False, True);", "Seed TaskTypes: INTERSTAT"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('TRAINING',  'Training / Meeting',              False, True);", "Seed TaskTypes: TRAINING"
    RunSQL "INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('ADMIN',     'Administrative Task',             False, True);", "Seed TaskTypes: ADMIN"

    ' CertifiedMailTypes
    RunSQL "INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('STD',   'Standard Certified Mail',                True);", "Seed CertMailType: STD"
    RunSQL "INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('RR',    'Return Receipt Requested',               True);", "Seed CertMailType: RR"
    RunSQL "INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('ERR',   'Electronic Return Receipt',              True);", "Seed CertMailType: ERR"
    RunSQL "INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('RD',    'Restricted Delivery',                    True);", "Seed CertMailType: RD"
    RunSQL "INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('RD-RR', 'Restricted Delivery with Return Receipt', True);", "Seed CertMailType: RD-RR"
    RunSQL "INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('SIG',   'Signature Confirmation',                 True);", "Seed CertMailType: SIG"

    ' CertifiedMailStatuses
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('SENT',      'Sent',                      False, True);", "Seed CertMailStatus: SENT"
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('IN-TRANS',  'In Transit',                False, True);", "Seed CertMailStatus: IN-TRANS"
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('DELIV',     'Delivered',                 True,  True);", "Seed CertMailStatus: DELIV"
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('RETURNED',  'Returned to Sender',        True,  True);", "Seed CertMailStatus: RETURNED"
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('REFUSED',   'Refused by Recipient',      True,  True);", "Seed CertMailStatus: REFUSED"
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('UNCLAIMED', 'Unclaimed / Undeliverable', True,  True);", "Seed CertMailStatus: UNCLAIMED"
    RunSQL "INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('PENDING',   'Pending Tracking Update',   False, True);", "Seed CertMailStatus: PENDING"

    ' ReviewStatuses
    RunSQL "INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('PENDING',   'Pending Review',          True);", "Seed ReviewStatus: PENDING"
    RunSQL "INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('APPROVED',  'Approved',                True);", "Seed ReviewStatus: APPROVED"
    RunSQL "INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('RETURNED',  'Returned for Correction', True);", "Seed ReviewStatus: RETURNED"
    RunSQL "INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('ESCALATED', 'Escalated',               True);", "Seed ReviewStatus: ESCALATED"

    ' BacklogStatuses
    RunSQL "INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('OPEN',      'Open',               False, True);", "Seed BacklogStatus: OPEN"
    RunSQL "INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('IN-PROG',   'In Progress',        False, True);", "Seed BacklogStatus: IN-PROG"
    RunSQL "INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('BLOCKED',   'Blocked',            False, True);", "Seed BacklogStatus: BLOCKED"
    RunSQL "INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('CLOSED',    'Closed / Completed', True,  True);", "Seed BacklogStatus: CLOSED"
    RunSQL "INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('CANCELLED', 'Cancelled',          True,  True);", "Seed BacklogStatus: CANCELLED"

    ' ReportingMonths (FY 2026)
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-01-01#, #2026-01-31#, True);",  "Seed Month: Jan 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-02-01#, #2026-02-28#, True);",  "Seed Month: Feb 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-03-01#, #2026-03-31#, True);",  "Seed Month: Mar 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-04-01#, #2026-04-30#, True);",  "Seed Month: Apr 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-05-01#, #2026-05-31#, True);",  "Seed Month: May 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-06-01#, #2026-06-30#, False);", "Seed Month: Jun 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-07-01#, #2026-07-31#, False);", "Seed Month: Jul 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-08-01#, #2026-08-31#, False);", "Seed Month: Aug 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-09-01#, #2026-09-30#, False);", "Seed Month: Sep 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-10-01#, #2026-10-31#, False);", "Seed Month: Oct 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-11-01#, #2026-11-30#, False);", "Seed Month: Nov 2026"
    RunSQL "INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-12-01#, #2026-12-31#, False);", "Seed Month: Dec 2026"

End Sub

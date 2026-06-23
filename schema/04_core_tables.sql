-- =============================================================================
-- OCSS Employee Productivity Log
-- Schema Script 04 -- Core Transactional Tables
-- Depends on: 01_lookup_tables.sql, 02_employees.sql, 03_reporting_months.sql
--
-- Creation order matters because of foreign key references:
--   1. ProductivityLogs
--   2. BacklogItems          (no FK to ProductivityLogDetails yet)
--   3. ProductivityLogDetails
--   4. SupervisorReviews
--
-- The circular FK BacklogItems -> ProductivityLogDetails is added
-- separately in 05_constraints.sql after both tables exist.
-- =============================================================================


-- -----------------------------------------------------------------------------
-- ProductivityLogs  (daily header -- one row per employee per work date)
-- The UNIQUE constraint on (EmployeeID, WorkDate) prevents duplicate entries.
-- -----------------------------------------------------------------------------
CREATE TABLE ProductivityLogs (
    ProductivityLogID  AUTOINCREMENT CONSTRAINT PK_ProductivityLogs PRIMARY KEY,
    EmployeeID         LONG      NOT NULL,
    WorkDate           DATETIME  NOT NULL,
    ReportingMonthID   LONG      NOT NULL,
    SubmittedOn        DATETIME,
    DailyComment       LONGTEXT,
    CONSTRAINT UQ_ProductivityLogs_EmployeeDate UNIQUE (EmployeeID, WorkDate),
    CONSTRAINT FK_ProductivityLogs_Employee
        FOREIGN KEY (EmployeeID) REFERENCES Employees (EmployeeID),
    CONSTRAINT FK_ProductivityLogs_ReportingMonth
        FOREIGN KEY (ReportingMonthID) REFERENCES ReportingMonths (ReportingMonthID)
);


-- -----------------------------------------------------------------------------
-- BacklogItems  (work items that could not be completed on the log date)
-- CreatedFromLogDetailID is populated when a backlog item is raised directly
-- from a detail line; the FK for that column is added in 05_constraints.sql.
-- -----------------------------------------------------------------------------
CREATE TABLE BacklogItems (
    BacklogItemID        AUTOINCREMENT CONSTRAINT PK_BacklogItems PRIMARY KEY,
    UnitID               LONG      NOT NULL,
    TaskTypeID           LONG      NOT NULL,
    AssignedEmployeeID   LONG,
    BacklogStatusID      LONG      NOT NULL,
    OpenedDate           DATETIME  NOT NULL,
    TargetDate           DATETIME,
    ClosedDate           DATETIME,
    InitialQuantity      LONG      NOT NULL,
    RemainingQuantity    LONG      NOT NULL,
    CreatedFromLogDetailID LONG,
    BacklogNote          LONGTEXT,
    CONSTRAINT FK_BacklogItems_Unit
        FOREIGN KEY (UnitID) REFERENCES Units (UnitID),
    CONSTRAINT FK_BacklogItems_TaskType
        FOREIGN KEY (TaskTypeID) REFERENCES TaskTypes (TaskTypeID),
    CONSTRAINT FK_BacklogItems_AssignedEmployee
        FOREIGN KEY (AssignedEmployeeID) REFERENCES Employees (EmployeeID),
    CONSTRAINT FK_BacklogItems_Status
        FOREIGN KEY (BacklogStatusID) REFERENCES BacklogStatuses (BacklogStatusID)
);


-- -----------------------------------------------------------------------------
-- ProductivityLogDetails  (one row per task line on a daily log)
-- CertifiedMailTypeID and CertifiedMailStatusID are required when the linked
-- TaskType has IsCertifiedMailTask = True (enforced in the entry form).
-- BacklogItemID optionally links this detail line to an open backlog item.
-- -----------------------------------------------------------------------------
CREATE TABLE ProductivityLogDetails (
    ProductivityLogDetailID AUTOINCREMENT CONSTRAINT PK_ProductivityLogDetails PRIMARY KEY,
    ProductivityLogID       LONG NOT NULL,
    TaskTypeID              LONG NOT NULL,
    QuantityCompleted       LONG NOT NULL,
    CertifiedMailTypeID     LONG,
    CertifiedMailStatusID   LONG,
    BacklogItemID           LONG,
    DetailComment           LONGTEXT,
    CONSTRAINT FK_LogDetails_Log
        FOREIGN KEY (ProductivityLogID) REFERENCES ProductivityLogs (ProductivityLogID),
    CONSTRAINT FK_LogDetails_TaskType
        FOREIGN KEY (TaskTypeID) REFERENCES TaskTypes (TaskTypeID),
    CONSTRAINT FK_LogDetails_CertType
        FOREIGN KEY (CertifiedMailTypeID) REFERENCES CertifiedMailTypes (CertifiedMailTypeID),
    CONSTRAINT FK_LogDetails_CertStatus
        FOREIGN KEY (CertifiedMailStatusID) REFERENCES CertifiedMailStatuses (CertifiedMailStatusID),
    CONSTRAINT FK_LogDetails_BacklogItem
        FOREIGN KEY (BacklogItemID) REFERENCES BacklogItems (BacklogItemID)
);


-- -----------------------------------------------------------------------------
-- SupervisorReviews  (one review record per daily log -- enforced by UNIQUE)
-- BacklogFlag = True signals that the supervisor flagged items for backlog.
-- -----------------------------------------------------------------------------
CREATE TABLE SupervisorReviews (
    SupervisorReviewID  AUTOINCREMENT CONSTRAINT PK_SupervisorReviews PRIMARY KEY,
    ProductivityLogID   LONG     NOT NULL,
    ReviewerEmployeeID  LONG     NOT NULL,
    ReviewStatusID      LONG     NOT NULL,
    ReviewDate          DATETIME NOT NULL,
    BacklogFlag         YESNO    NOT NULL,
    BacklogComment      LONGTEXT,
    CONSTRAINT UQ_SupervisorReviews_Log UNIQUE (ProductivityLogID),
    CONSTRAINT FK_SupervisorReviews_Log
        FOREIGN KEY (ProductivityLogID) REFERENCES ProductivityLogs (ProductivityLogID),
    CONSTRAINT FK_SupervisorReviews_Reviewer
        FOREIGN KEY (ReviewerEmployeeID) REFERENCES Employees (EmployeeID),
    CONSTRAINT FK_SupervisorReviews_Status
        FOREIGN KEY (ReviewStatusID) REFERENCES ReviewStatuses (ReviewStatusID)
);

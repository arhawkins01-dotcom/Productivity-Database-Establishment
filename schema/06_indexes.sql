-- =============================================================================
-- OCSS Employee Productivity Log
-- Schema Script 06 -- Indexes
-- Depends on: ALL previous schema scripts (01-05) must have run first.
--
-- Primary keys are already indexed automatically by Access.
-- These indexes cover the most common query join and filter columns.
-- =============================================================================


-- Employees
CREATE INDEX IX_Employees_UnitID
    ON Employees (UnitID);

CREATE INDEX IX_Employees_SupervisorID
    ON Employees (SupervisorEmployeeID);


-- ProductivityLogs
CREATE INDEX IX_ProductivityLogs_EmployeeID
    ON ProductivityLogs (EmployeeID);

CREATE INDEX IX_ProductivityLogs_WorkDate
    ON ProductivityLogs (WorkDate);

CREATE INDEX IX_ProductivityLogs_ReportingMonthID
    ON ProductivityLogs (ReportingMonthID);


-- ProductivityLogDetails
CREATE INDEX IX_LogDetails_ProductivityLogID
    ON ProductivityLogDetails (ProductivityLogID);

CREATE INDEX IX_LogDetails_TaskTypeID
    ON ProductivityLogDetails (TaskTypeID);

CREATE INDEX IX_LogDetails_CertifiedMailTypeID
    ON ProductivityLogDetails (CertifiedMailTypeID);

CREATE INDEX IX_LogDetails_CertifiedMailStatusID
    ON ProductivityLogDetails (CertifiedMailStatusID);

CREATE INDEX IX_LogDetails_BacklogItemID
    ON ProductivityLogDetails (BacklogItemID);


-- BacklogItems
CREATE INDEX IX_BacklogItems_UnitID
    ON BacklogItems (UnitID);

CREATE INDEX IX_BacklogItems_TaskTypeID
    ON BacklogItems (TaskTypeID);

CREATE INDEX IX_BacklogItems_AssignedEmployeeID
    ON BacklogItems (AssignedEmployeeID);

CREATE INDEX IX_BacklogItems_BacklogStatusID
    ON BacklogItems (BacklogStatusID);

CREATE INDEX IX_BacklogItems_OpenedDate
    ON BacklogItems (OpenedDate);


-- SupervisorReviews
CREATE INDEX IX_SupervisorReviews_ReviewerEmployeeID
    ON SupervisorReviews (ReviewerEmployeeID);

CREATE INDEX IX_SupervisorReviews_ReviewStatusID
    ON SupervisorReviews (ReviewStatusID);

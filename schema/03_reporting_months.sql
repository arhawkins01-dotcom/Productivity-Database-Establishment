-- =============================================================================
-- OCSS Employee Productivity Log
-- Schema Script 03 -- ReportingMonths Table
-- Depends on: 02_employees.sql (Employees must exist for ClosedByEmployeeID)
-- =============================================================================


-- -----------------------------------------------------------------------------
-- ReportingMonths
-- One row per calendar month reporting period.
-- MonthStart is always the first day of the month (e.g. 2026-06-01).
-- MonthEnd   is always the last day of the month  (e.g. 2026-06-30).
-- IsClosed   prevents further log entries once the supervisor signs off.
-- -----------------------------------------------------------------------------
CREATE TABLE ReportingMonths (
    ReportingMonthID    AUTOINCREMENT CONSTRAINT PK_ReportingMonths PRIMARY KEY,
    MonthStart          DATETIME NOT NULL,
    MonthEnd            DATETIME NOT NULL,
    IsClosed            YESNO    NOT NULL,
    ClosedByEmployeeID  LONG,
    ClosedOn            DATETIME,
    CONSTRAINT UQ_ReportingMonths_MonthStart UNIQUE (MonthStart),
    CONSTRAINT FK_ReportingMonths_ClosedBy
        FOREIGN KEY (ClosedByEmployeeID) REFERENCES Employees (EmployeeID)
);

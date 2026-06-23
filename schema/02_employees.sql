-- =============================================================================
-- OCSS Employee Productivity Log
-- Schema Script 02 -- Employees Table
-- Depends on: 01_lookup_tables.sql (Units must exist first)
--
-- NOTE: The self-referencing supervisor FK is added via ALTER TABLE in
--       05_constraints.sql after this table is created, because Access
--       cannot resolve a forward self-reference in the same CREATE TABLE.
-- =============================================================================


-- -----------------------------------------------------------------------------
-- Employees
-- One row per OCSS staff member or supervisor.
-- SupervisorEmployeeID links back to this same table (added in 05_constraints).
-- -----------------------------------------------------------------------------
CREATE TABLE Employees (
    EmployeeID         AUTOINCREMENT CONSTRAINT PK_Employees PRIMARY KEY,
    EmployeeNumber     TEXT(20)  NOT NULL,
    FirstName          TEXT(50)  NOT NULL,
    LastName           TEXT(50)  NOT NULL,
    EmailAddress       TEXT(120),
    UnitID             LONG      NOT NULL,
    SupervisorEmployeeID LONG,
    IsActive           YESNO     NOT NULL,
    HireDate           DATETIME,
    CONSTRAINT UQ_Employees_EmployeeNumber UNIQUE (EmployeeNumber),
    CONSTRAINT FK_Employees_Units
        FOREIGN KEY (UnitID) REFERENCES Units (UnitID)
);

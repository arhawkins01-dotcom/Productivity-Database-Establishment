-- =============================================================================
-- OCSS Employee Productivity Log
-- Schema Script 05 -- Deferred / Circular Constraints
-- Depends on: ALL previous schema scripts (01-04) must have run first.
--
-- Two constraints cannot be declared inline because the referenced table
-- did not yet exist at CREATE TABLE time:
--   1. Employees.SupervisorEmployeeID -> Employees (self-reference)
--   2. BacklogItems.CreatedFromLogDetailID -> ProductivityLogDetails (circular)
-- =============================================================================


-- Employees: supervisor self-reference
ALTER TABLE Employees
    ADD CONSTRAINT FK_Employees_Supervisor
    FOREIGN KEY (SupervisorEmployeeID) REFERENCES Employees (EmployeeID);


-- BacklogItems: optional link back to the log detail that spawned the item
ALTER TABLE BacklogItems
    ADD CONSTRAINT FK_BacklogItems_CreatedFromDetail
    FOREIGN KEY (CreatedFromLogDetailID)
    REFERENCES ProductivityLogDetails (ProductivityLogDetailID);

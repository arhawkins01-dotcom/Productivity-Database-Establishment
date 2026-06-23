-- =============================================================================
-- OCSS Employee Productivity Log
-- Report Query: Monthly Productivity Totals by Employee and Unit
-- Save this as a named query (e.g. qry_Rpt_MonthlyByEmployee) in Access.
--
-- Parameters (set via form or replace with literal values):
--   [Enter Reporting Month Start Date (e.g. 6/1/2026)]: prompt for month
-- =============================================================================

SELECT
    rm.MonthStart,
    u.UnitName,
    e.LastName & ", " & e.FirstName    AS EmployeeName,
    e.EmployeeNumber,
    COUNT(DISTINCT pl.ProductivityLogID) AS DaysLogged,
    SUM(pld.QuantityCompleted)           AS TotalUnitsCompleted,
    SUM(IIf(tt.IsCertifiedMailTask = True, pld.QuantityCompleted, 0)) AS CertifiedMailTotal
FROM
    (((((ProductivityLogDetails AS pld
        INNER JOIN ProductivityLogs     AS pl  ON pld.ProductivityLogID = pl.ProductivityLogID)
        INNER JOIN ReportingMonths      AS rm  ON pl.ReportingMonthID   = rm.ReportingMonthID)
        INNER JOIN Employees            AS e   ON pl.EmployeeID         = e.EmployeeID)
        INNER JOIN Units                AS u   ON e.UnitID              = u.UnitID)
        INNER JOIN TaskTypes            AS tt  ON pld.TaskTypeID        = tt.TaskTypeID)
WHERE
    rm.MonthStart = [Enter Reporting Month Start Date (e.g. 6/1/2026)]
    AND e.IsActive = True
GROUP BY
    rm.MonthStart,
    u.UnitName,
    e.LastName,
    e.FirstName,
    e.EmployeeNumber
ORDER BY
    u.UnitName,
    e.LastName,
    e.FirstName;

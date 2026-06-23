-- =============================================================================
-- OCSS Employee Productivity Log
-- Report Query: Certified Mail Detail
-- Lists every certified mail task line with type and delivery status.
-- Save as qry_Rpt_CertifiedMailDetail in Access.
-- =============================================================================

SELECT
    rm.MonthStart,
    u.UnitName,
    e.LastName & ", " & e.FirstName    AS EmployeeName,
    pl.WorkDate,
    tt.TaskName,
    cmt.MailTypeName,
    cms.StatusName                     AS DeliveryStatus,
    cms.IsFinalStatus,
    pld.QuantityCompleted,
    pld.DetailComment
FROM
    ((((((ProductivityLogDetails AS pld
        INNER JOIN ProductivityLogs      AS pl   ON pld.ProductivityLogID     = pl.ProductivityLogID)
        INNER JOIN ReportingMonths       AS rm   ON pl.ReportingMonthID       = rm.ReportingMonthID)
        INNER JOIN Employees             AS e    ON pl.EmployeeID             = e.EmployeeID)
        INNER JOIN Units                 AS u    ON e.UnitID                  = u.UnitID)
        INNER JOIN TaskTypes             AS tt   ON pld.TaskTypeID            = tt.TaskTypeID)
        INNER JOIN CertifiedMailTypes    AS cmt  ON pld.CertifiedMailTypeID   = cmt.CertifiedMailTypeID)
        INNER JOIN CertifiedMailStatuses AS cms  ON pld.CertifiedMailStatusID = cms.CertifiedMailStatusID
WHERE
    tt.IsCertifiedMailTask = True
    AND rm.MonthStart = [Enter Reporting Month Start Date (e.g. 6/1/2026)]
ORDER BY
    u.UnitName,
    e.LastName,
    e.FirstName,
    pl.WorkDate DESC;

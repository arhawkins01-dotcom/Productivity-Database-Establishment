-- =============================================================================
-- OCSS Employee Productivity Log
-- Report Query: Monthly Productivity Totals by Task Type
-- Includes certified mail breakdown by mail type and delivery status.
-- Save as qry_Rpt_MonthlyByTaskType in Access.
-- =============================================================================

SELECT
    rm.MonthStart,
    u.UnitName,
    tt.TaskName,
    tt.IsCertifiedMailTask,
    IIf(tt.IsCertifiedMailTask = True,
        IIf(cmt.MailTypeName IS NULL, "(Not Specified)", cmt.MailTypeName),
        NULL)                                           AS MailTypeName,
    IIf(tt.IsCertifiedMailTask = True,
        IIf(cms.StatusName IS NULL, "(Not Specified)", cms.StatusName),
        NULL)                                           AS MailStatusName,
    SUM(pld.QuantityCompleted)                          AS TotalQuantity
FROM
    ((((((ProductivityLogDetails AS pld
        INNER JOIN ProductivityLogs      AS pl   ON pld.ProductivityLogID   = pl.ProductivityLogID)
        INNER JOIN ReportingMonths       AS rm   ON pl.ReportingMonthID     = rm.ReportingMonthID)
        INNER JOIN Employees             AS e    ON pl.EmployeeID           = e.EmployeeID)
        INNER JOIN Units                 AS u    ON e.UnitID                = u.UnitID)
        INNER JOIN TaskTypes             AS tt   ON pld.TaskTypeID          = tt.TaskTypeID)
        LEFT  JOIN CertifiedMailTypes    AS cmt  ON pld.CertifiedMailTypeID = cmt.CertifiedMailTypeID)
        LEFT  JOIN CertifiedMailStatuses AS cms  ON pld.CertifiedMailStatusID = cms.CertifiedMailStatusID
WHERE
    rm.MonthStart = [Enter Reporting Month Start Date (e.g. 6/1/2026)]
GROUP BY
    rm.MonthStart,
    u.UnitName,
    tt.TaskName,
    tt.IsCertifiedMailTask,
    cmt.MailTypeName,
    cms.StatusName
ORDER BY
    u.UnitName,
    tt.TaskName,
    cmt.MailTypeName,
    cms.StatusName;

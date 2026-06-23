-- =============================================================================
-- OCSS Employee Productivity Log
-- Report Query: Supervisor Review Status
-- Shows the review status for every daily log in a given month,
-- flagging any logs that are still pending review.
-- Save as qry_Rpt_SupervisorReviews in Access.
-- =============================================================================

SELECT
    rm.MonthStart,
    u.UnitName,
    e.LastName & ", " & e.FirstName        AS EmployeeName,
    pl.WorkDate,
    pl.SubmittedOn,
    IIf(sr.SupervisorReviewID IS NULL,
        "Not Yet Reviewed",
        rs.StatusName)                      AS ReviewStatus,
    sr.ReviewDate,
    sup.LastName & ", " & sup.FirstName     AS ReviewerName,
    sr.BacklogFlag,
    sr.BacklogComment
FROM
    ((((((ProductivityLogs AS pl
        INNER JOIN ReportingMonths   AS rm  ON pl.ReportingMonthID       = rm.ReportingMonthID)
        INNER JOIN Employees         AS e   ON pl.EmployeeID             = e.EmployeeID)
        INNER JOIN Units             AS u   ON e.UnitID                  = u.UnitID)
        LEFT  JOIN SupervisorReviews AS sr  ON pl.ProductivityLogID      = sr.ProductivityLogID)
        LEFT  JOIN ReviewStatuses    AS rs  ON sr.ReviewStatusID         = rs.ReviewStatusID)
        LEFT  JOIN Employees         AS sup ON sr.ReviewerEmployeeID     = sup.EmployeeID)
WHERE
    rm.MonthStart = [Enter Reporting Month Start Date (e.g. 6/1/2026)]
ORDER BY
    u.UnitName,
    e.LastName,
    e.FirstName,
    pl.WorkDate;

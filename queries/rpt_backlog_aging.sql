-- =============================================================================
-- OCSS Employee Productivity Log
-- Report Query: Open Backlog Aging
-- Shows all open backlog items ordered by age (oldest first).
-- Save as qry_Rpt_BacklogAging in Access.
-- =============================================================================

SELECT
    bi.BacklogItemID,
    u.UnitName,
    tt.TaskName,
    e.LastName & ", " & e.FirstName    AS AssignedEmployee,
    bs.StatusName                      AS BacklogStatus,
    bi.OpenedDate,
    bi.TargetDate,
    DateDiff("d", bi.OpenedDate, Now()) AS DaysOpen,
    bi.InitialQuantity,
    bi.RemainingQuantity,
    bi.InitialQuantity - bi.RemainingQuantity AS QuantityCompleted,
    IIf(bi.TargetDate IS NOT NULL AND Now() > bi.TargetDate, True, False) AS IsPastDue,
    -- Supervisor who flagged the backlog item (from the originating review)
    sup.LastName & ", " & sup.FirstName AS OriginatingSupervisor,
    sr.ReviewDate                       AS FlaggedOn,
    bi.BacklogNote
FROM
    (((((BacklogItems AS bi
        INNER JOIN Units            AS u   ON bi.UnitID             = u.UnitID)
        INNER JOIN TaskTypes        AS tt  ON bi.TaskTypeID         = tt.TaskTypeID)
        INNER JOIN BacklogStatuses  AS bs  ON bi.BacklogStatusID    = bs.BacklogStatusID)
        LEFT  JOIN Employees        AS e   ON bi.AssignedEmployeeID = e.EmployeeID)
        -- Link back to the originating log detail -> log -> review
        LEFT  JOIN ProductivityLogDetails AS orig_pld
                                            ON bi.CreatedFromLogDetailID = orig_pld.ProductivityLogDetailID)
        LEFT  JOIN SupervisorReviews AS sr ON orig_pld.ProductivityLogID = sr.ProductivityLogID)
        LEFT  JOIN Employees AS sup        ON sr.ReviewerEmployeeID      = sup.EmployeeID
WHERE
    bs.IsClosedStatus = False
ORDER BY
    bi.OpenedDate,
    u.UnitName;

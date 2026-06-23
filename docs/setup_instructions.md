# OCSS Employee Productivity Log — Database Setup Guide

## Requirements
- Microsoft Access 2016 or later (Microsoft 365 recommended)
- Shared network drive accessible to all 35+ users (for split-database deployment)

---

## 1. Build the Backend Database

### Option A — Automated (recommended)

1. Create a new blank Access database: **OCSS_Productivity_BE.accdb**
2. Open the VBA Editor: **Alt + F11**
3. **Insert > Module**
4. Paste the entire contents of [vba/Build_Schema.bas](../vba/Build_Schema.bas) into the module
5. Press **F5** (or Run > Run Sub/UserForm) with the cursor inside `RunAllScripts`
6. Watch the **Immediate Window** (Ctrl + G) for confirmation messages
7. Close the VBA Editor when build is complete

### Option B — Manual (step by step)

Open **OCSS_Productivity_BE.accdb** in Access. For each script below:

1. Go to **Create > Query Design**
2. Close the "Show Table" dialog
3. Switch to **SQL View** (Home tab)
4. Paste one SQL statement at a time and click **Run (!)** after each

Run scripts in this order:

| Order | File | Purpose |
|-------|------|---------|
| 1 | [schema/01_lookup_tables.sql](../schema/01_lookup_tables.sql) | Reference / code tables |
| 2 | [schema/02_employees.sql](../schema/02_employees.sql) | Employee roster |
| 3 | [schema/03_reporting_months.sql](../schema/03_reporting_months.sql) | Monthly periods |
| 4 | [schema/04_core_tables.sql](../schema/04_core_tables.sql) | Logs + details + backlog + reviews |
| 5 | [schema/05_constraints.sql](../schema/05_constraints.sql) | Deferred / circular FKs |
| 6 | [schema/06_indexes.sql](../schema/06_indexes.sql) | Performance indexes |
| 7 | [seed_data/01_lookup_seed_data.sql](../seed_data/01_lookup_seed_data.sql) | Default lookup values |

---

## 2. Customise Lookup Data

Before distributing the database, update the following tables to match your agency:

| Table | What to customise |
|-------|--------------------|
| **Units** | Add/rename organisational units |
| **TaskTypes** | Add task codes specific to your workflow; set `IsCertifiedMailTask = True` for any mail tasks |
| **CertifiedMailTypes** | Match your USPS certified mail service options |
| **CertifiedMailStatuses** | Match your tracking status workflow |
| **ReviewStatuses** | Match your supervisor approval process |
| **ReportingMonths** | Extend rows for future months as needed |

---

## 3. Add Employees

Populate the **Employees** table with your 35+ staff:

- Set `UnitID` to the correct unit for each employee
- Set `SupervisorEmployeeID` to the `EmployeeID` of their direct supervisor
  (supervisors must be inserted first, or use a second-pass UPDATE)
- Set `IsActive = True` for current staff

---

## 4. Split the Database (Multi-User Deployment)

Access supports up to ~35–50 simultaneous users when properly split.

1. With **OCSS_Productivity_BE.accdb** complete, go to:
   **Database Tools > Access Database (Split Database)**
2. Save the split frontend as **OCSS_Productivity_FE.accdb** locally on your machine
3. Copy **OCSS_Productivity_BE.accdb** to a shared network folder (e.g. `\\server\OCSS\DB\`)
4. Re-link the frontend tables to point to the network backend:
   **External Data > Linked Table Manager**
5. Distribute a copy of the **frontend (.accdb or compiled .accde)** to each user's local machine
6. Each user opens their **local frontend**; all data is written to the shared **backend**

### Multi-user settings to enable in the backend

Open the backend, go to **File > Options > Client Settings**:

- Default open mode: **Shared**
- Default record locking: **Edited Record** (row-level locking)
- Open databases using record-level locking: **checked**

---

## 5. Create Saved Queries for Reports

For each file in the [queries/](../queries/) folder, save it as a named query in Access:

1. **Create > Query Design > SQL View**
2. Paste the query SQL
3. **File > Save** and name the query (suggested names are in the file header comments)

| File | Suggested query name | Description |
|------|----------------------|-------------|
| [rpt_monthly_by_employee.sql](../queries/rpt_monthly_by_employee.sql) | `qry_Rpt_MonthlyByEmployee` | Monthly totals per employee and unit |
| [rpt_monthly_by_tasktype.sql](../queries/rpt_monthly_by_tasktype.sql) | `qry_Rpt_MonthlyByTaskType` | Monthly totals by task with certified mail breakdown |
| [rpt_certified_mail_detail.sql](../queries/rpt_certified_mail_detail.sql) | `qry_Rpt_CertifiedMailDetail` | Certified mail line detail with delivery status |
| [rpt_backlog_aging.sql](../queries/rpt_backlog_aging.sql) | `qry_Rpt_BacklogAging` | Open backlog items with days-open calculation |
| [rpt_supervisor_reviews.sql](../queries/rpt_supervisor_reviews.sql) | `qry_Rpt_SupervisorReviews` | Review status for every daily log in a month |

All monthly reports use an Access parameter prompt — when you open the query or report, Access will ask:

> **Enter Reporting Month Start Date (e.g. 6/1/2026)**

---

## 6. Certified Mail Entry Rule

In your daily log entry form, add the following logic (VBA or macro):

- When the user selects a `TaskTypeID` where `IsCertifiedMailTask = True`:
  - **Require** `CertifiedMailTypeID` (mail type combo box must not be empty)
  - **Require** `CertifiedMailStatusID` (delivery status must not be empty)
- Otherwise, lock/hide those two fields to prevent accidental data entry

---

## 7. Closing a Reporting Month

When a month's data is complete and supervisor-approved:

1. A supervisor (or admin) sets `IsClosed = True` on the `ReportingMonths` row for that month
2. Set `ClosedByEmployeeID` and `ClosedOn` at the same time
3. Your entry form should check `IsClosed` and prevent new `ProductivityLogs` rows being created
   for dates within a closed month

---

## 8. Ongoing Maintenance

- **New employees**: insert a row in `Employees`; no schema change needed
- **New task types**: insert a row in `TaskTypes`; set `IsCertifiedMailTask` appropriately
- **New months**: insert rows in `ReportingMonths` at least one month ahead
- **Compact & repair**: run monthly on the backend (`Database Tools > Compact and Repair`)
- **Backup**: back up `OCSS_Productivity_BE.accdb` nightly to a separate location

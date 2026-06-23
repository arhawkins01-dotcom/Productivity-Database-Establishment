# OCSS Employee Productivity Log — Database

Microsoft Access database for tracking daily employee productivity across 35+ users. Covers task completion, certified mail processing, supervisor review, backlog management, and monthly reporting.

---

## Repository Structure

```
schema/
  01_lookup_tables.sql      -- Units, TaskTypes, CertifiedMailTypes/Statuses, ReviewStatuses, BacklogStatuses
  02_employees.sql          -- Employees table (links to Units)
  03_reporting_months.sql   -- Monthly reporting periods
  04_core_tables.sql        -- ProductivityLogs, ProductivityLogDetails, BacklogItems, SupervisorReviews
  05_constraints.sql        -- Deferred/circular foreign keys (run after 04)
  06_indexes.sql            -- Performance indexes

seed_data/
  01_lookup_seed_data.sql   -- Default rows for all lookup tables + FY2026 reporting months

queries/
  rpt_monthly_by_employee.sql       -- Monthly totals per employee and unit
  rpt_monthly_by_tasktype.sql       -- Monthly totals by task type with certified mail breakdown
  rpt_certified_mail_detail.sql     -- Certified mail line detail with delivery status
  rpt_backlog_aging.sql             -- Open backlog items with days-open calculation
  rpt_supervisor_reviews.sql        -- Review status for every daily log in a month

vba/
  Build_Schema.bas          -- VBA module: runs all DDL + seed data in one click

docs/
  setup_instructions.md     -- Full step-by-step setup and deployment guide
```

---

## Quick Start

### 1. Build the Database (one-time)

1. Create a new blank `.accdb` file in Microsoft Access
2. Open the VBA Editor: **Alt + F11**
3. **File > Import File** → select `vba/Build_Schema.bas`
4. Press **F5** with the cursor inside `RunAllScripts`
5. Watch the Immediate Window (**Ctrl + G**) for confirmation

All 12 tables, foreign keys, indexes, and lookup seed data are created automatically.

### 2. Customise Lookups

Edit these tables in Access before distributing to users:

| Table | What to update |
|-------|----------------|
| `Units` | Your org's actual unit names and codes |
| `TaskTypes` | Your task list; set `IsCertifiedMailTask = True` for mail tasks |
| `CertifiedMailTypes` | USPS service types you use |
| `CertifiedMailStatuses` | Your delivery tracking workflow |
| `ReportingMonths` | Extend rows for future months as needed |

### 3. Add Employees

Populate `Employees` with your staff. Set `UnitID`, `SupervisorEmployeeID`, and `IsActive`.

### 4. Save Report Queries

In Access, save each file from `queries/` as a named query (SQL View → paste → Save).

### 5. Split for Multi-User Deployment

**Database Tools > Access Database** → move the backend to a shared network folder, distribute the frontend locally to each of the 35+ users.

See [docs/setup_instructions.md](docs/setup_instructions.md) for the full guide.

---

## Tables at a Glance

| Table | Purpose |
|-------|---------|
| `Units` | Organisational units |
| `Employees` | Staff roster with supervisor chain |
| `TaskTypes` | Task catalogue; flags certified mail tasks |
| `CertifiedMailTypes` | USPS mail service types |
| `CertifiedMailStatuses` | Delivery outcome tracking |
| `ReviewStatuses` | Supervisor review outcomes |
| `BacklogStatuses` | Backlog item lifecycle states |
| `ReportingMonths` | Monthly reporting periods with close/sign-off |
| `ProductivityLogs` | Daily log header — one row per employee per day |
| `ProductivityLogDetails` | Task lines on a daily log |
| `BacklogItems` | Work items that could not be completed on the log date |
| `SupervisorReviews` | Supervisor review record for each daily log |
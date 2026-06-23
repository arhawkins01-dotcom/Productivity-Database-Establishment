-- =============================================================================
-- OCSS Employee Productivity Log
-- Seed Data -- Lookup Tables
-- Run AFTER all schema scripts (01-06) have completed successfully.
--
-- Customise these rows to match your agency's actual codes and names
-- before distributing the database to users.
-- =============================================================================


-- -----------------------------------------------------------------------------
-- Units  (add/rename to match your org chart)
-- -----------------------------------------------------------------------------
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('CS',  'Child Support',               True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('ENF', 'Enforcement',                 True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('EST', 'Establishment',               True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('FIN', 'Financial',                   True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('INT', 'Intake',                      True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('LGL', 'Legal',                       True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('MED', 'Medical Support',             True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('PAT', 'Paternity',                   True);
INSERT INTO Units (UnitCode, UnitName, IsActive) VALUES ('ADM', 'Administration',              True);


-- -----------------------------------------------------------------------------
-- TaskTypes  (IsCertifiedMailTask = True for any mail processing tasks)
-- -----------------------------------------------------------------------------
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CASE-REV',  'Case Review',                         False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('DATA-ENT',  'Data Entry',                          False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('PHONE-CTR', 'Phone Contact',                       False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CORR-GEN',  'General Correspondence',              False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CERT-MAIL', 'Certified Mail Processing',           True,  True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('CERT-RTN',  'Certified Mail Return Processing',    True,  True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('LOCATE',    'Locate Activity',                     False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('LEGAL-REV', 'Legal Document Review',               False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('HEARING',   'Hearing Preparation',                 False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('COLLECTIONS','Collections Processing',            False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('PATERNITY', 'Paternity Establishment',             False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('MED-ENROLL','Medical Enrollment',                 False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('INTERSTAT', 'Interstate Case Work',               False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('TRAINING',  'Training / Meeting',                  False, True);
INSERT INTO TaskTypes (TaskCode, TaskName, IsCertifiedMailTask, IsActive) VALUES ('ADMIN',     'Administrative Task',                 False, True);


-- -----------------------------------------------------------------------------
-- CertifiedMailTypes
-- -----------------------------------------------------------------------------
INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('STD',    'Standard Certified Mail',                       True);
INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('RR',     'Return Receipt Requested',                      True);
INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('ERR',    'Electronic Return Receipt',                     True);
INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('RD',     'Restricted Delivery',                           True);
INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('RD-RR',  'Restricted Delivery with Return Receipt',        True);
INSERT INTO CertifiedMailTypes (MailTypeCode, MailTypeName, IsActive) VALUES ('SIG',    'Signature Confirmation',                        True);


-- -----------------------------------------------------------------------------
-- CertifiedMailStatuses  (IsFinalStatus = True means no further update expected)
-- -----------------------------------------------------------------------------
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('SENT',      'Sent',                       False, True);
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('IN-TRANS',  'In Transit',                 False, True);
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('DELIV',     'Delivered',                  True,  True);
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('RETURNED',  'Returned to Sender',         True,  True);
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('REFUSED',   'Refused by Recipient',       True,  True);
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('UNCLAIMED', 'Unclaimed / Undeliverable',  True,  True);
INSERT INTO CertifiedMailStatuses (StatusCode, StatusName, IsFinalStatus, IsActive) VALUES ('PENDING',   'Pending Tracking Update',    False, True);


-- -----------------------------------------------------------------------------
-- ReviewStatuses
-- -----------------------------------------------------------------------------
INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('PENDING',   'Pending Review',       True);
INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('APPROVED',  'Approved',             True);
INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('RETURNED',  'Returned for Correction', True);
INSERT INTO ReviewStatuses (StatusCode, StatusName, IsActive) VALUES ('ESCALATED', 'Escalated',            True);


-- -----------------------------------------------------------------------------
-- BacklogStatuses  (IsClosedStatus = True hides item from open backlog report)
-- -----------------------------------------------------------------------------
INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('OPEN',       'Open',                False, True);
INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('IN-PROG',    'In Progress',         False, True);
INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('BLOCKED',    'Blocked',             False, True);
INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('CLOSED',     'Closed / Completed',  True,  True);
INSERT INTO BacklogStatuses (StatusCode, StatusName, IsClosedStatus, IsActive) VALUES ('CANCELLED',  'Cancelled',           True,  True);


-- -----------------------------------------------------------------------------
-- ReportingMonths  (add one row per month before employees begin logging)
-- Pattern: MonthStart = first day, MonthEnd = last day of that month.
-- Extend this list to cover your entire fiscal year.
-- -----------------------------------------------------------------------------
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-01-01#, #2026-01-31#, True);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-02-01#, #2026-02-28#, True);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-03-01#, #2026-03-31#, True);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-04-01#, #2026-04-30#, True);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-05-01#, #2026-05-31#, True);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-06-01#, #2026-06-30#, False);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-07-01#, #2026-07-31#, False);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-08-01#, #2026-08-31#, False);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-09-01#, #2026-09-30#, False);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-10-01#, #2026-10-31#, False);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-11-01#, #2026-11-30#, False);
INSERT INTO ReportingMonths (MonthStart, MonthEnd, IsClosed) VALUES (#2026-12-01#, #2026-12-31#, False);

-- =============================================================================
-- OCSS Employee Productivity Log
-- Schema Script 01 -- Lookup / Reference Tables
-- Run these statements ONE AT A TIME in the Access Query Editor,
-- or execute all at once using the VBA script in vba/Build_Schema.bas.
-- =============================================================================


-- -----------------------------------------------------------------------------
-- Units
-- Organisational units (teams / sections) within OCSS.
-- -----------------------------------------------------------------------------
CREATE TABLE Units (
    UnitID    AUTOINCREMENT CONSTRAINT PK_Units PRIMARY KEY,
    UnitCode  TEXT(20)  NOT NULL,
    UnitName  TEXT(100) NOT NULL,
    IsActive  YESNO     NOT NULL,
    CONSTRAINT UQ_Units_UnitCode UNIQUE (UnitCode)
);


-- -----------------------------------------------------------------------------
-- TaskTypes
-- Catalogue of every work task that can appear on a daily log line.
-- IsCertifiedMailTask = True forces certified-mail fields on the entry form.
-- -----------------------------------------------------------------------------
CREATE TABLE TaskTypes (
    TaskTypeID          AUTOINCREMENT CONSTRAINT PK_TaskTypes PRIMARY KEY,
    TaskCode            TEXT(20)  NOT NULL,
    TaskName            TEXT(100) NOT NULL,
    IsCertifiedMailTask YESNO     NOT NULL,
    IsActive            YESNO     NOT NULL,
    CONSTRAINT UQ_TaskTypes_TaskCode UNIQUE (TaskCode)
);


-- -----------------------------------------------------------------------------
-- CertifiedMailTypes
-- Classifies the kind of certified mail piece (e.g. Standard, Restricted).
-- -----------------------------------------------------------------------------
CREATE TABLE CertifiedMailTypes (
    CertifiedMailTypeID AUTOINCREMENT CONSTRAINT PK_CertifiedMailTypes PRIMARY KEY,
    MailTypeCode        TEXT(20)  NOT NULL,
    MailTypeName        TEXT(100) NOT NULL,
    IsActive            YESNO     NOT NULL,
    CONSTRAINT UQ_CertifiedMailTypes_Code UNIQUE (MailTypeCode)
);


-- -----------------------------------------------------------------------------
-- CertifiedMailStatuses
-- Tracks the delivery outcome of each certified mail piece.
-- IsFinalStatus = True means no further status change is expected.
-- -----------------------------------------------------------------------------
CREATE TABLE CertifiedMailStatuses (
    CertifiedMailStatusID AUTOINCREMENT CONSTRAINT PK_CertifiedMailStatuses PRIMARY KEY,
    StatusCode            TEXT(20)  NOT NULL,
    StatusName            TEXT(100) NOT NULL,
    IsFinalStatus         YESNO     NOT NULL,
    IsActive              YESNO     NOT NULL,
    CONSTRAINT UQ_CertifiedMailStatuses_Code UNIQUE (StatusCode)
);


-- -----------------------------------------------------------------------------
-- ReviewStatuses
-- Supervisor review outcome codes for a daily productivity log.
-- -----------------------------------------------------------------------------
CREATE TABLE ReviewStatuses (
    ReviewStatusID AUTOINCREMENT CONSTRAINT PK_ReviewStatuses PRIMARY KEY,
    StatusCode     TEXT(20)  NOT NULL,
    StatusName     TEXT(100) NOT NULL,
    IsActive       YESNO     NOT NULL,
    CONSTRAINT UQ_ReviewStatuses_Code UNIQUE (StatusCode)
);


-- -----------------------------------------------------------------------------
-- BacklogStatuses
-- Lifecycle states for a backlog work item.
-- IsClosedStatus = True hides the item from the open backlog report.
-- -----------------------------------------------------------------------------
CREATE TABLE BacklogStatuses (
    BacklogStatusID AUTOINCREMENT CONSTRAINT PK_BacklogStatuses PRIMARY KEY,
    StatusCode      TEXT(20)  NOT NULL,
    StatusName      TEXT(100) NOT NULL,
    IsClosedStatus  YESNO     NOT NULL,
    IsActive        YESNO     NOT NULL,
    CONSTRAINT UQ_BacklogStatuses_Code UNIQUE (StatusCode)
);

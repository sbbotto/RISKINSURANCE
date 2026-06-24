-- Run this in SQL Server if you prefer to provision the tables manually.
-- The application can also create them automatically at startup.

IF OBJECT_ID('dbo.CaseAssessmentReport', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CaseAssessmentReport (
        report_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        source_file_name NVARCHAR(260) NOT NULL,
        source_file_path NVARCHAR(1000) NOT NULL UNIQUE,
        source_file_hash CHAR(64) NOT NULL,
        source_last_modified_utc DATETIME2(0) NULL,
        claim_number NVARCHAR(255) NULL,
        claim_manager NVARCHAR(255) NULL,
        report_date NVARCHAR(255) NULL,
        insured NVARCHAR(500) NULL,
        div_hos_notice NVARCHAR(255) NULL,
        date_of_loss NVARCHAR(255) NULL,
        date_of_report NVARCHAR(255) NULL,
        date_of_est NVARCHAR(255) NULL,
        date_of_suit NVARCHAR(255) NULL,
        plaintiff_name NVARCHAR(500) NULL,
        hca_named NVARCHAR(255) NULL,
        jurisdiction NVARCHAR(255) NULL,
        primary_insurance_limit NVARCHAR(255) NULL,
        indemnity_reserve NVARCHAR(255) NULL,
        lae_paid NVARCHAR(255) NULL,
        defense_counsel NVARCHAR(500) NULL,
        defense_firm NVARCHAR(500) NULL,
        plaintiff_counsel NVARCHAR(500) NULL,
        plaintiff_firm NVARCHAR(500) NULL,
        authority_required NVARCHAR(255) NULL,
        demand_offer NVARCHAR(255) NULL,
        trial_date NVARCHAR(255) NULL,
        mediation_date NVARCHAR(255) NULL,
        chance_dv NVARCHAR(255) NULL,
        verdict_value NVARCHAR(255) NULL,
        settlement_value NVARCHAR(255) NULL,
        executive_summary NVARCHAR(MAX) NULL,
        resolution_strategy NVARCHAR(MAX) NULL,
        facts NVARCHAR(MAX) NULL,
        injury NVARCHAR(MAX) NULL,
        plaintiff_section NVARCHAR(MAX) NULL,
        damages NVARCHAR(MAX) NULL,
        allegations NVARCHAR(MAX) NULL,
        defenses NVARCHAR(MAX) NULL,
        peer_review_remediation NVARCHAR(MAX) NULL,
        internal_review NVARCHAR(MAX) NULL,
        experts NVARCHAR(MAX) NULL,
        defense_section NVARCHAR(MAX) NULL,
        raw_json NVARCHAR(MAX) NULL,
        ingested_at_utc DATETIME2(0) NOT NULL CONSTRAINT DF_CaseAssessmentReport_ingested_at DEFAULT SYSUTCDATETIME(),
        last_seen_at_utc DATETIME2(0) NOT NULL CONSTRAINT DF_CaseAssessmentReport_last_seen_at DEFAULT SYSUTCDATETIME()
    );
END;

IF OBJECT_ID('dbo.CaseAssessmentReportProvider', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.CaseAssessmentReportProvider (
        provider_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        report_id INT NOT NULL,
        row_index INT NOT NULL,
        involved_provider_specialty NVARCHAR(500) NULL,
        relationship_to_facility NVARCHAR(500) NULL,
        carrier_limits NVARCHAR(500) NULL,
        prior_hci_claim_involvement NVARCHAR(500) NULL,
        hci_insured_status NVARCHAR(500) NULL,
        raw_text NVARCHAR(MAX) NULL,
        CONSTRAINT FK_CaseAssessmentReportProvider_report FOREIGN KEY (report_id)
            REFERENCES dbo.CaseAssessmentReport(report_id)
            ON DELETE CASCADE
    );
END;

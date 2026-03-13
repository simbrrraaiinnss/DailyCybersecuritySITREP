-- Database schema for Daily Cybersecurity SITREP system

-- Sources table: tracks all intelligence sources
CREATE TABLE IF NOT EXISTS sources (
    source_id SERIAL PRIMARY KEY,
    source_name VARCHAR(255) NOT NULL UNIQUE,
    source_type VARCHAR(50) NOT NULL, -- CISA_CA, CISA_KEV, NEWS, BLOG, VENDOR_ADVISORY, RSS
    source_url TEXT,
    last_checked_at TIMESTAMP,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Intel items table: stores all collected threat intelligence
CREATE TABLE IF NOT EXISTS intel_items (
    item_id SERIAL PRIMARY KEY,
    source_id INTEGER REFERENCES sources(source_id),
    title TEXT NOT NULL,
    summary TEXT,
    url TEXT,
    published_at TIMESTAMP,
    last_seen_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    severity VARCHAR(20), -- CRITICAL, HIGH, MEDIUM, LOW
    cve_list TEXT[], -- Array of CVE IDs
    affected_products TEXT[],
    sector_tags TEXT[], -- healthcare, finance, critical-infrastructure, etc.
    exploitation_status VARCHAR(50), -- KNOWN_EXPLOITED, POC_AVAILABLE, NOT_EXPLOITED, UNKNOWN
    intel_category VARCHAR(50), -- CISA, KEV, NEWS, BLOG, VENDOR_ADVISORY
    relevance_to_hospital VARCHAR(20), -- DIRECT, INDIRECT, LOW
    is_healthcare_relevant BOOLEAN DEFAULT FALSE,
    is_hospital_relevant BOOLEAN DEFAULT FALSE,
    relevance_reason TEXT,
    threat_category VARCHAR(20), -- IMMEDIATE, PRIORITY, ROUTINE
    risk_rating VARCHAR(20), -- HIGH, MED, LOW
    recommended_actions TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- SITREPs table: tracks generated reports
CREATE TABLE IF NOT EXISTS sitreps (
    sitrep_id SERIAL PRIMARY KEY,
    report_date DATE NOT NULL UNIQUE,
    report_id VARCHAR(50) NOT NULL UNIQUE, -- YYYYMMDD-SITREP
    threat_posture VARCHAR(20) NOT NULL, -- ELEVATED, GUARDED, LOW
    document_path TEXT,
    email_sent BOOLEAN DEFAULT FALSE,
    email_sent_at TIMESTAMP,
    immediate_count INTEGER DEFAULT 0,
    priority_count INTEGER DEFAULT 0,
    routine_count INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Actions table: tracks cybersecurity actions taken
CREATE TABLE IF NOT EXISTS actions (
    action_id SERIAL PRIMARY KEY,
    action_date DATE NOT NULL,
    action_type VARCHAR(100), -- PATCHING, CONFIGURATION_CHANGE, INCIDENT_RESPONSE, MONITORING, etc.
    description TEXT NOT NULL,
    system_affected TEXT,
    performed_by VARCHAR(100),
    related_item_id INTEGER REFERENCES intel_items(item_id),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Open items / watch list table
CREATE TABLE IF NOT EXISTS open_items (
    open_item_id SERIAL PRIMARY KEY,
    title TEXT NOT NULL,
    description TEXT,
    status VARCHAR(20) DEFAULT 'OPEN', -- OPEN, IN_PROGRESS, CLOSED
    priority VARCHAR(20), -- HIGH, MEDIUM, LOW
    owner VARCHAR(100),
    due_date DATE,
    related_item_id INTEGER REFERENCES intel_items(item_id),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    closed_at TIMESTAMP
);

-- Deadlines table: tracks upcoming cybersecurity deadlines
CREATE TABLE IF NOT EXISTS deadlines (
    deadline_id SERIAL PRIMARY KEY,
    deadline_date DATE NOT NULL,
    title TEXT NOT NULL,
    description TEXT,
    deadline_type VARCHAR(50), -- VENDOR_ENFORCEMENT, REGULATORY, INTERNAL, PATCH_DEADLINE
    source VARCHAR(100),
    is_completed BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Create indexes for better query performance
CREATE INDEX IF NOT EXISTS idx_intel_items_published_at ON intel_items(published_at DESC);
CREATE INDEX IF NOT EXISTS idx_intel_items_threat_category ON intel_items(threat_category);
CREATE INDEX IF NOT EXISTS idx_intel_items_relevance ON intel_items(relevance_to_hospital);
CREATE INDEX IF NOT EXISTS idx_sitreps_report_date ON sitreps(report_date DESC);
CREATE INDEX IF NOT EXISTS idx_open_items_status ON open_items(status);
CREATE INDEX IF NOT EXISTS idx_deadlines_date ON deadlines(deadline_date);

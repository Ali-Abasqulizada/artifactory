--
-- Artifacts table maintains every artifact known to the artifactory.
--
CREATE TABLE artifacts (
    -- Unique artifact ID
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    -- Name of the artifact. For example, if it's a model database, it's going to be something like: falcon900b
    name TEXT NOT NULL,
    -- Commit hash is the full commit hash from the version control (e.g. git or bitbucket or etc)
    commit_hash TEXT NOT NULL,
    -- Freeform JSON encoded string with tags. Can be anything.
    -- For example: {"deployed": true,...}
    -- Type key has a special meaning, can only be: model_database, icopilot_database or other.
    -- TODO: enforce type values in the incoming JSON.
    tags TEXT,
    -- The timestamp at which the artifact was created
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    -- The combination of the artifact name and commit hash is unique
    UNIQUE (name, commit_hash)
);

-- Artifacts storage table maintains the paths at which the artifacts are stored.
CREATE TABLE artifact_storage (
    -- Unique artifact ID reference from the artifacts table
    artifact_id INTEGER NOT NULL,
    -- The md5 checksum of the artifact.
    checksum TEXT NOT NULL,
    -- Absolute filesystem path at which the artifact is stored
    path TEXT UNIQUE NOT NULL,

    -- The artifact ID maps to ID field in the artifacts table
    FOREIGN KEY (artifact_id) REFERENCES artifacts(id) ON DELETE CASCADE
);

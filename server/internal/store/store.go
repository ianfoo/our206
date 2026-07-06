// Package store is the canonical our206 datastore: SQLite, accessed via
// the pure-Go modernc.org/sqlite driver. Everything else in the system —
// website, calendar, feeds — is a projection of what lives here.
package store

import (
	"database/sql"
	"errors"
	"fmt"
	"strings"
	"time"

	_ "modernc.org/sqlite"
)

const schema = `
CREATE TABLE IF NOT EXISTS venues (
	id               INTEGER PRIMARY KEY,
	name             TEXT NOT NULL,
	address          TEXT NOT NULL DEFAULT '',
	city             TEXT NOT NULL DEFAULT '',
	state            TEXT NOT NULL DEFAULT '',
	postal_code      TEXT NOT NULL DEFAULT '',
	lat              REAL,
	lng              REAL,
	website          TEXT NOT NULL DEFAULT '',
	place_id         TEXT NOT NULL DEFAULT '',
	place_confidence REAL,
	last_checked_at  TEXT,
	created_at       TEXT NOT NULL DEFAULT (datetime('now')),
	updated_at       TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE UNIQUE INDEX IF NOT EXISTS venues_name_ci ON venues (lower(name));

CREATE TABLE IF NOT EXISTS venue_aliases (
	id       INTEGER PRIMARY KEY,
	alias    TEXT NOT NULL,
	venue_id INTEGER NOT NULL REFERENCES venues(id) ON DELETE CASCADE
);
CREATE UNIQUE INDEX IF NOT EXISTS venue_aliases_alias_ci ON venue_aliases (lower(alias));

CREATE TABLE IF NOT EXISTS events (
	id         INTEGER PRIMARY KEY,
	uid        TEXT NOT NULL UNIQUE,
	date       TEXT NOT NULL,             -- YYYY-MM-DD
	artist     TEXT NOT NULL,
	venue_id   INTEGER NOT NULL REFERENCES venues(id),
	score      INTEGER NOT NULL DEFAULT 0 CHECK (score BETWEEN 0 AND 4),
	notes      TEXT NOT NULL DEFAULT '',
	ticket_url TEXT NOT NULL DEFAULT '',
	status     TEXT NOT NULL DEFAULT 'published',  -- published | canceled
	created_at TEXT NOT NULL DEFAULT (datetime('now')),
	updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS events_date ON events (date);

CREATE TABLE IF NOT EXISTS submissions (
	id          INTEGER PRIMARY KEY,
	source      TEXT NOT NULL,            -- paste | web | csv | signal | ...
	submitter   TEXT NOT NULL DEFAULT '',
	raw         TEXT NOT NULL,
	received_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS proposals (
	id            INTEGER PRIMARY KEY,
	submission_id INTEGER NOT NULL REFERENCES submissions(id),
	raw_line      TEXT NOT NULL,
	date          TEXT NOT NULL DEFAULT '',
	artist        TEXT NOT NULL DEFAULT '',
	venue_raw     TEXT NOT NULL DEFAULT '',
	venue_id      INTEGER REFERENCES venues(id),
	score         INTEGER NOT NULL DEFAULT 0,
	state         TEXT NOT NULL DEFAULT 'pending',
	              -- pending | approved | rejected | expired | superseded
	disposition   TEXT NOT NULL DEFAULT '',
	              -- appended | exact_duplicate | fuzzy_duplicate | ignored | needs_review | error
	note          TEXT NOT NULL DEFAULT '',
	event_id      INTEGER REFERENCES events(id),
	created_at    TEXT NOT NULL DEFAULT (datetime('now')),
	resolved_at   TEXT
);
CREATE INDEX IF NOT EXISTS proposals_state ON proposals (state);
`

// Store wraps the SQLite database.
type Store struct {
	db *sql.DB
}

// Open opens (creating if necessary) the database at path and applies the
// schema. Use ":memory:" for tests.
func Open(path string) (*Store, error) {
	dsn := fmt.Sprintf("file:%s?_pragma=journal_mode(WAL)&_pragma=foreign_keys(ON)&_pragma=busy_timeout(5000)", path)
	if path == ":memory:" {
		dsn = "file::memory:?_pragma=foreign_keys(ON)"
	}
	db, err := sql.Open("sqlite", dsn)
	if err != nil {
		return nil, err
	}
	// SQLite handles one writer at a time; a single connection avoids
	// SQLITE_BUSY surprises at this scale.
	db.SetMaxOpenConns(1)
	if _, err := db.Exec(schema); err != nil {
		db.Close()
		return nil, fmt.Errorf("apply schema: %w", err)
	}
	return &Store{db: db}, nil
}

func (s *Store) Close() error { return s.db.Close() }

// Venue is a canonical venue record.
type Venue struct {
	ID      int64
	Name    string
	Address string
}

// Event is a canonical event joined with its venue.
type Event struct {
	ID           int64
	UID          string
	DateKey      string
	Artist       string
	VenueID      int64
	VenueName    string
	VenueAddress string
	Score        int
	Notes        string
	TicketURL    string
	Status       string
}

// ErrNotFound is returned by lookups that match nothing.
var ErrNotFound = errors.New("not found")

// ResolveVenue maps a raw venue string to a canonical venue via the alias
// table, falling back to a case-insensitive match on canonical names.
func (s *Store) ResolveVenue(raw string) (Venue, error) {
	key := strings.ToLower(strings.TrimSpace(raw))
	var v Venue
	err := s.db.QueryRow(`
		SELECT v.id, v.name, v.address FROM venue_aliases a
		JOIN venues v ON v.id = a.venue_id
		WHERE lower(a.alias) = ?`, key).Scan(&v.ID, &v.Name, &v.Address)
	if err == nil {
		return v, nil
	}
	if !errors.Is(err, sql.ErrNoRows) {
		return Venue{}, err
	}
	err = s.db.QueryRow(`SELECT id, name, address FROM venues WHERE lower(name) = ?`, key).
		Scan(&v.ID, &v.Name, &v.Address)
	if errors.Is(err, sql.ErrNoRows) {
		return Venue{}, ErrNotFound
	}
	return v, err
}

// EnsureVenue resolves raw to a canonical venue, creating a new venue named
// exactly as given when nothing matches (the spreadsheet behaved the same
// way: unknown venues passed through verbatim).
func (s *Store) EnsureVenue(raw string) (Venue, bool, error) {
	v, err := s.ResolveVenue(raw)
	if err == nil {
		return v, false, nil
	}
	if !errors.Is(err, ErrNotFound) {
		return Venue{}, false, err
	}
	name := strings.TrimSpace(raw)
	res, err := s.db.Exec(`INSERT INTO venues (name) VALUES (?)`, name)
	if err != nil {
		return Venue{}, false, err
	}
	id, _ := res.LastInsertId()
	return Venue{ID: id, Name: name}, true, nil
}

// AddAlias registers alias -> venueID, ignoring duplicates.
func (s *Store) AddAlias(alias string, venueID int64) error {
	_, err := s.db.Exec(`INSERT OR IGNORE INTO venue_aliases (alias, venue_id) VALUES (?, ?)`,
		strings.ToLower(strings.TrimSpace(alias)), venueID)
	return err
}

// InsertEvent stores a new canonical event.
func (s *Store) InsertEvent(e Event) (int64, error) {
	res, err := s.db.Exec(`
		INSERT INTO events (uid, date, artist, venue_id, score, notes, ticket_url, status)
		VALUES (?, ?, ?, ?, ?, ?, ?, 'published')`,
		e.UID, e.DateKey, e.Artist, e.VenueID, e.Score, e.Notes, e.TicketURL)
	if err != nil {
		return 0, err
	}
	return res.LastInsertId()
}

// ListEvents returns published events with date in [fromKey, toKey]
// (inclusive; either may be empty for open-ended), ordered by date then
// artist.
func (s *Store) ListEvents(fromKey, toKey string) ([]Event, error) {
	q := `SELECT e.id, e.uid, e.date, e.artist, e.venue_id, v.name, v.address,
	             e.score, e.notes, e.ticket_url, e.status
	      FROM events e JOIN venues v ON v.id = e.venue_id
	      WHERE e.status = 'published'`
	var args []any
	if fromKey != "" {
		q += ` AND e.date >= ?`
		args = append(args, fromKey)
	}
	if toKey != "" {
		q += ` AND e.date <= ?`
		args = append(args, toKey)
	}
	q += ` ORDER BY e.date, e.artist`
	rows, err := s.db.Query(q, args...)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	var out []Event
	for rows.Next() {
		var e Event
		if err := rows.Scan(&e.ID, &e.UID, &e.DateKey, &e.Artist, &e.VenueID, &e.VenueName,
			&e.VenueAddress, &e.Score, &e.Notes, &e.TicketURL, &e.Status); err != nil {
			return nil, err
		}
		out = append(out, e)
	}
	return out, rows.Err()
}

// ListVenues returns all canonical venues ordered by name.
func (s *Store) ListVenues() ([]Venue, error) {
	rows, err := s.db.Query(`SELECT id, name, address FROM venues ORDER BY name`)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	var out []Venue
	for rows.Next() {
		var v Venue
		if err := rows.Scan(&v.ID, &v.Name, &v.Address); err != nil {
			return nil, err
		}
		out = append(out, v)
	}
	return out, rows.Err()
}

// CreateSubmission records one intake action and returns its id.
func (s *Store) CreateSubmission(source, submitter, raw string) (int64, error) {
	res, err := s.db.Exec(`INSERT INTO submissions (source, submitter, raw) VALUES (?, ?, ?)`,
		source, submitter, raw)
	if err != nil {
		return 0, err
	}
	return res.LastInsertId()
}

// Proposal mirrors one row of the proposals table.
type Proposal struct {
	ID           int64
	SubmissionID int64
	RawLine      string
	DateKey      string
	Artist       string
	VenueRaw     string
	VenueID      sql.NullInt64
	Score        int
	State        string
	Disposition  string
	Note         string
	EventID      sql.NullInt64
}

// CreateProposal inserts a proposal; resolved states get a resolved_at
// timestamp.
func (s *Store) CreateProposal(p Proposal) (int64, error) {
	var resolvedAt any
	if p.State != "pending" {
		resolvedAt = time.Now().UTC().Format(time.RFC3339)
	}
	res, err := s.db.Exec(`
		INSERT INTO proposals
			(submission_id, raw_line, date, artist, venue_raw, venue_id, score,
			 state, disposition, note, event_id, resolved_at)
		VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
		p.SubmissionID, p.RawLine, p.DateKey, p.Artist, p.VenueRaw, p.VenueID, p.Score,
		p.State, p.Disposition, p.Note, p.EventID, resolvedAt)
	if err != nil {
		return 0, err
	}
	return res.LastInsertId()
}

// PendingProposals lists proposals awaiting review, oldest first.
func (s *Store) PendingProposals() ([]Proposal, error) {
	rows, err := s.db.Query(`
		SELECT id, submission_id, raw_line, date, artist, venue_raw, venue_id,
		       score, state, disposition, note, event_id
		FROM proposals WHERE state = 'pending' ORDER BY id`)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	var out []Proposal
	for rows.Next() {
		var p Proposal
		if err := rows.Scan(&p.ID, &p.SubmissionID, &p.RawLine, &p.DateKey, &p.Artist,
			&p.VenueRaw, &p.VenueID, &p.Score, &p.State, &p.Disposition, &p.Note, &p.EventID); err != nil {
			return nil, err
		}
		out = append(out, p)
	}
	return out, rows.Err()
}

// ExpirePastPending marks pending proposals whose date has passed as
// expired, returning how many were affected.
func (s *Store) ExpirePastPending(todayKey string) (int64, error) {
	res, err := s.db.Exec(`
		UPDATE proposals SET state = 'expired', resolved_at = datetime('now')
		WHERE state = 'pending' AND date != '' AND date < ?`, todayKey)
	if err != nil {
		return 0, err
	}
	return res.RowsAffected()
}

// DedupKeys returns the exact and fuzzy duplicate-detection key sets for
// all events, matching the spreadsheet's scheme:
//
//	exact: date|venueID|artistCompareKey
//	fuzzy: date|venueID|primaryArtistKey
//
// The caller supplies the key functions to keep this package free of
// normalization logic.
func (s *Store) DedupKeys(compareKey, primaryKey func(string) string) (exact, fuzzy map[string]bool, err error) {
	rows, err := s.db.Query(`SELECT date, venue_id, artist FROM events`)
	if err != nil {
		return nil, nil, err
	}
	defer rows.Close()
	exact = make(map[string]bool)
	fuzzy = make(map[string]bool)
	for rows.Next() {
		var date, artist string
		var venueID int64
		if err := rows.Scan(&date, &venueID, &artist); err != nil {
			return nil, nil, err
		}
		exact[fmt.Sprintf("%s|%d|%s", date, venueID, compareKey(artist))] = true
		fuzzy[fmt.Sprintf("%s|%d|%s", date, venueID, primaryKey(artist))] = true
	}
	return exact, fuzzy, rows.Err()
}

// Package importer runs intake batches through the editorial pipeline:
// parse -> normalize -> venue resolution -> duplicate detection ->
// canonical events plus a proposal record for every line.
//
// It preserves the spreadsheet's dispositions with one deliberate change:
// fuzzy duplicates used to be silently dropped; here they become pending
// proposals so an editor decides ("never silently merge").
package importer

import (
	"database/sql"
	"encoding/csv"
	"fmt"
	"io"
	"strings"
	"time"

	"github.com/ianfoo/our206/server/internal/normalize"
	"github.com/ianfoo/our206/server/internal/parse"
	"github.com/ianfoo/our206/server/internal/store"
	"github.com/ianfoo/our206/server/internal/uid"
)

// Result summarizes one processed batch.
type Result struct {
	SubmissionID int64
	Appended     int
	ExactDups    int
	NeedsReview  int
	Ignored      int
	Notes        []string
}

func (r Result) String() string {
	return fmt.Sprintf("appended=%d exact_duplicates=%d needs_review=%d ignored=%d",
		r.Appended, r.ExactDups, r.NeedsReview, r.Ignored)
}

// Shorthand processes a batch of contributor shorthand lines. ref anchors
// year inference (normally time.Now()).
func Shorthand(st *store.Store, source, submitter, text string, ref time.Time) (Result, error) {
	subID, err := st.CreateSubmission(source, submitter, text)
	if err != nil {
		return Result{}, err
	}
	res := Result{SubmissionID: subID}

	exact, fuzzy, err := st.DedupKeys(normalize.ArtistCompareKey, normalize.PrimaryArtistKey)
	if err != nil {
		return res, err
	}

	for _, line := range strings.Split(text, "\n") {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}
		if err := processLine(st, subID, line, ref, exact, fuzzy, &res); err != nil {
			return res, err
		}
	}
	return res, nil
}

func processLine(st *store.Store, subID int64, line string, ref time.Time,
	exact, fuzzy map[string]bool, res *Result) error {

	parsed, ok := parse.Shorthand(line, ref)
	if !ok {
		res.Ignored++
		_, err := st.CreateProposal(store.Proposal{
			SubmissionID: subID, RawLine: line,
			State: "rejected", Disposition: "ignored",
			Note: "could not parse (expected: M/D: Artist @ Venue)",
		})
		return err
	}

	venue, created, err := st.EnsureVenue(parsed.Venue)
	if err != nil {
		return err
	}
	if created {
		res.Notes = append(res.Notes, fmt.Sprintf("new venue created: %q", venue.Name))
	}
	artist := normalize.DeShoutifyArtist(parsed.Artist)
	if artist != parsed.Artist {
		res.Notes = append(res.Notes, fmt.Sprintf("artist: %q -> %q", parsed.Artist, artist))
	}
	if venue.Name != parsed.Venue {
		res.Notes = append(res.Notes, fmt.Sprintf("venue: %q -> %q", parsed.Venue, venue.Name))
	}

	exactKey := fmt.Sprintf("%s|%d|%s", parsed.DateKey, venue.ID, normalize.ArtistCompareKey(artist))
	fuzzyKey := fmt.Sprintf("%s|%d|%s", parsed.DateKey, venue.ID, normalize.PrimaryArtistKey(artist))

	base := store.Proposal{
		SubmissionID: subID, RawLine: line,
		DateKey: parsed.DateKey, Artist: artist, VenueRaw: parsed.Venue,
		VenueID: sql.NullInt64{Int64: venue.ID, Valid: true}, Score: parsed.Score,
	}

	switch {
	case exact[exactKey]:
		res.ExactDups++
		base.State, base.Disposition = "rejected", "exact_duplicate"
		base.Note = "exact duplicate of an existing event"
		_, err = st.CreateProposal(base)
		return err

	case fuzzy[fuzzyKey]:
		res.NeedsReview++
		base.State, base.Disposition = "pending", "fuzzy_duplicate"
		base.Note = "same date/venue/primary artist as an existing event — needs editor review"
		_, err = st.CreateProposal(base)
		return err

	default:
		eventID, err := st.InsertEvent(store.Event{
			UID:     uid.Build(parsed.DateKey, artist, venue.Name),
			DateKey: parsed.DateKey,
			Artist:  artist,
			VenueID: venue.ID,
			Score:   parsed.Score,
		})
		if err != nil {
			return err
		}
		res.Appended++
		exact[exactKey], fuzzy[fuzzyKey] = true, true
		base.State, base.Disposition = "approved", "appended"
		base.EventID = sql.NullInt64{Int64: eventID, Valid: true}
		_, err = st.CreateProposal(base)
		return err
	}
}

// CSV imports a spreadsheet export with the prototype's column layout:
//
//	Date, Artist, Venue, Skoi Rating, Notes, Ticket Link[, ...]
//
// A header row is detected and skipped. Dates may be DD-MMM-YYYY, M/D/YYYY,
// or YYYY-MM-DD. Ratings count ✅/!/🔥 marks. Rows import directly as
// published events (this is the trusted one-time migration path), with
// exact duplicates skipped.
func CSV(st *store.Store, source string, r io.Reader) (Result, error) {
	subID, err := st.CreateSubmission(source, "", "(csv import)")
	if err != nil {
		return Result{}, err
	}
	res := Result{SubmissionID: subID}

	exact, fuzzy, err := st.DedupKeys(normalize.ArtistCompareKey, normalize.PrimaryArtistKey)
	if err != nil {
		return res, err
	}

	cr := csv.NewReader(r)
	cr.FieldsPerRecord = -1
	for {
		rec, err := cr.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return res, err
		}
		if len(rec) < 3 {
			continue
		}
		get := func(i int) string {
			if i < len(rec) {
				return strings.TrimSpace(rec[i])
			}
			return ""
		}
		dateKey := parse.DateKey(get(0))
		artist, venueRaw := get(1), get(2)
		if dateKey == "" || artist == "" || venueRaw == "" {
			res.Ignored++ // header row or malformed row
			continue
		}

		venue, _, err := st.EnsureVenue(venueRaw)
		if err != nil {
			return res, err
		}
		artist = normalize.DeShoutifyArtist(artist)
		exactKey := fmt.Sprintf("%s|%d|%s", dateKey, venue.ID, normalize.ArtistCompareKey(artist))
		if exact[exactKey] {
			res.ExactDups++
			continue
		}

		eventID, err := st.InsertEvent(store.Event{
			UID:       uid.Build(dateKey, artist, venue.Name),
			DateKey:   dateKey,
			Artist:    artist,
			VenueID:   venue.ID,
			Score:     normalize.ScoreFromMarks(get(3)),
			Notes:     get(4),
			TicketURL: get(5),
		})
		if err != nil {
			return res, err
		}
		res.Appended++
		exact[exactKey] = true
		fuzzy[fmt.Sprintf("%s|%d|%s", dateKey, venue.ID, normalize.PrimaryArtistKey(artist))] = true
		_ = eventID
	}
	return res, nil
}

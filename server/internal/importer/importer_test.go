package importer

import (
	"strings"
	"testing"
	"time"

	"github.com/ianfoo/our206/server/internal/store"
)

var ref = time.Date(2026, 7, 6, 10, 0, 0, 0, time.Local)

func newStore(t *testing.T) *store.Store {
	t.Helper()
	st, err := store.Open(":memory:")
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { st.Close() })
	if err := st.Seed(); err != nil {
		t.Fatal(err)
	}
	return st
}

func TestShorthandBatch(t *testing.T) {
	st := newStore(t)
	batch := strings.Join([]string{
		"8/20: UMPHREY'S MCGEE @ Showbox ✅✅✅✅",
		"8/21: Band of Horses @ Showbox",
		"8/24: Machine Girl @ SoDo Showbox ✅✅",
		"just some chatter",
	}, "\n")

	res, err := Shorthand(st, "paste", "tester", batch, ref)
	if err != nil {
		t.Fatal(err)
	}
	if res.Appended != 3 || res.Ignored != 1 || res.ExactDups != 0 || res.NeedsReview != 0 {
		t.Fatalf("unexpected result: %+v", res)
	}

	events, err := st.ListEvents("", "")
	if err != nil {
		t.Fatal(err)
	}
	if len(events) != 3 {
		t.Fatalf("got %d events, want 3", len(events))
	}
	// Venue aliases resolved to canonical names; artist de-shoutified.
	if events[0].Artist != "Umphrey's Mcgee" || events[0].VenueName != "The Showbox" {
		t.Errorf("normalization failed: %q @ %q", events[0].Artist, events[0].VenueName)
	}
	if events[2].VenueName != "Showbox SoDo" {
		t.Errorf("alias 'SoDo Showbox' resolved to %q, want 'Showbox SoDo'", events[2].VenueName)
	}
	if events[0].Score != 4 {
		t.Errorf("score = %d, want 4", events[0].Score)
	}
}

func TestExactDuplicateRejected(t *testing.T) {
	st := newStore(t)
	if _, err := Shorthand(st, "paste", "", "8/20: Some Band @ Neumos", ref); err != nil {
		t.Fatal(err)
	}
	// Same event, different formatting: still an exact duplicate.
	res, err := Shorthand(st, "paste", "", "8/20 SOME BAND @ neumos ✅", ref)
	if err != nil {
		t.Fatal(err)
	}
	if res.ExactDups != 1 || res.Appended != 0 {
		t.Fatalf("unexpected result: %+v", res)
	}
	events, _ := st.ListEvents("", "")
	if len(events) != 1 {
		t.Fatalf("got %d events, want 1", len(events))
	}
}

func TestFuzzyDuplicateNeedsReview(t *testing.T) {
	st := newStore(t)
	if _, err := Shorthand(st, "paste", "", "8/20: Tom Hamilton @ Neumos", ref); err != nil {
		t.Fatal(err)
	}
	res, err := Shorthand(st, "paste", "", "8/20: Tom Hamilton x Swindler @ Neumos", ref)
	if err != nil {
		t.Fatal(err)
	}
	if res.NeedsReview != 1 || res.Appended != 0 {
		t.Fatalf("unexpected result: %+v", res)
	}
	// No event was created; a pending proposal awaits an editor.
	events, _ := st.ListEvents("", "")
	if len(events) != 1 {
		t.Fatalf("got %d events, want 1", len(events))
	}
	pending, err := st.PendingProposals()
	if err != nil {
		t.Fatal(err)
	}
	if len(pending) != 1 || pending[0].Disposition != "fuzzy_duplicate" {
		t.Fatalf("unexpected pending proposals: %+v", pending)
	}
}

func TestUnknownVenueCreated(t *testing.T) {
	st := newStore(t)
	res, err := Shorthand(st, "paste", "", "8/20: New Band @ Some Basement", ref)
	if err != nil {
		t.Fatal(err)
	}
	if res.Appended != 1 {
		t.Fatalf("unexpected result: %+v", res)
	}
	v, err := st.ResolveVenue("some basement")
	if err != nil {
		t.Fatalf("venue not created: %v", err)
	}
	if v.Name != "Some Basement" {
		t.Errorf("venue name = %q", v.Name)
	}
}

func TestCSVImport(t *testing.T) {
	st := newStore(t)
	csv := `Date,Artist,Venue,Skoi Rating,Notes,Ticket Link
20-Mar-2026,Umphrey's McGee,Showbox,🔥🔥🔥🔥,epic,https://tickets.example/umph
3/21/2026,Band of Horses,The Showbox,,,
2026-03-24,Machine Girl,SoDo Showbox,✅✅,,
`
	res, err := CSV(st, "csv", strings.NewReader(csv))
	if err != nil {
		t.Fatal(err)
	}
	if res.Appended != 3 || res.Ignored != 1 { // header row ignored
		t.Fatalf("unexpected result: %+v", res)
	}
	events, _ := st.ListEvents("", "")
	if len(events) != 3 {
		t.Fatalf("got %d events, want 3", len(events))
	}
	if events[0].Score != 4 || events[0].Notes != "epic" || events[0].TicketURL == "" {
		t.Errorf("row fields not imported: %+v", events[0])
	}
	// Both "Showbox" and "The Showbox" resolve to the same canonical venue.
	if events[0].VenueID != events[1].VenueID {
		t.Error("Showbox and The Showbox did not resolve to the same venue")
	}
}

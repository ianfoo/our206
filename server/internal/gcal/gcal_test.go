package gcal

import (
	"fmt"
	"testing"
	"time"

	"github.com/ianfoo/our206/server/internal/store"
)

// fakeClient records mutations against an in-memory event list.
type fakeClient struct {
	events  []Event
	created []Event
	updated []Event
	deleted []string
	nextID  int
}

func (f *fakeClient) List(_, _ time.Time) ([]Event, error) { return f.events, nil }
func (f *fakeClient) Create(e Event) error {
	f.nextID++
	e.ExternalID = fmt.Sprintf("ext%d", f.nextID)
	f.created = append(f.created, e)
	return nil
}
func (f *fakeClient) Update(e Event) error { f.updated = append(f.updated, e); return nil }
func (f *fakeClient) Delete(id string) error {
	f.deleted = append(f.deleted, id)
	return nil
}

func TestUIDMarkerRoundTrip(t *testing.T) {
	desc := AttachUID("abc123def456abc123def456", "Great show\nTicket link: x")
	if got := ExtractUID(desc); got != "abc123def456abc123def456" {
		t.Errorf("ExtractUID = %q", got)
	}
	if ExtractUID("no marker here") != "" {
		t.Error("ExtractUID found a UID where none exists")
	}
	// Marker-only description (no user text).
	if got := ExtractUID(AttachUID("abcdef1234567890abcdef12", "")); got != "abcdef1234567890abcdef12" {
		t.Errorf("ExtractUID marker-only = %q", got)
	}
}

func TestReconcile(t *testing.T) {
	desired := Desired([]store.Event{
		{UID: "uid_keep_unchanged_000001", DateKey: "2026-08-01", Artist: "Keeper",
			VenueName: "Neumos", VenueAddress: "925 E Pike St, Seattle, WA 98122"},
		{UID: "uid_needs_update_0000002", DateKey: "2026-08-02", Artist: "Changed Band",
			VenueName: "Tractor Tavern", Score: 2},
		{UID: "uid_to_create_000000003", DateKey: "2026-08-03", Artist: "New Band",
			VenueName: "The Crocodile"},
	})

	fake := &fakeClient{
		events: []Event{
			// Unchanged: mirror the desired representation exactly.
			desired[0].withExternalID("ext-keep"),
			// Same UID but stale title.
			{ExternalID: "ext-update", DateKey: "2026-08-02", Title: "Old Name",
				Location: "Tractor Tavern", Description: desired[1].Description},
			// Orphan: tagged with a UID no longer desired.
			{ExternalID: "ext-orphan", DateKey: "2026-08-04", Title: "Cancelled Band",
				Description: AttachUID("uid_orphaned_00000000004", "")},
			// Human-created event with no marker: must never be touched.
			{ExternalID: "ext-human", DateKey: "2026-08-05", Title: "Someone's Birthday"},
		},
	}

	sum, err := Reconcile(fake, desired, time.Now(), time.Now().AddDate(2, 0, 0), false)
	if err != nil {
		t.Fatal(err)
	}
	if sum.Created != 1 || sum.Updated != 1 || sum.Deleted != 1 {
		t.Fatalf("summary = %+v", sum)
	}
	if len(fake.created) != 1 || fake.created[0].Title != "New Band" {
		t.Errorf("created = %+v", fake.created)
	}
	if len(fake.updated) != 1 || fake.updated[0].ExternalID != "ext-update" ||
		fake.updated[0].Title != "Changed Band" {
		t.Errorf("updated = %+v", fake.updated)
	}
	if len(fake.deleted) != 1 || fake.deleted[0] != "ext-orphan" {
		t.Errorf("deleted = %+v", fake.deleted)
	}
}

func TestReconcileDryRunTouchesNothing(t *testing.T) {
	desired := Desired([]store.Event{
		{UID: "uid_to_create_000000001", DateKey: "2026-08-03", Artist: "New Band", VenueName: "Neumos"},
	})
	fake := &fakeClient{}
	sum, err := Reconcile(fake, desired, time.Now(), time.Now().AddDate(2, 0, 0), true)
	if err != nil {
		t.Fatal(err)
	}
	if sum.Created != 1 {
		t.Fatalf("summary = %+v", sum)
	}
	if len(fake.created)+len(fake.updated)+len(fake.deleted) != 0 {
		t.Error("dry run mutated the calendar")
	}
}

func TestDesiredDescription(t *testing.T) {
	d := Desired([]store.Event{{
		UID: "u", DateKey: "2026-08-01", Artist: "Band",
		VenueName: "Neumos", VenueAddress: "925 E Pike St",
		Score: 2, Notes: "great openers", TicketURL: "https://t.example",
	}})[0]
	want := "great openers\nSkoi rating: 🔥🔥\nTicket link: https://t.example\n\n" + UIDMarkerPrefix + "u"
	if d.Description != want {
		t.Errorf("description = %q, want %q", d.Description, want)
	}
	if d.Location != "Neumos\n925 E Pike St" {
		t.Errorf("location = %q", d.Location)
	}
}

// withExternalID clones a desired event as an existing calendar event.
func (e Event) withExternalID(id string) Event {
	e.ExternalID = id
	return e
}

// Package gcal synchronizes canonical events to Google Calendar, one-way.
// It preserves the Apps Script prototype's reconciliation exactly: desired
// events are derived from the canonical store, existing calendar events are
// matched by a hidden "[our206_uid]:<uid>" marker in their description, and
// the calendar is made to match (create missing, update changed, delete
// orphans). Because the UID scheme and marker are unchanged, this reconciles
// cleanly against calendar entries created by the spreadsheet system.
package gcal

import (
	"fmt"
	"regexp"
	"strings"
	"time"

	"github.com/ianfoo/our206/server/internal/normalize"
	"github.com/ianfoo/our206/server/internal/store"
)

// UIDMarkerPrefix matches CFG.uidMarkerPrefix in the prototype.
const UIDMarkerPrefix = "[our206_uid]:"

// HorizonYears bounds how far ahead the sync window extends.
const HorizonYears = 2

var uidMarkerRe = regexp.MustCompile(regexp.QuoteMeta(UIDMarkerPrefix) + `(\w{16,64})`)

// Event is a calendar-side event (all-day).
type Event struct {
	ExternalID  string // provider event id; empty for desired events
	UID         string
	Title       string
	Location    string
	Description string // includes the UID marker
	DateKey     string // YYYY-MM-DD
}

// Client abstracts the calendar provider so reconciliation is testable.
type Client interface {
	List(from, to time.Time) ([]Event, error)
	Create(e Event) error
	Update(e Event) error
	Delete(externalID string) error
}

// AttachUID appends the hidden UID marker to a user-facing description.
func AttachUID(uid, description string) string {
	d := strings.TrimSpace(description)
	if d == "" {
		return UIDMarkerPrefix + uid
	}
	return d + "\n\n" + UIDMarkerPrefix + uid
}

// ExtractUID pulls the UID marker out of a calendar event description,
// returning "" when absent.
func ExtractUID(description string) string {
	m := uidMarkerRe.FindStringSubmatch(description)
	if m == nil {
		return ""
	}
	return m[1]
}

// Desired converts canonical events into their calendar representation:
// title is the artist, location is "Venue\nAddress", description carries
// notes, rating flames, ticket link, and the UID marker.
func Desired(events []store.Event) []Event {
	out := make([]Event, 0, len(events))
	for _, e := range events {
		var desc []string
		if e.Notes != "" {
			desc = append(desc, e.Notes)
		}
		if e.Score > 0 {
			desc = append(desc, "Skoi rating: "+normalize.Flames(e.Score))
		}
		if e.TicketURL != "" {
			desc = append(desc, "Ticket link: "+e.TicketURL)
		}
		location := e.VenueName
		if e.VenueAddress != "" {
			location += "\n" + e.VenueAddress
		}
		out = append(out, Event{
			UID:         e.UID,
			Title:       e.Artist,
			Location:    location,
			Description: AttachUID(e.UID, strings.Join(desc, "\n")),
			DateKey:     e.DateKey,
		})
	}
	return out
}

// Summary reports what a reconciliation did (or would do, when dry).
type Summary struct {
	Created, Updated, Deleted int
	Desired, ExistingTagged   int
	Log                       []string
}

func (s Summary) String() string {
	return fmt.Sprintf("created=%d updated=%d deleted=%d desired=%d existing_tagged=%d",
		s.Created, s.Updated, s.Deleted, s.Desired, s.ExistingTagged)
}

// Reconcile makes the calendar window [from, to] match desired. Calendar
// events without a UID marker are never touched (they belong to humans).
func Reconcile(c Client, desired []Event, from, to time.Time, dryRun bool) (Summary, error) {
	var sum Summary
	sum.Desired = len(desired)

	existing, err := c.List(from, to)
	if err != nil {
		return sum, fmt.Errorf("list calendar events: %w", err)
	}
	existingByUID := make(map[string]Event)
	for _, ev := range existing {
		if uid := ExtractUID(ev.Description); uid != "" {
			ev.UID = uid
			existingByUID[uid] = ev
		}
	}
	sum.ExistingTagged = len(existingByUID)

	for _, d := range desired {
		ev, ok := existingByUID[d.UID]
		if !ok {
			sum.Created++
			sum.Log = append(sum.Log, fmt.Sprintf("CREATE %s — %s @ %s", d.DateKey, d.Title, firstLine(d.Location)))
			if !dryRun {
				if err := c.Create(d); err != nil {
					return sum, fmt.Errorf("create %q: %w", d.Title, err)
				}
			}
			continue
		}
		if ev.Title != d.Title || ev.Location != d.Location ||
			ev.Description != d.Description || ev.DateKey != d.DateKey {
			sum.Updated++
			sum.Log = append(sum.Log, fmt.Sprintf("UPDATE %s — %s @ %s", d.DateKey, d.Title, firstLine(d.Location)))
			if !dryRun {
				d.ExternalID = ev.ExternalID
				if err := c.Update(d); err != nil {
					return sum, fmt.Errorf("update %q: %w", d.Title, err)
				}
			}
		}
	}

	desiredByUID := make(map[string]bool, len(desired))
	for _, d := range desired {
		desiredByUID[d.UID] = true
	}
	for uid, ev := range existingByUID {
		if desiredByUID[uid] {
			continue
		}
		sum.Deleted++
		sum.Log = append(sum.Log, fmt.Sprintf("DELETE %s — %s", ev.DateKey, ev.Title))
		if !dryRun {
			if err := c.Delete(ev.ExternalID); err != nil {
				return sum, fmt.Errorf("delete %q: %w", ev.Title, err)
			}
		}
	}
	return sum, nil
}

func firstLine(s string) string {
	if i := strings.IndexByte(s, '\n'); i >= 0 {
		return s[:i]
	}
	return s
}

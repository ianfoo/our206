package gcal

import (
	"context"
	"fmt"
	"time"

	"google.golang.org/api/calendar/v3"
	"google.golang.org/api/option"
)

// GoogleClient implements Client against the Google Calendar API. It
// authenticates with Application Default Credentials — set
// GOOGLE_APPLICATION_CREDENTIALS to a service-account key file and share
// the target calendar with the service account's email address (with
// "Make changes to events" permission).
type GoogleClient struct {
	svc        *calendar.Service
	calendarID string
}

// NewGoogleClient builds a calendar client for calendarID.
func NewGoogleClient(ctx context.Context, calendarID string, opts ...option.ClientOption) (*GoogleClient, error) {
	opts = append([]option.ClientOption{option.WithScopes(calendar.CalendarEventsScope)}, opts...)
	svc, err := calendar.NewService(ctx, opts...)
	if err != nil {
		return nil, fmt.Errorf("create calendar service: %w", err)
	}
	return &GoogleClient{svc: svc, calendarID: calendarID}, nil
}

func (g *GoogleClient) List(from, to time.Time) ([]Event, error) {
	var out []Event
	pageToken := ""
	for {
		call := g.svc.Events.List(g.calendarID).
			TimeMin(from.Format(time.RFC3339)).
			TimeMax(to.Format(time.RFC3339)).
			SingleEvents(true).
			MaxResults(2500)
		if pageToken != "" {
			call = call.PageToken(pageToken)
		}
		resp, err := call.Do()
		if err != nil {
			return nil, err
		}
		for _, item := range resp.Items {
			ev := Event{
				ExternalID:  item.Id,
				Title:       item.Summary,
				Location:    item.Location,
				Description: item.Description,
			}
			if item.Start != nil && item.Start.Date != "" {
				ev.DateKey = item.Start.Date
			}
			out = append(out, ev)
		}
		pageToken = resp.NextPageToken
		if pageToken == "" {
			return out, nil
		}
	}
}

func (g *GoogleClient) Create(e Event) error {
	_, err := g.svc.Events.Insert(g.calendarID, toGoogleEvent(e)).Do()
	return err
}

func (g *GoogleClient) Update(e Event) error {
	_, err := g.svc.Events.Update(g.calendarID, e.ExternalID, toGoogleEvent(e)).Do()
	return err
}

func (g *GoogleClient) Delete(externalID string) error {
	return g.svc.Events.Delete(g.calendarID, externalID).Do()
}

// toGoogleEvent renders an all-day event; per the Calendar API, End.Date is
// exclusive, so it is the day after Start.Date.
func toGoogleEvent(e Event) *calendar.Event {
	start, _ := time.Parse("2006-01-02", e.DateKey)
	return &calendar.Event{
		Summary:     e.Title,
		Location:    e.Location,
		Description: e.Description,
		Start:       &calendar.EventDateTime{Date: e.DateKey},
		End:         &calendar.EventDateTime{Date: start.AddDate(0, 0, 1).Format("2006-01-02")},
	}
}

// Package web exposes the Phase 1 HTTP surface: a small JSON API, an
// import endpoint, and a minimal server-rendered event listing that stands
// in for the real website until Phase 2.
package web

import (
	"crypto/subtle"
	"encoding/json"
	"html/template"
	"io"
	"log"
	"net/http"
	"time"

	"github.com/ianfoo/our206/server/internal/gcal"
	"github.com/ianfoo/our206/server/internal/importer"
	"github.com/ianfoo/our206/server/internal/normalize"
	"github.com/ianfoo/our206/server/internal/store"
)

// Server carries handler dependencies.
type Server struct {
	store      *store.Store
	adminToken string // empty disables mutating endpoints
}

// New builds the HTTP handler.
func New(st *store.Store, adminToken string) http.Handler {
	s := &Server{store: st, adminToken: adminToken}
	mux := http.NewServeMux()
	mux.HandleFunc("GET /healthz", func(w http.ResponseWriter, _ *http.Request) {
		w.Write([]byte("ok"))
	})
	mux.HandleFunc("GET /api/events", s.listEvents)
	mux.HandleFunc("GET /api/venues", s.listVenues)
	mux.HandleFunc("POST /api/import", s.importShorthand)
	mux.HandleFunc("GET /{$}", s.home)
	return mux
}

func todayKey() string { return time.Now().Format("2006-01-02") }

func horizonKey() string {
	return time.Now().AddDate(gcal.HorizonYears, 0, 0).Format("2006-01-02")
}

type eventJSON struct {
	UID       string `json:"uid"`
	Date      string `json:"date"`
	Artist    string `json:"artist"`
	Venue     string `json:"venue"`
	Address   string `json:"address,omitempty"`
	Score     int    `json:"score"`
	Flames    string `json:"flames,omitempty"`
	Notes     string `json:"notes,omitempty"`
	TicketURL string `json:"ticket_url,omitempty"`
}

func (s *Server) listEvents(w http.ResponseWriter, r *http.Request) {
	from := r.URL.Query().Get("from")
	to := r.URL.Query().Get("to")
	if from == "" {
		from = todayKey()
	}
	if to == "" {
		to = horizonKey()
	}
	events, err := s.store.ListEvents(from, to)
	if err != nil {
		httpError(w, err)
		return
	}
	out := make([]eventJSON, 0, len(events))
	for _, e := range events {
		out = append(out, eventJSON{
			UID: e.UID, Date: e.DateKey, Artist: e.Artist,
			Venue: e.VenueName, Address: e.VenueAddress,
			Score: e.Score, Flames: normalize.Flames(e.Score),
			Notes: e.Notes, TicketURL: e.TicketURL,
		})
	}
	writeJSON(w, map[string]any{"events": out})
}

func (s *Server) listVenues(w http.ResponseWriter, _ *http.Request) {
	venues, err := s.store.ListVenues()
	if err != nil {
		httpError(w, err)
		return
	}
	type venueJSON struct {
		Name    string `json:"name"`
		Address string `json:"address,omitempty"`
	}
	out := make([]venueJSON, 0, len(venues))
	for _, v := range venues {
		out = append(out, venueJSON{Name: v.Name, Address: v.Address})
	}
	writeJSON(w, map[string]any{"venues": out})
}

// importShorthand accepts a text/plain body of shorthand lines. It requires
// the admin bearer token; with no token configured the endpoint is
// disabled entirely rather than left open.
func (s *Server) importShorthand(w http.ResponseWriter, r *http.Request) {
	if s.adminToken == "" {
		http.Error(w, "import disabled: no admin token configured", http.StatusForbidden)
		return
	}
	auth := r.Header.Get("Authorization")
	if subtle.ConstantTimeCompare([]byte(auth), []byte("Bearer "+s.adminToken)) != 1 {
		http.Error(w, "unauthorized", http.StatusUnauthorized)
		return
	}
	body, err := io.ReadAll(io.LimitReader(r.Body, 1<<20))
	if err != nil {
		httpError(w, err)
		return
	}
	res, err := importer.Shorthand(s.store, "web", "", string(body), time.Now())
	if err != nil {
		httpError(w, err)
		return
	}
	writeJSON(w, map[string]any{
		"submission_id": res.SubmissionID,
		"appended":      res.Appended,
		"exact_dups":    res.ExactDups,
		"needs_review":  res.NeedsReview,
		"ignored":       res.Ignored,
		"notes":         res.Notes,
	})
}

var homeTmpl = template.Must(template.New("home").
	Funcs(template.FuncMap{"flames": normalize.Flames}).
	Parse(`<!doctype html>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>our206 — upcoming shows</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 44rem; margin: 2rem auto; padding: 0 1rem; }
  h1 { letter-spacing: .04em; }
  .day { margin-top: 1.5rem; font-weight: 600; color: #555; }
  .show { padding: .4rem 0; border-bottom: 1px solid #eee; }
  .venue { color: #777; }
  a { color: inherit; }
</style>
<h1>our206</h1>
<p>{{len .Events}} upcoming shows</p>
{{$day := ""}}
{{range .Events}}
  {{if ne .DateKey $day}}<div class="day">{{.DateKey}}</div>{{$day = .DateKey}}{{end}}
  <div class="show">
    {{if .TicketURL}}<a href="{{.TicketURL}}">{{.Artist}}</a>{{else}}{{.Artist}}{{end}}
    <span class="venue">@ {{.VenueName}}</span>
    {{if .Score}}<span>{{flames .Score}}</span>{{end}}
  </div>
{{end}}
`))

func (s *Server) home(w http.ResponseWriter, _ *http.Request) {
	events, err := s.store.ListEvents(todayKey(), horizonKey())
	if err != nil {
		httpError(w, err)
		return
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := homeTmpl.Execute(w, map[string]any{"Events": events}); err != nil {
		log.Printf("render home: %v", err)
	}
}

func writeJSON(w http.ResponseWriter, v any) {
	w.Header().Set("Content-Type", "application/json")
	if err := json.NewEncoder(w).Encode(v); err != nil {
		log.Printf("write json: %v", err)
	}
}

func httpError(w http.ResponseWriter, err error) {
	log.Printf("http error: %v", err)
	http.Error(w, "internal error", http.StatusInternalServerError)
}

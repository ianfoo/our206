// Command our206d is the our206 platform service and admin CLI.
//
// Subcommands:
//
//	serve          run the HTTP server (also expires stale pending proposals)
//	import [file]  import contributor shorthand from a file or stdin
//	import-csv F   one-time migration from a spreadsheet CSV export
//	events         list published events
//	pending        list proposals awaiting review
//	sync           reconcile upcoming events to Google Calendar
//
// Environment:
//
//	OUR206_DB           SQLite path (default ./our206.db)
//	OUR206_ADDR         listen address for serve (default :8080)
//	OUR206_ADMIN_TOKEN  bearer token required by POST /api/import
//	OUR206_CALENDAR_ID  Google Calendar id for sync
//	GOOGLE_APPLICATION_CREDENTIALS  service-account key for sync
package main

import (
	"context"
	"flag"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"time"

	"github.com/ianfoo/our206/server/internal/gcal"
	"github.com/ianfoo/our206/server/internal/importer"
	"github.com/ianfoo/our206/server/internal/normalize"
	"github.com/ianfoo/our206/server/internal/store"
	"github.com/ianfoo/our206/server/internal/web"
)

func main() {
	log.SetFlags(0)
	if len(os.Args) < 2 {
		usage()
		os.Exit(2)
	}

	st, err := store.Open(envOr("OUR206_DB", "our206.db"))
	if err != nil {
		log.Fatalf("open database: %v", err)
	}
	defer st.Close()
	if err := st.Seed(); err != nil {
		log.Fatalf("seed venues: %v", err)
	}

	cmd, args := os.Args[1], os.Args[2:]
	switch cmd {
	case "serve":
		err = serve(st, args)
	case "import":
		err = importShorthand(st, args)
	case "import-csv":
		err = importCSV(st, args)
	case "events":
		err = listEvents(st, args)
	case "pending":
		err = listPending(st)
	case "sync":
		err = sync(st, args)
	default:
		usage()
		os.Exit(2)
	}
	if err != nil {
		log.Fatalf("%s: %v", cmd, err)
	}
}

func usage() {
	fmt.Fprintln(os.Stderr, "usage: our206d <serve|import|import-csv|events|pending|sync> [flags]")
}

func envOr(key, def string) string {
	if v := os.Getenv(key); v != "" {
		return v
	}
	return def
}

func serve(st *store.Store, args []string) error {
	fs := flag.NewFlagSet("serve", flag.ExitOnError)
	addr := fs.String("addr", envOr("OUR206_ADDR", ":8080"), "listen address")
	fs.Parse(args)

	if n, err := st.ExpirePastPending(time.Now().Format("2006-01-02")); err != nil {
		return err
	} else if n > 0 {
		log.Printf("expired %d stale pending proposal(s)", n)
	}

	handler := web.New(st, os.Getenv("OUR206_ADMIN_TOKEN"))
	log.Printf("our206d listening on %s", *addr)
	return http.ListenAndServe(*addr, handler)
}

func importShorthand(st *store.Store, args []string) error {
	fs := flag.NewFlagSet("import", flag.ExitOnError)
	submitter := fs.String("submitter", "", "who submitted this batch")
	fs.Parse(args)

	text, err := readInput(fs.Args())
	if err != nil {
		return err
	}
	res, err := importer.Shorthand(st, "paste", *submitter, text, time.Now())
	if err != nil {
		return err
	}
	report(res)
	return nil
}

func importCSV(st *store.Store, args []string) error {
	fs := flag.NewFlagSet("import-csv", flag.ExitOnError)
	fs.Parse(args)
	if fs.NArg() != 1 {
		return fmt.Errorf("usage: our206d import-csv <file.csv>")
	}
	f, err := os.Open(fs.Arg(0))
	if err != nil {
		return err
	}
	defer f.Close()
	res, err := importer.CSV(st, "csv", f)
	if err != nil {
		return err
	}
	report(res)
	return nil
}

func report(res importer.Result) {
	fmt.Println(res)
	for _, n := range res.Notes {
		fmt.Println("  -", n)
	}
}

func readInput(args []string) (string, error) {
	if len(args) == 1 && args[0] != "-" {
		b, err := os.ReadFile(args[0])
		return string(b), err
	}
	b, err := io.ReadAll(os.Stdin)
	return string(b), err
}

func listEvents(st *store.Store, args []string) error {
	fs := flag.NewFlagSet("events", flag.ExitOnError)
	from := fs.String("from", time.Now().Format("2006-01-02"), "start date (YYYY-MM-DD, empty for all)")
	to := fs.String("to", "", "end date (YYYY-MM-DD)")
	fs.Parse(args)

	events, err := st.ListEvents(*from, *to)
	if err != nil {
		return err
	}
	for _, e := range events {
		line := fmt.Sprintf("%s  %-40s @ %s", e.DateKey, e.Artist, e.VenueName)
		if e.Score > 0 {
			line += "  " + normalize.Flames(e.Score)
		}
		fmt.Println(line)
	}
	fmt.Printf("%d event(s)\n", len(events))
	return nil
}

func listPending(st *store.Store) error {
	pending, err := st.PendingProposals()
	if err != nil {
		return err
	}
	for _, p := range pending {
		fmt.Printf("#%d  %s  %s @ %s  [%s] %s\n",
			p.ID, p.DateKey, p.Artist, p.VenueRaw, p.Disposition, p.Note)
	}
	fmt.Printf("%d pending proposal(s)\n", len(pending))
	return nil
}

func sync(st *store.Store, args []string) error {
	fs := flag.NewFlagSet("sync", flag.ExitOnError)
	dryRun := fs.Bool("dry-run", false, "log actions without touching the calendar")
	fs.Parse(args)

	calendarID := os.Getenv("OUR206_CALENDAR_ID")
	if calendarID == "" {
		return fmt.Errorf("OUR206_CALENDAR_ID is not set")
	}

	now := time.Now()
	from := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.Local)
	to := from.AddDate(gcal.HorizonYears, 0, 0)

	events, err := st.ListEvents(from.Format("2006-01-02"), to.Format("2006-01-02"))
	if err != nil {
		return err
	}

	client, err := gcal.NewGoogleClient(context.Background(), calendarID)
	if err != nil {
		return err
	}
	sum, err := gcal.Reconcile(client, gcal.Desired(events), from, to, *dryRun)
	for _, l := range sum.Log {
		fmt.Println(l)
	}
	fmt.Println(sum)
	return err
}

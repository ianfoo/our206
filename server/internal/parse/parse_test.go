package parse

import (
	"testing"
	"time"
)

// ref is a fixed reference date for year inference: 2026-07-06.
var ref = time.Date(2026, 7, 6, 10, 0, 0, 0, time.Local)

func TestShorthandGrammar(t *testing.T) {
	tests := []struct {
		name string
		line string
		want Line
	}{
		{
			name: "canonical form with checkmarks",
			line: "3/20: Umphrey's McGee @ Showbox ✅✅✅✅",
			want: Line{DateKey: "2027-03-20", Artist: "Umphrey's McGee", Venue: "Showbox", Score: 4},
		},
		{
			name: "no score",
			line: "3/21: Band of Horses @ Showbox",
			want: Line{DateKey: "2027-03-21", Artist: "Band of Horses", Venue: "Showbox", Score: 0},
		},
		{
			name: "missing colon",
			line: "8/24 Machine Girl @ SoDo Showbox ✅✅",
			want: Line{DateKey: "2026-08-24", Artist: "Machine Girl", Venue: "SoDo Showbox", Score: 2},
		},
		{
			name: "bang marks",
			line: "9/1: The Beths @ Neumos !!!",
			want: Line{DateKey: "2026-09-01", Artist: "The Beths", Venue: "Neumos", Score: 3},
		},
		{
			name: "score capped at four",
			line: "9/2: Hype Band @ Neumos ✅✅✅✅✅✅",
			want: Line{DateKey: "2026-09-02", Artist: "Hype Band", Venue: "Neumos", Score: 4},
		},
		{
			name: "trailing commentary stripped",
			line: "10/5: Slow Dive @ Paramount - can't wait for this one",
			want: Line{DateKey: "2026-10-05", Artist: "Slow Dive", Venue: "Paramount", Score: 0},
		},
		{
			name: "trailing parenthetical stripped from venue",
			line: "10/6: Sharp Pins @ Vera Project (all ages)",
			want: Line{DateKey: "2026-10-06", Artist: "Sharp Pins", Venue: "Vera Project", Score: 0},
		},
		{
			name: "artist containing @ uses last @ as separator",
			line: "10/7: DJ A@B @ Chop Suey",
			want: Line{DateKey: "2026-10-07", Artist: "DJ A@B", Venue: "Chop Suey", Score: 0},
		},
		{
			name: "explicit year honored",
			line: "1/15/2028: Future Band @ Tractor",
			want: Line{DateKey: "2028-01-15", Artist: "Future Band", Venue: "Tractor", Score: 0},
		},
		{
			name: "past month rolls to next year",
			line: "1/15: Winter Band @ Tractor",
			want: Line{DateKey: "2027-01-15", Artist: "Winter Band", Venue: "Tractor", Score: 0},
		},
		{
			name: "today stays this year",
			line: "7/6: Tonight Band @ Neumos",
			want: Line{DateKey: "2026-07-06", Artist: "Tonight Band", Venue: "Neumos", Score: 0},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, ok := Shorthand(tt.line, ref)
			if !ok {
				t.Fatalf("Shorthand(%q) not parsed", tt.line)
			}
			if got != tt.want {
				t.Errorf("Shorthand(%q) = %+v, want %+v", tt.line, got, tt.want)
			}
		})
	}
}

func TestShorthandRejects(t *testing.T) {
	lines := []string{
		"",
		"random chatter about a band",
		"3/20: No venue here",            // no @
		"Umphrey's McGee @ Showbox",      // no date
		"2/30: Impossible Date @ Neumos", // invalid calendar date
		"3/20: @ Showbox",                // empty artist
		"3/20: Some Band @",              // empty venue
	}
	for _, line := range lines {
		if _, ok := Shorthand(line, ref); ok {
			t.Errorf("Shorthand(%q) parsed, want rejection", line)
		}
	}
}

func TestDateKey(t *testing.T) {
	tests := map[string]string{
		"20-Mar-2026": "2026-03-20",
		"3/20/2026":   "2026-03-20",
		"2026-03-20":  "2026-03-20",
		"":            "",
		"Date":        "", // header cell
	}
	for in, want := range tests {
		if got := DateKey(in); got != want {
			t.Errorf("DateKey(%q) = %q, want %q", in, got, want)
		}
	}
}

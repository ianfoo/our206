// Package parse implements the contributor shorthand grammar:
//
//	3/20: Umphrey's McGee @ Showbox ✅✅✅✅
//	3/21 Band of Horses @ Showbox
//	3/24: Machine Girl @ SoDo Showbox (all ages) - should be wild
//
// It tolerates a missing colon, arbitrary capitalization, "- commentary"
// suffixes, trailing parenthetical notes, and both "!" and "✅" as
// excitement marks. Ports parseIncomingLine_ from the Apps Script
// prototype, with one deliberate improvement: instead of a hardcoded year,
// the year is inferred as the next occurrence of the month/day relative to
// a reference date, and an explicit M/D/YYYY year is honored when present.
package parse

import (
	"fmt"
	"regexp"
	"strings"
	"time"
)

// Line is one successfully parsed shorthand line.
type Line struct {
	DateKey string // YYYY-MM-DD
	Artist  string // raw artist text (not yet normalized)
	Venue   string // raw venue text (not yet canonicalized)
	Score   int    // 0..4 excitement marks
}

var (
	datePrefixRe    = regexp.MustCompile(`^\s*(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?:?\s*`)
	marksRe         = regexp.MustCompile(`[✅!]+`)
	commentaryRe    = regexp.MustCompile(`\s+-\s+.*$`)
	trailingParenRe = regexp.MustCompile(`\s*\([^)]*\)\s*$`)
)

// Shorthand parses one line relative to ref (used for year inference).
// It returns ok=false for lines that don't match the grammar; callers
// decide how to record those.
func Shorthand(line string, ref time.Time) (Line, bool) {
	m := datePrefixRe.FindStringSubmatch(line)
	if m == nil || !strings.Contains(line, "@") {
		return Line{}, false
	}

	month, day := atoi(m[1]), atoi(m[2])
	if month < 1 || month > 12 || day < 1 || day > 31 {
		return Line{}, false
	}
	year := inferYear(month, day, m[3], ref)

	rest := strings.TrimSpace(datePrefixRe.ReplaceAllString(line, ""))
	score := countMarks(rest)
	rest = strings.TrimSpace(marksRe.ReplaceAllString(rest, ""))
	rest = strings.TrimSpace(commentaryRe.ReplaceAllString(rest, ""))

	atIdx := strings.LastIndex(rest, "@")
	if atIdx == -1 {
		return Line{}, false
	}
	artist := strings.TrimSpace(rest[:atIdx])
	venue := strings.TrimSpace(rest[atIdx+1:])
	venue = strings.TrimSpace(trailingParenRe.ReplaceAllString(venue, ""))
	if artist == "" || venue == "" {
		return Line{}, false
	}

	// Validate the calendar date (rejects 2/30 etc.).
	d := time.Date(year, time.Month(month), day, 12, 0, 0, 0, time.Local)
	if int(d.Month()) != month || d.Day() != day {
		return Line{}, false
	}

	return Line{
		DateKey: fmt.Sprintf("%04d-%02d-%02d", year, month, day),
		Artist:  artist,
		Venue:   venue,
		Score:   score,
	}, true
}

// inferYear picks the year for a M/D shorthand date. An explicit year wins
// (2-digit years are 2000-based). Otherwise the date is assumed to be the
// next occurrence on or after ref's date.
func inferYear(month, day int, explicit string, ref time.Time) int {
	if explicit != "" {
		y := atoi(explicit)
		if y < 100 {
			y += 2000
		}
		return y
	}
	y := ref.Year()
	candidate := time.Date(y, time.Month(month), day, 12, 0, 0, 0, time.Local)
	today := time.Date(ref.Year(), ref.Month(), ref.Day(), 0, 0, 0, 0, time.Local)
	if candidate.Before(today) {
		return y + 1
	}
	return y
}

func countMarks(s string) int {
	n := 0
	for _, r := range s {
		if r == '✅' || r == '!' {
			n++
		}
	}
	if n > 4 {
		n = 4
	}
	return n
}

func atoi(s string) int {
	n := 0
	for _, r := range s {
		n = n*10 + int(r-'0')
	}
	return n
}

var (
	ddMMMyyyyRe = regexp.MustCompile(`^(\d{1,2})-([A-Za-z]{3})-(\d{4})$`)
	mdyRe       = regexp.MustCompile(`^(\d{1,2})/(\d{1,2})/(\d{4})$`)
	isoRe       = regexp.MustCompile(`^(\d{4})-(\d{2})-(\d{2})$`)
)

// DateKey normalizes the date formats that appear in the spreadsheet
// (DD-MMM-YYYY, M/D/YYYY, YYYY-MM-DD) to YYYY-MM-DD. Ports
// sheetDateToKey_. Returns "" for unrecognized input.
func DateKey(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return ""
	}
	if m := ddMMMyyyyRe.FindStringSubmatch(s); m != nil {
		t, err := time.ParseInLocation("2-Jan-2006", fmt.Sprintf("%d-%s-%s", atoi(m[1]), m[2], m[3]), time.Local)
		if err != nil {
			return ""
		}
		return t.Format("2006-01-02")
	}
	if m := mdyRe.FindStringSubmatch(s); m != nil {
		return fmt.Sprintf("%04d-%02d-%02d", atoi(m[3]), atoi(m[1]), atoi(m[2]))
	}
	if isoRe.MatchString(s) {
		return s
	}
	return ""
}

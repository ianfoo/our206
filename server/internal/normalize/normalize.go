// Package normalize ports the artist/score normalization behavior of the
// Apps Script prototype (calendar-sync/incoming-raw.js). Behavior parity
// with the spreadsheet system is intentional: these functions feed dedup
// keys and calendar UIDs, so changes here change event identity.
package normalize

import (
	"regexp"
	"strings"
)

// knownArtistFixes ports fixKnownArtistNames_.
var knownArtistFixes = map[string]string{
	"Royksopp (DJ Set)": "Röyksopp (DJ Set)",
	"Royksopp":          "Röyksopp",
	"Devin The Due":     "Devin The Dude",
}

var (
	titleWordRe    = regexp.MustCompile(`\b[a-z][a-z']*`)
	nonLetterRe    = regexp.MustCompile(`[^A-Za-z]`)
	ampRe          = regexp.MustCompile(`&`)
	xJoinRe        = regexp.MustCompile(`\bx\b`)
	nonWordSpaceRe = regexp.MustCompile(`[^\w\s]`)
	spacesRe       = regexp.MustCompile(`\s+`)
	andFriendsRe   = regexp.MustCompile(`\band friends\b`)
	andSplitRe     = regexp.MustCompile(`\s+and\s+`)
)

// DeShoutifyArtist converts ALL-CAPS artist names to title case and applies
// known typo corrections. Mixed-case names pass through untouched (aside
// from typo fixes). Ports deShoutifyArtist_.
func DeShoutifyArtist(artist string) string {
	s := strings.TrimSpace(artist)
	letters := nonLetterRe.ReplaceAllString(s, "")
	if letters == "" {
		return s
	}
	if letters == strings.ToUpper(letters) {
		return titleCase(strings.ToLower(s))
	}
	if fixed, ok := knownArtistFixes[s]; ok {
		return fixed
	}
	return s
}

// titleCase ports toTitleCase_: uppercase the first letter of each
// lowercase word, preserving apostrophes within words.
func titleCase(s string) string {
	return titleWordRe.ReplaceAllStringFunc(s, func(w string) string {
		return strings.ToUpper(w[:1]) + w[1:]
	})
}

// ArtistCompareKey produces the canonical comparison form of an artist name
// used for exact duplicate detection and UID construction. Ports
// normalizeArtistForCompare_.
func ArtistCompareKey(artist string) string {
	s := strings.ToLower(artist)
	s = ampRe.ReplaceAllString(s, " and ")
	s = xJoinRe.ReplaceAllString(s, " and ")
	s = nonWordSpaceRe.ReplaceAllString(s, "")
	s = spacesRe.ReplaceAllString(s, " ")
	return strings.TrimSpace(s)
}

// PrimaryArtistKey reduces a possibly-collaborative billing to its primary
// act ("Tom Hamilton x Swindler" -> "tom hamilton") for fuzzy duplicate
// detection. Ports primaryArtistKey_.
func PrimaryArtistKey(artist string) string {
	s := ArtistCompareKey(artist)
	s = strings.TrimSpace(andFriendsRe.ReplaceAllString(s, ""))
	parts := andSplitRe.Split(s, -1)
	return strings.TrimSpace(parts[0])
}

// MaxScore is the cap on the excitement score.
const MaxScore = 4

// ScoreFromMarks counts excitement marks (✅, !, 🔥) in a string and caps
// the result at MaxScore. Used both for shorthand lines and for importing
// spreadsheet rating cells.
func ScoreFromMarks(s string) int {
	n := 0
	for _, r := range s {
		switch r {
		case '✅', '!', '🔥':
			n++
		}
	}
	if n > MaxScore {
		return MaxScore
	}
	return n
}

// Flames renders a score as the 🔥 display string used in calendar
// descriptions and the website.
func Flames(score int) string {
	if score < 0 {
		score = 0
	}
	if score > MaxScore {
		score = MaxScore
	}
	return strings.Repeat("🔥", score)
}

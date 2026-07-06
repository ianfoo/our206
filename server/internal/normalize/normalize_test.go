package normalize

import "testing"

func TestDeShoutifyArtist(t *testing.T) {
	tests := map[string]string{
		"SOME BAND":         "Some Band",
		"UMPHREY'S MCGEE":   "Umphrey's Mcgee", // title-case per word, apostrophe kept in word
		"Band of Horses":    "Band of Horses",  // mixed case untouched
		"Royksopp":          "Röyksopp",        // known fix
		"Royksopp (DJ Set)": "Röyksopp (DJ Set)",
		"Devin The Due":     "Devin The Dude",
		"MGMT":              "Mgmt", // all-caps names do get folded, as in the prototype
		"":                  "",
		"123":               "123", // no letters: unchanged
	}
	for in, want := range tests {
		if got := DeShoutifyArtist(in); got != want {
			t.Errorf("DeShoutifyArtist(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestArtistCompareKey(t *testing.T) {
	tests := map[string]string{
		"Tom Hamilton x Swindler": "tom hamilton and swindler",
		"Scott Law & Friends":     "scott law and friends",
		"Umphrey's McGee":         "umphreys mcgee",
		"  Big   Gap  ":           "big gap",
		"X Ambassadors":           "and ambassadors", // known quirk of \bx\b, preserved from prototype
	}
	for in, want := range tests {
		if got := ArtistCompareKey(in); got != want {
			t.Errorf("ArtistCompareKey(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestPrimaryArtistKey(t *testing.T) {
	tests := map[string]string{
		"Tom Hamilton x Swindler": "tom hamilton",
		"Scott Law & Friends":     "scott law",
		"Scott Law x Jay Cobb":    "scott law",
		"Band of Horses":          "band of horses",
	}
	for in, want := range tests {
		if got := PrimaryArtistKey(in); got != want {
			t.Errorf("PrimaryArtistKey(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestScores(t *testing.T) {
	if got := ScoreFromMarks("✅✅✅✅✅"); got != 4 {
		t.Errorf("ScoreFromMarks cap = %d, want 4", got)
	}
	if got := ScoreFromMarks("!✅"); got != 2 {
		t.Errorf("ScoreFromMarks mixed = %d, want 2", got)
	}
	if got := ScoreFromMarks("🔥🔥"); got != 2 {
		t.Errorf("ScoreFromMarks flames = %d, want 2", got)
	}
	if got := Flames(3); got != "🔥🔥🔥" {
		t.Errorf("Flames(3) = %q", got)
	}
	if got := Flames(0); got != "" {
		t.Errorf("Flames(0) = %q, want empty", got)
	}
}

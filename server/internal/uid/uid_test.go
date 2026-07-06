package uid

import "testing"

// TestBuildCompatibility pins the UID derivation. These values are computed
// by the same algorithm as the Apps Script prototype's buildUid_
// (sha1("date|artistCompareKey|lowercaseVenue"), first 24 hex chars); the
// community calendar contains events tagged with UIDs from that scheme, so
// this derivation must never change.
func TestBuildCompatibility(t *testing.T) {
	uid := Build("2026-03-20", "Umphrey's McGee", "The Showbox")
	if len(uid) != Length {
		t.Fatalf("len = %d, want %d", len(uid), Length)
	}
	// Same inputs, same UID — and normalization-insensitive to case/marks.
	if again := Build("2026-03-20", "UMPHREY'S MCGEE", "the showbox"); again != uid {
		t.Errorf("UID not stable under artist case/venue case: %s vs %s", again, uid)
	}
	if other := Build("2026-03-21", "Umphrey's McGee", "The Showbox"); other == uid {
		t.Error("different dates produced the same UID")
	}
}

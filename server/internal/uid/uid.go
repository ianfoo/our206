// Package uid derives the deterministic event identity used to reconcile
// canonical events with Google Calendar. It must stay byte-compatible with
// buildUid_ in the Apps Script prototype: the community's calendar already
// contains events tagged with these UIDs, and changing the derivation would
// orphan them all.
package uid

import (
	"crypto/sha1"
	"encoding/hex"
	"strings"

	"github.com/ianfoo/our206/server/internal/normalize"
)

// Length is the number of hex characters kept from the SHA-1 digest.
const Length = 24

// Build derives the UID from the event date (YYYY-MM-DD), the artist as
// displayed, and the canonical venue name.
func Build(dateKey, artist, canonicalVenue string) string {
	seed := dateKey + "|" + normalize.ArtistCompareKey(artist) + "|" + strings.ToLower(canonicalVenue)
	sum := sha1.Sum([]byte(seed))
	return hex.EncodeToString(sum[:])[:Length]
}

package store

// Seed data carried over from the Apps Script prototype
// (calendar-sync/config.js): canonical venues with known addresses, and the
// alias map used for normalization. Seeding is idempotent; a Venue Map
// sheet export can layer on top of this via the importer later.

type seedVenue struct {
	name    string
	address string
}

var seedVenues = []seedVenue{
	{"Airport Tavern", ""},
	{"Baba Yaga", ""},
	{"Barboza", ""},
	{"Belltown Yacht Club", ""},
	{"Chop Suey", "1325 E Madison St, Seattle, WA 98122"},
	{"Clock-Out Lounge", "4864 Beacon Ave S, Seattle, WA 98108"},
	{"Darrell’s Tavern", ""},
	{"Edmonds Center for the Arts", "410 4th Ave N, Edmonds, WA 98020"},
	{"El Corazon", ""},
	{"Hidden Hall", "400 N 35th St, Seattle, WA 98103"},
	{"Massive", ""},
	{"Moore Theatre", "1932 2nd Ave, Seattle, WA 98101"},
	{"Nectar Lounge", "412 N 36th St, Seattle, WA 98103"},
	{"Neptune Theatre", "1303 NE 45th St, Seattle, WA 98105"},
	{"Neumos", "925 E Pike St, Seattle, WA 98122"},
	{"Paramount Theatre", "911 Pine St, Seattle, WA 98101"},
	{"Pony", "1221 E Madison St, Seattle, WA 98122"},
	{"Q Nightclub", "1426 Broadway, Seattle, WA 98122"},
	{"Real Art Tacoma", ""},
	{"Showbox SoDo", "1700 1st Ave S, Seattle, WA 98134"},
	{"Stumpfest PDX", ""},
	{"Substation Seattle", "645 NW 45th St, Seattle, WA 98107"},
	{"T-Mobile Park", ""},
	{"The Chapel", "4649 Sunnyside Ave N, Seattle, WA 98103"},
	{"The Crocodile", "2505 1st Ave, Seattle, WA 98121"},
	{"The Showbox", "1426 1st Ave, Seattle, WA 98101"},
	{"The Sunset Tavern", ""},
	{"The Triple Door", ""},
	{"Town Hall Seattle", "1119 8th Ave, Seattle, WA 98101"},
	{"Tractor Tavern", "5213 Ballard Ave NW, Seattle, WA 98107"},
	{"WaMu Theater", ""},
	{"Wheelie Pop Brewing", ""},
}

// seedAliases maps lowercase raw names to canonical venue names.
var seedAliases = map[string]string{
	"sodo showbox":        "Showbox SoDo",
	"showbox sodo":        "Showbox SoDo",
	"showbox":             "The Showbox",
	"croc":                "The Crocodile",
	"neptune":             "Neptune Theatre",
	"paramount":           "Paramount Theatre",
	"substation":          "Substation Seattle",
	"nectar":              "Nectar Lounge",
	"tractor":             "Tractor Tavern",
	"chop":                "Chop Suey",
	"clock-out":           "Clock-Out Lounge",
	"wheelie pop":         "Wheelie Pop Brewing",
	"sunset":              "The Sunset Tavern",
	"the moore":           "Moore Theatre",
	"q":                   "Q Nightclub",
	"edmonds arts center": "Edmonds Center for the Arts",
	"wamu":                "WaMu Theater",
	"darrel’s tavern":     "Darrell’s Tavern",
	"darrel's tavern":     "Darrell’s Tavern",
}

// Seed inserts the baseline venues and aliases, skipping anything already
// present. Safe to run on every startup.
func (s *Store) Seed() error {
	for _, v := range seedVenues {
		if _, err := s.db.Exec(`INSERT OR IGNORE INTO venues (name, address) VALUES (?, ?)`,
			v.name, v.address); err != nil {
			return err
		}
		// Backfill addresses onto venues created earlier without one.
		if v.address != "" {
			if _, err := s.db.Exec(
				`UPDATE venues SET address = ? WHERE lower(name) = lower(?) AND address = ''`,
				v.address, v.name); err != nil {
				return err
			}
		}
	}
	for alias, canonical := range seedAliases {
		var venueID int64
		if err := s.db.QueryRow(`SELECT id FROM venues WHERE lower(name) = lower(?)`, canonical).
			Scan(&venueID); err != nil {
			return err
		}
		if err := s.AddAlias(alias, venueID); err != nil {
			return err
		}
	}
	return nil
}

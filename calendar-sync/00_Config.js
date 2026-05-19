/**
 * Our206 — combined sheet/calendar automation
 *
 * Includes:
 * - Debounced future-event calendar sync
 * - Dry-run sync
 * - Daily maintenance: move past events -> Past Concerts, sort, compact, sync
 * - Incoming Raw processor: parse messy source lines, normalize, dedupe, append to Concerts
 * - Venue normalization via reference tab "Venue Map" if present, with baked-in fallback map
 * - Automatic copying of Google Form input to main concert tab
 * - Logging + custom menu
 *
 * Tabs used:
 * - Concerts
 * - Past Concerts
 * - Incoming Raw   (optional, for ingestion)
 * - Venue Map      (optional, for venue KV normalization)
 *
 * Venue Map format:
 *   A: Raw Venue
 *   B: Normalized Venue
 * Starting at row 2. Header row 1 can be anything.
 */

const SHEET_KEYS = {
  CONCERTS: "concerts",
  PAST_CONCERTS: "pastConcerts",
  INCOMING_RAW: "incomingRaw",
  VENUE_MAP: "venueMap"
};

const CFG = {
  calendarId: "our206wa@gmail.com",
  sheets: {
    [SHEET_KEYS.CONCERTS]: {
      name: "Concerts",
      headerRow: 3
    },
    [SHEET_KEYS.PAST_CONCERTS]: {
      name: "Past Concerts",
      headerRow: 3
    },
    [SHEET_KEYS.INCOMING_RAW]: {
      name: "Incoming Raw",
      headerRow: 1
    },
    [SHEET_KEYS.VENUE_MAP]: {
      name: "Venue Map",
      headerRow: 1
    }
  },
  headerRowFallback: 3,
  debounceMinutes: 10,
  debounceGuardMinutes: 8,
  horizonYears: 2,
  uidHeader: "UID",
  uidMarkerPrefix: "[our206_uid]:",
  headerMatchers: {
    date: "date",
    artist: "artist",
    venue: "venue",
    rating: "skoi",
    notes: "notes",
    ticket: "ticket",
    addedOn: "added"
  },
  keepLastColumnHeader: "cap"
};

// Fallback venue normalization map.
// If Venue Map tab exists, its values take precedence.
const VENUE_ADDRESS = {
  "Chop Suey": "1325 E Madison St, Seattle, WA 98122",
  "Clock-Out Lounge": "4864 Beacon Ave S, Seattle, WA 98108",
  "Edmonds Center for the Arts": "410 4th Ave N, Edmonds, WA 98020",
  "Hidden Hall": "400 N 35th St, Seattle, WA 98103",
  "Moore Theatre": "1932 2nd Ave, Seattle, WA 98101",
  "Nectar Lounge": "412 N 36th St, Seattle, WA 98103",
  "Neptune Theatre": "1303 NE 45th St, Seattle, WA 98105",
  "Neumos": "925 E Pike St, Seattle, WA 98122",
  "Paramount Theatre": "911 Pine St, Seattle, WA 98101",
  "Pony": "1221 E Madison St, Seattle, WA 98122",
  "Q Nightclub": "1426 Broadway, Seattle, WA 98122",
  "Showbox SoDo": "1700 1st Ave S, Seattle, WA 98134",
  "Substation Seattle": "645 NW 45th St, Seattle, WA 98107",
  "The Chapel": "4649 Sunnyside Ave N, Seattle, WA 98103",
  "The Crocodile": "2505 1st Ave, Seattle, WA 98121",
  "The Showbox": "1426 1st Ave, Seattle, WA 98101",
  "Town Hall Seattle": "1119 8th Ave, Seattle, WA 98101",
  "Tractor Tavern": "5213 Ballard Ave NW, Seattle, WA 98107",
  "Wheelie Pop Brewing": "",
  "Darrell’s Tavern": "",
  "The Triple Door": "",
  "El Corazon": "",
  "Belltown Yacht Club": "",
  "Airport Tavern": "",
  "Massive": "",
  "Real Art Tacoma": "",
  "WaMu Theater": "",
  "T-Mobile Park": "",
  "Barboza": "",
  "Stumpfest PDX": ""
};

const FALLBACK_VENUE_NORMALIZATION = {
  "sodo showbox": "Showbox SoDo",
  "showbox sodo": "Showbox SoDo",
  "showbox": "The Showbox",
  "croc": "The Crocodile",
  "the crocodile": "The Crocodile",
  "neptune": "Neptune Theatre",
  "paramount": "Paramount Theatre",
  "substation": "Substation Seattle",
  "nectar": "Nectar Lounge",
  "tractor": "Tractor Tavern",
  "chop": "Chop Suey",
  "clock-out": "Clock-Out Lounge",
  "clock-out lounge": "Clock-Out Lounge",
  "wheelie pop": "Wheelie Pop Brewing",
  "sunset": "The Sunset Tavern",
  "the moore": "Moore Theatre",
  "q": "Q Nightclub",
  "edmonds arts center": "Edmonds Center for the Arts",
  "edmonds center for the arts": "Edmonds Center for the Arts",
  "wamu": "WaMu Theater",
  "darrel’s tavern": "Darrell’s Tavern",
  "darrel's tavern": "Darrell’s Tavern",
  "massive": "Massive",
  "real art tacoma": "Real Art Tacoma",
  "airport tavern": "Airport Tavern",
  "t-mobile park": "T-Mobile Park",
  "barboza": "Barboza",
  "hidden hall": "Hidden Hall",
  "the triple door": "The Triple Door",
  "el corazon": "El Corazon",
  "belltown yacht club": "Belltown Yacht Club",
  "stumpfest pdx": "Stumpfest PDX",
  "neumos": "Neumos",
  "paramount theatre": "Paramount Theatre",
  "neptune theatre": "Neptune Theatre",
  "tractor tavern": "Tractor Tavern",
  "nectar lounge": "Nectar Lounge",
  "substation seattle": "Substation Seattle",
  "the showbox": "The Showbox",
  "showbox sodo": "Showbox SoDo"
};

// ---------------------- MENU ----------------------

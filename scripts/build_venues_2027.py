import json

OUT = "/tmp/claude-0/-home-claude/5f5137a5-0701-5ed4-a98c-f2f53eaefc9b/scratchpad/bbg/research/venues_2027.json"

def sc(crowd, venue, fighters, competition, distance):
    return {"crowd": crowd, "venue": venue, "fighters": fighters, "competition": competition,
            "distance": distance, "overall": round((crowd + venue + fighters + competition + distance) / 5, 1)}

def v(name, typ, address, cap, capnote, winter, wnote, combat, rate, contact, src, s):
    return {"name": name, "type": typ, "address": address, "capacity_event": cap, "capacity_source_note": capnote,
            "winter_ok": winter, "winter_note": wnote, "hosted_combat": combat, "rental_rate": rate,
            "booking_contact": contact, "source_url": src, "scores": s}

TCC_SRC = "https://www.tributecommunitiescentre.com/host-an-event"

towns = [
 {
  "town": "Peterborough",
  "region": "City of Peterborough (Kawarthas)",
  "population": "83,651 (City of Peterborough, 2021 census)",
  "drive_from_bancroft": "About 1 h 37 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("Peterborough Memorial Centre (PMC)", "arena", "151 Lansdowne Street West, Peterborough, ON K9J 1Y4",
     "Concerts: 'customizable capacity ranging from 1722 to 5501 including floor capacity'; ice events 4,050 max. No boxing/MMA configuration published.",
     "PMC Tech Pack 2024 (venuecoalition.com). Wikipedia lists hockey seating 3,729 (4,329 with standing room). The tech pack lists 'Full event subfloor decking' among event accessories.",
     "yes", "Holds dry-floor events during Petes (OHL) season: Dean Brody & The Reklaws concert Sat Nov 21, 2026, and OFSAA wrestling Mar 2-4, 2026. Dates must fit around Petes home games.",
     "No 2024-2026 boxing/MMA found. Earlier pro 'Homecoming' cards (e.g., Feb 9, 2019, where Cody Purtell was billed; confirm with Cody). OFSAA high-school wrestling, Mar 2-4, 2026.",
     "Not published - request a quote",
     "Jose Garcia-Perez, Event Services Supervisor, 705-743-3561 ext. 2504, jgarcia-perez@peterborough.ca",
     "https://venuecoalition.com/wp-content/uploads/2025/02/Tech-Pack-2024.pdf",
     sc(5, 4, 5, 4, 4)),
   v("Healthy Planet Arena (formerly Evinrude Centre)", "arena", "911 Monaghan Rd, Peterborough, ON K9J 5K6",
     "Not published", "Two pads plus a 5,000 sq ft multipurpose room (City page, from earlier CCCL research).",
     "no", "Floor season only: the 2026 spring/summer floor allocation ran Apr 12 to Aug 1, 2026.",
     "None found 2024-2026 (a kickboxing event was held here in 2011 per ptbocanada.com).",
     "Special events on a floor pad: $3,700.75 per day per pad (City of Peterborough)",
     "Facility Booking 705-742-7777 ext. 2577, facilitybooking@peterborough.ca",
     "https://www.peterborough.ca/arenas",
     sc(4, 3, 5, 4, 4))
  ],
  "gyms_nearby": [
   {"name": "Kawartha Combat (BJJ, Muay Thai, kickboxing), 910 High St #15", "source_url": "https://letsroll.co/gym/kawartha-combat-peterborough-on"},
   {"name": "10th Planet Jiu Jitsu Peterborough", "source_url": "https://www.10thplanetjj.com/?p=418"},
   {"name": "Corcoran's Boxing Club and Eastside Boxing (Boxing Ontario clubs)", "source_url": "https://boxingontario.com/city/peterborough"},
   {"name": "BEL Boxing Club (1625 Chemong Rd, Selwyn Twp)", "source_url": "https://kawarthanow.com/venue/1500/bel-boxing-club"},
   {"name": "Big Country - Lindsay (full MMA gym, ~45-50 min, estimate)", "source_url": "https://bigcountrylindsay.gymdesk.com/page/home"}
  ],
  "combat_events": [
   {"name": "OFSAA Wrestling Championship (high-school wrestling, PMC)", "date": "2026-03-02 to 2026-03-04", "venue": "Peterborough Memorial Centre", "source_url": "https://www.ptbotoday.ca/?p=94603"},
   {"name": "No 2024-2026 MMA or pro boxing card found in Peterborough", "date": "", "venue": "", "source_url": ""}
  ],
  "clashes": [
   {"name": "Peterborough Petes (OHL) home games at PMC", "typical_dates": "Late Sept to March (plus playoffs)", "source_url": "https://en.wikipedia.org/wiki/Peterborough_Memorial_Centre"},
   {"name": "Dean Brody & The Reklaws at PMC", "typical_dates": "Sat Nov 21, 2026 (one-off; shows PMC books November Saturdays)", "source_url": "https://kawarthanow.com/event/74476/dean-brody-and-the-reklaws-hometown-heroes-tour/"}
  ],
  "scores": sc(5, 4, 5, 4, 4),
  "verdict": "Best 2027 bet for a 1,000+ winter show: closest city arena that holds events in Nov-Feb, smallest published setup 1,722, deepest local fighter pool, and no recent MMA competition. Cost is the unknown."
 },
 {
  "town": "Bancroft",
  "region": "North Hastings (Hastings County)",
  "population": "4,065 (Town of Bancroft, 2021 census); the Town says the regional population is 'more than 40 000'",
  "drive_from_bancroft": "0 (home base)",
  "over_3h": False,
  "venues": [
   v("North Hastings Community Centre (NHCC) arena", "arena", "103 Newkirk Blvd, Bancroft, ON K0L 1C0",
     "Not published", "Single pad, opened 1972. The ~$2.4M retrofit (new ice plant, accessible washroom, extra change rooms) was essentially done in Sept 2026.",
     "no", "Ice runs from about the end of August to mid-April (arena-guide.com). No sign the Town covers the ice for winter events. Dry-floor window is about April to July (estimate). The Gemboree uses NHCC over the Civic Holiday weekend.",
     "No MMA or boxing found. Smash Wrestling (pro wrestling) ran here May 15, 2015.",
     "Not published. The Town has said it wants year-round rentals; the arena ran a $207,350 deficit in 2023.",
     "613-332-2062; communitycentre@bancroft.ca (2023)",
     "https://bancroft.hosted.civiclive.com/living/recreation___leisure/north_hastings_community_centre",
     sc(2, 3, 4, 5, 5))
  ],
  "gyms_nearby": [
   {"name": "Bancroft Boxing & Grappling (Cody Purtell's gym, 17 Snow Road)", "source_url": "https://thehighlander.ca/2026/03/12/fight-night-coming-to-haliburton-county/"}
  ],
  "combat_events": [
   {"name": "None found 2024-2026. Show 1 was planned for Bancroft but moved to Haliburton because of the NHCC renovation.", "date": "", "venue": "", "source_url": "https://thehighlander.ca/2026/03/12/fight-night-coming-to-haliburton-county/"}
  ],
  "clashes": [
   {"name": "Rockhound Gemboree (fills NHCC)", "typical_dates": "Civic Holiday long weekend; 2026 opened Thu Jul 30. In 2027 the Civic Holiday is Mon Aug 2 (calendar).", "source_url": "https://bancroft.hosted.civiclive.com/living/recreation___leisure/north_hastings_community_centre"}
  ],
  "scores": sc(2, 3, 4, 5, 5),
  "verdict": "Home show for spring or early summer only (no winter option). 1,000+ is unproven: show 1 sold ~400 tickets. Get the NHCC floor capacity before booking."
 },
 {
  "town": "Haliburton",
  "region": "Haliburton County (Dysart et al)",
  "population": "7,182 (Dysart et al, 2021 census); Haliburton County 20,571 (2021)",
  "drive_from_bancroft": "About 54 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("A.J. LaRue Arena & Community Centre", "arena", "728 Mountain Street, Haliburton, ON K0M 1S0",
     "Not published. Show 1 (May 2, 2026) was planned for 800-1,000 (Highlander). Cody's actual figures: about 400 tickets sold and 400-600 people in the building.",
     "Single pad, no published seating (arena-guide.com). Expected attendance from The Highlander, Mar 12, 2026.",
     "no", "Ice season fall to spring; no sign of covering the ice for winter events. Spring dry floor proven (May 2). Summer ice may be sold through Labour Day, so confirm the 2027 ice-out and ice-in dates.",
     "CCCL 1 - Cottage Country Combat League amateur MMA, May 2, 2026",
     "By-Law 2023-39: ice surface $1,030.00/day for profit events ($463.50 non-profit); auditorium $42.45/hr (max $265.25/day); staff set-up $30/hr. These are 2023-24 rates.",
     "705-457-1740 ext. 635 (Municipality of Dysart et al)",
     "https://www.dysartetal.ca/en/explore-and-play/resources/Arena/2023-39%20Parks%20and%20Recreation%20Fees%20and%20Charges.pdf",
     sc(2, 3, 3, 5, 5))
  ],
  "gyms_nearby": [
   {"name": "The Chop Shop MMA (BJJ + kickboxing, opened May 30, 2026, 175 Industrial Park Rd)", "source_url": "https://thehighlander.ca/2026/05/21/bringing-hi-yah-to-haliburton/"},
   {"name": "Bancroft Boxing & Grappling (~54 min)", "source_url": "https://thehighlander.ca/2026/03/12/fight-night-coming-to-haliburton-county/"}
  ],
  "combat_events": [
   {"name": "CCCL 1 (amateur MMA, OAC-sanctioned)", "date": "2026-05-02", "venue": "A.J. LaRue Arena", "source_url": "https://thehighlander.ca/2026/03/12/fight-night-coming-to-haliburton-county/"}
  ],
  "clashes": [
   {"name": "Haliburton County Fair (Minden fairgrounds, 20 min away)", "typical_dates": "Mid-June; Sat June 13, 2026", "source_url": "https://www.mindenhills.ca/explore-discover/rent-a-facility/arena-and-ice-rentals/"}
  ],
  "scores": sc(2, 3, 3, 5, 5),
  "verdict": "Proven and cheap, but the crowd was about 400-600, half of the 1,000 target. Keep it as the yearly spring anniversary show only if Cody accepts a smaller house there."
 },
 {
  "town": "Apsley",
  "region": "Peterborough County (North Kawartha)",
  "population": "2,877 (North Kawartha, 2021 census); cottage area between Bancroft and Peterborough",
  "drive_from_bancroft": "About 38 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("North Kawartha Community Centre (NKCC)", "arena", "340 McFadden Road, Apsley, ON K0L 1A0",
     "Arena floor not published; banquet hall up to 325",
     "Township rentals page. The NHL-size pad is rented in the off-season for weddings, trade shows and sporting events.",
     "no", "Ice from late August to mid-March, so the dry floor runs about mid-March to late August.",
     "None found",
     "2026, plus HST: arena special-event flat rate $600.00/day; hall $300/day + $50 kitchen/bar",
     "705-656-4445; Shawn Tucker, Parks & Rec Supervisor, s.tucker@northkawartha.ca",
     "https://northkawartha.ca/living-here/facility-rentals",
     sc(2, 3, 2, 5, 5))
  ],
  "gyms_nearby": [],
  "combat_events": [],
  "clashes": [
   {"name": "Apsley Lions Home & Cottage Show (at NKCC)", "typical_dates": "Early May; May 2-3, 2026", "source_url": "https://northkawartha.ca/living-here/facility-rentals"},
   {"name": "Apsley Lions Giant Flea Market (at NKCC)", "typical_dates": "Mid-July; July 18-19, 2026", "source_url": "https://northkawartha.ca/living-here/facility-rentals"}
  ],
  "scores": sc(2, 3, 2, 5, 5),
  "verdict": "Cheapest rental ($600/day) and close, but the population is tiny and floor capacity is unpublished. Use it only for a smaller summer show or one that pulls from Peterborough."
 },
 {
  "town": "Madoc",
  "region": "Hastings County (Centre Hastings)",
  "population": "4,801 (Centre Hastings, 2021 census)",
  "drive_from_bancroft": "About 1 h 01 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("Madoc & District Recreation Centre (arena)", "arena", "45 Cooper Rd, Madoc, ON K0K 2K0",
     "Not published", "Single pad built 1976 on the fairgrounds; 250 parking spaces; has hosted dances, craft shows and trade fairs (Centre Hastings master plan 2023).",
     "no", "Ice from September to May, so the dry floor is about June to August.",
     "None found", "Not published", "613-473-4206",
     "https://www.centrehastings.com/media/ntcdeugb/centre-hastings-prc-master-plan-technical-compendium_may2023.pdf",
     sc(2, 2, 2, 5, 4))
  ],
  "gyms_nearby": [],
  "combat_events": [],
  "clashes": [
   {"name": "Madoc Fair (at the arena grounds)", "typical_dates": "3rd weekend of September; Sept 18-20, 2026", "source_url": "https://www.centrehastings.com/media/ntcdeugb/centre-hastings-prc-master-plan-technical-compendium_may2023.pdf"}
  ],
  "scores": sc(2, 2, 2, 5, 4),
  "verdict": "Not a 1,000+ market and there's no winter option. Marmora (25 min away) publishes a 1,000-person dry floor at $600/day and is the better summer pick in this area."
 },
 {
  "town": "Lindsay",
  "region": "City of Kawartha Lakes",
  "population": "79,247 (City of Kawartha Lakes, 2021 census)",
  "drive_from_bancroft": "About 2 h 02 min (Rome2rio estimate; earlier estimate ~1 h 50 min)",
  "over_3h": False,
  "venues": [
   v("Lindsay Recreation Complex", "arena", "133 Adelaide Street South, Lindsay",
     "Not published", "Two pads (Pad One/Pad Two); community room 180.",
     "no", "The City says many of its arenas switch to floors April to August for large special events; ice packages run fall/winter. No events over ice found.",
     "None found", "Not published (booked through the City's portal)",
     "Facility Booking 705-324-9411 ext. 1226, facilitybooking@kawarthalakes.ca; complex 705-324-9112",
     "https://www.kawarthalakes.ca/explore-play/facility-and-park-rentals/arena-and-ice-rentals/",
     sc(3, 2, 5, 4, 3)),
   v("Lindsay Exhibition Grounds", "fairgrounds", "50 Exhibition Drive, Lindsay",
     "Not published", "Home of the Lindsay Exhibition (LEX).",
     "no", "Fairgrounds, summer only.", "None found", "Not found", "Lindsay Agricultural Society",
     "https://ontariovisited.ca/las-presents-the-lex-2026/",
     sc(3, 2, 5, 4, 3))
  ],
  "gyms_nearby": [
   {"name": "Big Country - Lindsay (MMA, BJJ, Muay Thai, boxing, wrestling; sent an amateur champion to CCCL 1)", "source_url": "https://bigcountrylindsay.gymdesk.com/page/home"}
  ],
  "combat_events": [],
  "clashes": [
   {"name": "Lindsay Exhibition (LEX)", "typical_dates": "Late September; Sept 23-27, 2026", "source_url": "https://ontariovisited.ca/las-presents-the-lex-2026/"}
  ],
  "scores": sc(3, 2, 5, 4, 3),
  "verdict": "Strong fighter partner (Big Country), but no verified 1,000+ venue and no winter option. Better to bring Big Country fighters to Peterborough than to run in Lindsay."
 },
 {
  "town": "Kingston",
  "region": "Kingston / Frontenac (Eastern Ontario)",
  "population": "132,485 (City of Kingston, 2021 census)",
  "drive_from_bancroft": "About 2 h 36 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("Slush Puppie Place (formerly Leon's Centre / K-Rock Centre)", "arena", "1 The Tragically Hip Way, Kingston, ON K7K 0B4",
     "6,700 seats, or 5,700 for cold-season events (Visit Kingston guide); 5,000 for hockey (Wikipedia). No MMA configuration published.",
     "Visit Kingston Sport Facilities Guide; City of Kingston page (6,700 seats, 33 suites, in-house kitchen, 'available for private event rental').",
     "yes", "Proven for MMA in-season: BTC held cards here in March 2023 and March 2024, during the Frontenacs season. The guide's 'cold-season' capacity implies events over the ice. Dates must avoid Frontenacs home games (e.g., Jan 2, 13, 15, 30 and Feb 26, 2027 are listed).",
     "BTC 19: Fight Night Kingston (pro MMA, Mar 18, 2023) and BTC 23: Revolution (Mar 16, 2024), both as Leon's Centre",
     "Not published",
     "Managed by ASM Global; venue 613-650-5000; 'Private Venue Bookings' page on slushpuppieplace.com. Tourism Kingston sport team: rob@tourismkingston.com, 613-305-1735",
     "https://www.visitkingston.ca/media/documents/sports/Sport-Facilities-Guide.pdf",
     sc(4, 4, 3, 3, 2)),
   v("Kingston Memorial Centre", "arena", "303 York Street, Kingston",
     "3,300 seats (Wikipedia); Visit Kingston guide says 2,500, or bleachers for up to 1,000. Figures conflict.",
     "Home of Queen's Gaels hockey. City report (June 2026) on expanding and redeveloping the arena with Queen's.",
     "unclear", "Queen's hockey has ice in winter; no events over ice found. Redevelopment timing unknown.",
     "None found. The fairgrounds host the Kingston Fall Fair.", "Not published", "City of Kingston recreation",
     "https://www.visitkingston.ca/media/documents/sports/Sport-Facilities-Guide.pdf",
     sc(3, 2, 3, 3, 2)),
   v("Queen's Athletics & Recreation Centre (ARC) main gym", "event centre", "284 Earl Street, Kingston",
     "Main gym seats 1,904 (15,488 sq ft floor with seats retracted)", "Visit Kingston Sport Facilities Guide",
     "unclear", "Indoor gym with no ice, so winter is possible, but it's a university facility and no combat-event rental policy was found.",
     "None found", "Not published", "Queen's Athletics & Recreation",
     "https://www.visitkingston.ca/media/documents/sports/Sport-Facilities-Guide.pdf",
     sc(3, 2, 3, 3, 2)),
   v("INVISTA Centre", "arena", "1350 Gardiners Road, Kingston",
     "725 seats on the main pad (250 on each of three others) plus 1,000 standing room", "Visit Kingston Sport Facilities Guide",
     "no", "Four ice pads; no events over ice found.", "None found", "Not published", "City of Kingston",
     "https://www.visitkingston.ca/media/documents/sports/Sport-Facilities-Guide.pdf",
     sc(3, 2, 3, 3, 2))
  ],
  "gyms_nearby": [
   {"name": "YGK Combatives Group (BJJ, Muay Thai, Krav Maga), 745 Development Dr", "source_url": "https://letsroll.co/gym/ygk-combatives-group-kingston-on"},
   {"name": "Kingston Jiu Jitsu (BJJ, Muay Thai classes)", "source_url": "https://www.kingstonjiujitsu.com/new-muay-thai-classes/"}
  ],
  "combat_events": [
   {"name": "BTC 23: Revolution (BTC Fight Promotions, pro MMA)", "date": "2024-03-16", "venue": "Leon's Centre (now Slush Puppie Place)", "source_url": "https://www.sherdog.com/events/btc-23-revolution-101517"},
   {"name": "BTC 19: Fight Night Kingston (9 pro MMA bouts) - pre-2024 precedent", "date": "2023-03-18", "venue": "Leon's Centre", "source_url": "https://www.kingstonist.com/?p=138619"},
   {"name": "No 2025-2026 MMA or pro boxing card found in Kingston", "date": "", "venue": "", "source_url": ""}
  ],
  "clashes": [
   {"name": "Kingston Frontenacs (OHL) home games", "typical_dates": "Sept-March; e.g., Nov 27, 2026, Jan 2/13/15/30 and Feb 26, 2027", "source_url": "https://www.rateyourseats.com/tickets/events/london-knights-at-kingston-frontenacs-tickets-slush-puppie-place-january-2-2027-7174512"},
   {"name": "Spirit of Winter festival (downtown)", "typical_dates": "February; began Feb 1, 2026", "source_url": "https://www.cfrc.ca/news-events/news/2026/january/spirit-of-winter-festival-returns-to-downtown-for-a-second-year"},
   {"name": "Kingston Fall Fair (Memorial Centre fairgrounds, 16,000+ attendance)", "typical_dates": "Four days each September", "source_url": "https://en.wikipedia.org/wiki/Kingston_Memorial_Centre"}
  ],
  "scores": sc(4, 4, 3, 3, 2),
  "verdict": "Arena already proven for winter MMA (BTC held two March cards), in a 132k city. Good for a November show, but the 2.5 h+ drive and a thin known amateur MMA gym base are the drawbacks."
 },
 {
  "town": "Oshawa",
  "region": "Durham Region (east GTA)",
  "population": "175,383 (City of Oshawa, 2021 census); neighbours Whitby 138,501 and Clarington 101,427",
  "drive_from_bancroft": "About 2 h 29 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("Tribute Communities Centre (TCC)", "arena", "99 Athol Street East, Oshawa, ON L1H 1J8",
     "Main rink seating is flexible: 'can host events that require seating from 1,000 to 7,600'. Hockey 5,180 (6,125 with standing room).",
     "TCC Host an Event page; Wikipedia for hockey capacity.",
     "yes", "Holds non-hockey events in the Generals/FireWolves season (K-Pop Dreamworld Nov 7, 2026; Oshawa Beer Fest Nov 21, 2026). Risk: a $50M City modernization starts in 2027, phased over two years.",
     "Rise of the Phoenix (Durham Fight Series, Oshawa's first pro MMA card), June 3, 2023; boxing at the 2015 Pan Am Games",
     "Not published - 'Request a Proposal' via Oak View Group",
     "Oak View Group; 905-438-8881 x116; tributecommunitiescentre.com/host-an-event/request-a-proposal",
     TCC_SRC, sc(5, 4, 5, 2, 3)),
   v("Children's Arena (City of Oshawa)", "arena", "155 Arena St, Oshawa (Kinsmen Valleyview Park)",
     "Not published", "City lists it as 'floor only, no ice'.",
     "yes", "Floor-only building with no ice, so it's usable year-round. Indoor and heated status not confirmed.",
     "BTC 25: Legacy (BTC Fight Promotions), Sept 14, 2024",
     "Floor, commercial: $120/hr + HST (effective Apr 1, 2026); no event day rate published",
     "Service Oshawa 905-436-3311, service@oshawa.ca; online Arena Rental Application",
     "https://www.oshawa.ca/explore-play/recreation/facilities-and-rentals/arena-rentals/",
     sc(4, 3, 5, 2, 3)),
   v("Bond Street Event Centre (BOND|ST)", "event centre", "44 Bond Street East, Oshawa",
     "Not published", "Former UAW Local 222 hall turned into an event centre (Mainstage music venue, lounges).",
     "yes", "Indoor event centre with no ice.",
     "PFC 24 (Prospect Fighting Championships) was booked here for June 20, 2026, then moved to Caledonia because of the FIFA World Cup",
     "Not published", "bondst.ca",
     "https://doorsopenontario.on.ca/oshawa/bond-st-event-centre",
     sc(4, 2, 5, 2, 3)),
   v("General Sikorski Hall", "hall", "1551 Stevenson Rd N, Oshawa",
     "565 max (Grand Ballroom 500), so under 1,000", "Eventective listing",
     "yes", "Indoor banquet hall.", "BTC 31: Invictus, Aug 9, 2025",
     "Not published", "289-356-6399",
     "https://www.eventective.com/oshawa-on/general-sikorski-hall-596553.html",
     sc(2, 1, 5, 2, 3))
  ],
  "gyms_nearby": [
   {"name": "Oshawa BJJ (BJJ, Muay Thai, kickboxing, wrestling)", "source_url": "https://oshawabjj.com"},
   {"name": "Wider Durham/GTA-east gym base (estimate - not individually verified)", "source_url": ""}
  ],
  "combat_events": [
   {"name": "BTC 25: Legacy (BTC Fight Promotions)", "date": "2024-09-14", "venue": "Children's Arena, 155 Arena St", "source_url": "https://calendar.oshawa.ca/default/Detail/2024-09-14-1700-BTC-25-Legacy-Presented-by-BTC-Fight-Promotions"},
   {"name": "BTC 31: Invictus (10 MMA bouts)", "date": "2025-08-09", "venue": "Sikorski Hall", "source_url": "https://www.tapology.com/forum/threads/87443?page=1"},
   {"name": "PFC 24 (amateur, Prospect Fighting Championships) - booked, then moved to Caledonia; organizer said it plans to return to Oshawa in November", "date": "2026-06-20 (postponed)", "venue": "Bond Street Event Centre", "source_url": "https://mm-eh.ca/upcoming-fights/prospect-fighting-championship-24-1/"},
   {"name": "Rise of the Phoenix (Durham Fight Series, pro MMA) - pre-2024", "date": "2023-06-03", "venue": "Tribute Communities Centre", "source_url": "https://calendar.oshawa.ca/default/Detail/2023-06-03-1900-Rise-of-the-Phoenix-MMA"}
  ],
  "clashes": [
   {"name": "Oshawa Generals (OHL) and Oshawa FireWolves (NLL) home dates at TCC", "typical_dates": "Generals Sept-March; FireWolves home opener Nov 27, 2026, season into spring", "source_url": TCC_SRC},
   {"name": "Oshawa Peony Festival (next to Children's Arena, 155 Arena St)", "typical_dates": "Early June; June 6-7, 2026", "source_url": "https://ontariovisited.ca/?p=284728"},
   {"name": "PFC's planned November return to Oshawa", "typical_dates": "November (2026 plan)", "source_url": "https://mm-eh.ca/upcoming-fights/prospect-fighting-championship-24-1/"}
  ],
  "scores": sc(5, 4, 5, 2, 3),
  "verdict": "Biggest market within 2.5 h and a winter-capable arena that scales down to 1,000, but crowded: BTC ran there in 2024 and 2025, PFC wants back in, and pro boxing runs in Pickering. TCC's renovation starts in 2027."
 },
 {
  "town": "Niagara Falls",
  "region": "Niagara Region",
  "population": "94,415 (City of Niagara Falls, 2021 census) plus heavy tourist traffic",
  "drive_from_bancroft": "About 4 h 29 min (Rome2rio estimate) - OVER 3 HOURS",
  "over_3h": True,
  "venues": [
   v("Niagara Falls Convention Centre", "event centre", "6815 Stanley Avenue, Niagara Falls, ON L2G 3Y9",
     "Exhibit Hall A 2,142 theatre; Exhibit Hall ABC 6,400 theatre; Peller Estates Ballroom 1,722 theatre. No boxing configuration published.",
     "fallsconventions.com rooms & spaces",
     "yes", "Indoor convention centre with no ice.",
     "Pro boxing: Lucas Bahdi bouts Jul 2, 2022, May 21, 2023 and Jun 1, 2024. Tarps Off Fight Club 'Brawl by the Falls' MMA, Dec 18, 2021, listed as 'ScotiaBank Convention Centre'.",
     "Not published", "Niagara Falls Convention Centre sales (fallsconventions.com)",
     "https://fallsconventions.com/plan/rooms-spaces",
     sc(4, 4, 3, 2, 1)),
   v("OLG Stage at Fallsview Casino", "casino", "Fallsview Casino Resort, Niagara Falls",
     "5,000-seat theatre", "Canadian Gaming Business, Sept 2022",
     "yes", "Indoor casino theatre.", "None found (concerts and comedy)",
     "Not published", "MGE Niagara Entertainment (Mohegan), operator for OLG",
     "https://www.canadiangamingbusiness.com/2022/09/01/fallsview-olg-partner-up-on-new-casino-theatre/",
     sc(4, 3, 3, 2, 1)),
   v("Americana Resort (Grande Ballroom)", "event centre", "8444 Lundy's Lane, Niagara Falls, ON L2H 1H4",
     "Grande Ballroom 750 theatre / 625 banquet, so under 1,000", "meetings-conventions.com listing",
     "yes", "Indoor hotel ballroom.", "BTC 32: The Battle of Lundy's Lane (pro-am MMA), Sept 13, 2025",
     "Not published", "Americana Resort sales",
     "https://www.meetings-conventions.com/Meeting-Event-Venues/Niagara-Falls-ON/Convention-Hotel/Americana-Resort-Spa-Waterpark-p2357166",
     sc(3, 2, 3, 2, 1))
  ],
  "gyms_nearby": [
   {"name": "Niagara BJJ, 4536 Montrose Rd unit 4, Niagara Falls", "source_url": "https://letsroll.co/gym/niagara-bjj-niagara-falls-niagara-falls-on"},
   {"name": "Renzo Gracie Niagara (BJJ), 486 Grantham Ave, St. Catharines (~20 min, estimate)", "source_url": "https://gymdesk.com/originals/renzo-gracie-niagara"}
  ],
  "combat_events": [
   {"name": "BTC 32: The Battle of Lundy's Lane (4 pro + 6 amateur MMA bouts)", "date": "2025-09-13", "venue": "Americana Resort", "source_url": "https://cagesidepress.com/2025/09/13/btc-32-the-battle-of-lundys-lane-full-results"},
   {"name": "Lucas Bahdi pro boxing card", "date": "2024-06-01", "venue": "Niagara Falls Convention Centre", "source_url": "https://en.wikipedia.org/wiki/Lucas_Bahdi"},
   {"name": "PFC 23: Appel vs. Anaya (Prospect Fighting Championships; nearby Niagara-on-the-Lake)", "date": "2026-05-09", "venue": "Central Niagara Center, Niagara-on-the-Lake", "source_url": "https://en.wikipedia.org/wiki/List_of_PFC_events"}
  ],
  "clashes": [
   {"name": "Winter Festival of Lights", "typical_dates": "Mid-Nov to early Jan; Nov 14, 2026 to Jan 10, 2027", "source_url": "https://www.cliftonhill.com/events/niagara-falls-winter-festival-of-lights"}
  ],
  "scores": sc(4, 4, 3, 2, 1),
  "verdict": "Good venues and tourist traffic, but 4.5 h from Bancroft, outside Cody's 3 h limit, and already served by BTC, PFC and pro boxing. Not for 2027."
 },
 {
  "town": "Alliston (New Tecumseth)",
  "region": "Simcoe County",
  "population": "43,948 (Town of New Tecumseth, 2021 census)",
  "drive_from_bancroft": "About 3 h 16 min (Rome2rio estimate) - OVER 3 HOURS",
  "over_3h": True,
  "venues": [
   v("Nottawasaga Inn Resort - Trade Centre", "event centre", "6015 Highway 89, Alliston, ON L9R 1A4",
     "Trade Centre 20,358 sq ft: 1,000 theatre / 1,000 banquet", "successfulmeetings.com venue listing",
     "yes", "Indoor resort conference hall.", "None found", "Not published", "Nottawasaga Resort sales",
     "https://www.successfulmeetings.com/Meeting-Event-Venues/Alliston-ON/Convention-Hotel/Nottawasaga-Inn-Resort-Conference-Ctr-p5578841",
     sc(3, 3, 2, 4, 1)),
   v("New Tecumseth Recreation Centre (main arena)", "arena", "Alliston (address not retrieved)",
     "Main arena seats 750 (2021)", "New Tecumseth Times, Sept 16, 2021 (roof replacement article)",
     "no", "Ice arena; no events over ice found.", "None found", "Not published", "Town of New Tecumseth",
     "https://newtectimes.com/new-tecumseth-recreation-centre-getting-a-new-roof/",
     sc(3, 2, 2, 4, 1))
  ],
  "gyms_nearby": [
   {"name": "No combat gym verified in Alliston. Nearest found: Empire Martial Arts and Fitness, Barrie (~35-40 min, estimate)", "source_url": "https://bradfordtoday.ca/local-sports/great-thing-veteran-mma-fighter-gives-local-gym-a-punch-in-the-arm-7416978"}
  ],
  "combat_events": [
   {"name": "None found in Alliston. Regional: OFF1: First Strike (Official Fight Federation, pro-am MMA, Barrie; promoter expected 3,500+)", "date": "2025-05-09", "venue": "Sadlon Arena, Barrie", "source_url": "https://orilliamatters.com/local-sports/barries-first-ever-pro-am-mma-event-happening-in-may-10382588"}
  ],
  "clashes": [
   {"name": "Alliston Potato Festival (Church Street Fairgrounds)", "typical_dates": "Early-mid August; Aug 8-10, 2025", "source_url": "https://ontariovisited.ca/alliston-potato-festival-2025"}
  ],
  "scores": sc(3, 3, 2, 4, 1),
  "verdict": "Over the 3 h limit, with no local fight gym found and only a 1,000-seat hotel hall. Not worth it in 2027."
 },
 {
  "town": "Rama / Orillia (Casino Rama)",
  "region": "Simcoe County (Rama First Nation, beside Orillia)",
  "population": "Orillia 33,411 (2021 census); Casino Rama draws regionally",
  "drive_from_bancroft": "About 2 h 20 min to Orillia (Rome2rio estimate); Casino Rama is just northeast of Orillia",
  "over_3h": False,
  "venues": [
   v("Casino Rama Resort - Entertainment Centre", "casino", "5899 Rama Rd, Rama, ON",
     "Over 5,000 seats (Wikipedia); no boxing/MMA configuration published", "Wikipedia",
     "yes", "Indoor casino theatre running live entertainment year-round.",
     "Pro boxing Jan 19, 2008 (Orion Sports Management); MMA Apr 2, 2011, billed as Ontario's first sanctioned MMA event; Bellator events 2011 and Bellator 67 (May 2012). Nothing found for 2024-2026.",
     "Not published", "Casino Rama Resort entertainment / events department",
     "https://en.wikipedia.org/wiki/Casino_Rama",
     sc(4, 3, 3, 4, 3))
  ],
  "gyms_nearby": [
   {"name": "Big Country network (Lindsay/Orillia/Bracebridge, per Big Country Lindsay site)", "source_url": "https://bigcountrylindsay.gymdesk.com/page/home"}
  ],
  "combat_events": [
   {"name": "None found 2024-2026 at Casino Rama. Regional: OFF1 at Sadlon Arena, Barrie", "date": "2025-05-09", "venue": "Sadlon Arena, Barrie", "source_url": "https://www.ctvnews.ca/barrie/article/barrie-hosts-first-pro-amateur-mixed-martial-arts-fight-night/"}
  ],
  "clashes": [],
  "scores": sc(4, 3, 3, 4, 3),
  "verdict": "Big winter-proof room with MMA history, but no combat since about 2012 and no rental terms. Worth one inquiry as a future 'big show', not a 2027 base."
 },
 {
  "town": "Pickering (Pickering Casino Resort)",
  "region": "Durham Region (next to Oshawa)",
  "population": "Not retrieved; part of Durham Region beside Oshawa (175,383) and Whitby (138,501)",
  "drive_from_bancroft": "About 2 h 40 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("The Arena at Pickering Casino Resort", "casino", "Pickering Casino Resort (Durham Live), Pickering - street address not retrieved",
     "2,500 seats (opened Apr 26, 2023)", "blogTO, Mar 2023",
     "yes", "Indoor casino arena; held a pro boxing card in December 2024.",
     "United Boxing Promotions 'Fight Night at the Casino' pro boxing: 4th edition Oct 21, 2023 (first three reported sold out); Sept 20, 2024; Oct 19, 2024 (announced); about Dec 14, 2024 (inferred from a Dec 17 article)",
     "Not published", "Pickering Casino Resort events",
     "https://blogto.com/music/2023/03/arena-pickering-casino-resort-officially-opening",
     sc(4, 4, 5, 2, 2))
  ],
  "gyms_nearby": [
   {"name": "Oshawa BJJ (~20-30 min, estimate)", "source_url": "https://oshawabjj.com"}
  ],
  "combat_events": [
   {"name": "United Boxing Promotions fight night (pro boxing)", "date": "2024-09-20", "venue": "Pickering Casino Resort", "source_url": "https://tcn-now.com/united-boxing-promotions-finally-back-after-a-summer-break-i-dont-think-we-were-ready/"},
   {"name": "United Boxing Promotions fight night (pro boxing; date inferred as Sat before the Dec 17 article)", "date": "2024-12-14 (approx.)", "venue": "Pickering Casino Resort", "source_url": "https://tcn-now.com/lights-out-united-boxing-promotions-shocks-fans-with-unstoppable-knockouts/"}
  ],
  "clashes": [
   {"name": "United Boxing Promotions' recurring casino fight nights at the same venue", "typical_dates": "Several per year (2023-2024)", "source_url": "https://tcn-now.com/fourth-exhilarating-edition-of-united-boxing-promotions-fight-night-at-pickering-casino-resort/amp/"}
  ],
  "scores": sc(4, 4, 5, 2, 2),
  "verdict": "Right-sized (2,500) winter-proof casino arena that already does combat nights, but it's 2 h 40 min away in a crowded market. A fallback if TCC or PMC fall through."
 },
 {
  "town": "Port Perry (Great Blue Heron Casino)",
  "region": "Durham Region (Scugog)",
  "population": "21,581 (Township of Scugog, 2021 census)",
  "drive_from_bancroft": "About 2 h 27 min (Rome2rio estimate)",
  "over_3h": False,
  "venues": [
   v("Great Blue Heron Casino & Hotel - live entertainment/special events venue", "casino", "21777 Island Road, Port Perry",
     "About 800 seats (12,637 sq ft; banquet, theatre or GA), so under 1,000", "CTV News / CDC Gaming, Jan 18, 2024",
     "yes", "Indoor casino venue.", "None found", "Not published", "Great Blue Heron Casino events",
     "https://toronto.ctvnews.ca/a-new-entertainment-venue-just-opened-in-the-gta-1.6736507",
     sc(3, 2, 3, 3, 3))
  ],
  "gyms_nearby": [
   {"name": "Oshawa BJJ (~30 min, estimate)", "source_url": "https://oshawabjj.com"}
  ],
  "combat_events": [],
  "clashes": [],
  "scores": sc(3, 2, 3, 3, 3),
  "verdict": "Too small (about 800) for the 1,000+ target. Skip."
 }
]

promotions = [
 {"name": "BTC Fight Promotions (Burlington; owner Irene Starr; pro-am MMA since amateur MMA was legalized on Jan 1, 2025)",
  "markets": "Kingston, Oshawa, Belleville, Niagara Falls, Burlington, London",
  "venues": "Leon's Centre/Slush Puppie Place Kingston (Mar 18, 2023; Mar 16, 2024); Children's Arena Oshawa (Sept 14, 2024); Belleville Armouries (BTC 27, Mar 15, 2025, 6 am + 6 pro); Sikorski Hall Oshawa (BTC 31, Aug 9, 2025); Americana Resort Niagara Falls (BTC 32, Sept 13, 2025); Central Arena Burlington (BTC 37, June 6, 2026)",
  "frequency": "About 5-7 cards a year (estimate from numbering: BTC 25 in Sept 2024 to BTC 37 in June 2026). Moves around Ontario and has already worked Kingston and Oshawa.",
  "source_url": "https://anishinabeknews.ca/2025/01/historic-mixed-martial-arts-event-will-be-staged-in-belleville/"},
 {"name": "Prospect Fighting Championships (PFC; London-based; founder Jamie Champion; pro and amateur; OAC-sanctioned)",
  "markets": "London, Sarnia, Niagara-on-the-Lake, Caledonia; Oshawa planned",
  "venues": "2026: RBC Place London (Feb 28; Sept 26), Royal SCITS Academy Sarnia (Apr 25), Central Niagara Center NOTL (May 9), Caledonia Fairgrounds (Jul 18, moved from Bond Street Event Centre Oshawa), Sarnia (Aug 15)",
  "frequency": "Six events in 2026, about one every 1-2 months. Organizer said it plans to return to Oshawa in November.",
  "source_url": "https://en.wikipedia.org/wiki/List_of_PFC_events"},
 {"name": "Official Fight Federation (OFF; 'presented by Gary Goodridge')",
  "markets": "Barrie / Simcoe County",
  "venues": "Sadlon Arena, Barrie - OFF1: First Strike, pro-am, May 9, 2025 (promoter expected 3,500+)",
  "frequency": "One event found; no follow-up found",
  "source_url": "https://orilliamatters.com/local-sports/barries-first-ever-pro-am-mma-event-happening-in-may-10382588"},
 {"name": "United Boxing Promotions (pro boxing; president Tyler Buxton)",
  "markets": "Pickering / Durham",
  "venues": "The Arena, Pickering Casino Resort (2,500 seats)",
  "frequency": "Recurring 'Fight Night at the Casino': at least 4 by Oct 2023 plus Sept, Oct and Dec 2024; 2025-26 schedule not found",
  "source_url": "https://tcn-now.com/fourth-exhilarating-edition-of-united-boxing-promotions-fight-night-at-pickering-casino-resort/amp/"},
 {"name": "Durham Fight Series (pro MMA)",
  "markets": "Oshawa",
  "venues": "Tribute Communities Centre - Rise of the Phoenix, June 3, 2023",
  "frequency": "One event found (2023); no 2024-2026 events found",
  "source_url": "https://calendar.oshawa.ca/default/Detail/2023-06-03-1900-Rise-of-the-Phoenix-MMA"},
 {"name": "Cottage Country Combat League (CCCL) - Cody's promotion",
  "markets": "Haliburton / Bancroft",
  "venues": "A.J. LaRue Arena, Haliburton - CCCL 1, May 2, 2026",
  "frequency": "One so far",
  "source_url": "https://thehighlander.ca/2026/03/12/fight-night-coming-to-haliburton-county/"}
]

notes = (
 "SCORING: 1-5 each; overall = simple average. Venue scores are attached to each venue; the town score uses its best realistic venue. "
 "Distance follows the brief (5 <1 h, 4 1-2 h, 3 2-2.5 h, 2 2.5-3 h, 1 >3 h) using Rome2rio drive times (estimates; Rome2rio km figures looked unreliable, so only times are given). "
 "'crowd' rates the chance of 1,000+; CCCL 1 sold about 400 tickets with 400-600 in the building, so no small-town venue is proven at 1,000. "
 "WINTER (Jan/Feb/Nov) - venues with evidence they hold dry events in winter: Peterborough Memorial Centre (Nov 21, 2026 concert; OFSAA wrestling Mar 2026), Slush Puppie Place Kingston (BTC MMA in March 2023 and 2024; '5,700 for cold-season events'), Tribute Communities Centre Oshawa (Nov 2026 shows; renovation from 2027), Oshawa Children's Arena (floor only, no ice; capacity unknown), Pickering Casino Resort Arena (2,500; pro boxing Dec 2024), Casino Rama (5,000; no combat since ~2012), Niagara Falls Convention Centre (over 3 h), Nottawasaga Trade Centre Alliston (1,000; over 3 h). "
 "No small-town arena (Bancroft, Haliburton, Apsley, Madoc, Lindsay) showed evidence of covering ice for winter events. Shorelines Casino Peterborough has no event space mentioned (Global News 2018). Great Blue Heron's room is about 800. "
 "SUMMER 1,000 OPTION FROM EARLIER RESEARCH: Marmora's Dr. Hamilton Crawford Memorial Arena publishes a 1,000 dry-floor capacity at $600/day, May-Aug (https://marmoraandlake.ca/community-centre). It isn't on Cody's town list but is the only small-town venue with a published 1,000 figure. "
 "LEAD TIME: OAC amateur event permits need to be filed at least 30 days ahead, and modified-ruleset exemptions 120 days ahead (ontario.ca, from earlier CCCL research). A January 2027 show (under 3 months away) is very tight; February is possible only if the venue is booked now. "
 "KEY UNKNOWNS: rental costs at PMC, Slush Puppie Place, TCC, Pickering and Casino Rama (none published); arena-floor capacities in Bancroft, Haliburton, Apsley, Madoc and Lindsay; OHL home schedules for 2027; TCC renovation phasing; whether casinos will rent to an amateur promotion; PFC's November Oshawa plan and BTC's 2027 calendar."
)

data = {"towns": towns, "promotions": promotions, "notes": notes}
with open(OUT, "w") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
print("written", OUT)

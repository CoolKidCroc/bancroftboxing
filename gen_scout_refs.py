"""Write cccl-event-scout references: venues-2027.md (from research JSON) and money-model.md (from the recalculated Predictor).
usage: python3 gen_scout_refs.py research/venues_2027.json predictor.xlsx skills/cccl-event-scout/references"""
import json, sys, os
from openpyxl import load_workbook

VJ, XL, OUTDIR = sys.argv[1], sys.argv[2], sys.argv[3]
d = json.load(open(VJ))

# ---------------------------------------------------------------- venues-2027.md
L = ["# CCCL 2027 venues (researched October 2026)", "",
     "Cody's rules for 2027 (Oct 2026): 4+ shows, up to 3 hours from Bancroft, crowds of 1,000 or more, months Jan, Feb, Apr, May, Jun, Aug, Nov. "
     "Venue types he's open to: hockey arena, community hall, fairgrounds (summer), casino or event centre.", "",
     "Scores 1-5 (5 = best): crowd (chance of 1,000+), venue, fighters (gyms nearby), competition (5 = little), distance "
     "(5 under 1 h, 4 1-2 h, 3 2-2.5 h, 2 2.5-3 h, 1 over 3 h). Overall = simple average. Drive times are Rome2rio estimates. "
     "Capacities and rates are quoted as published; re-check before Cody relies on them.", "",
     "## Key points", "",
     "- **Winter (Jan, Feb, Nov)**: small-town arenas have ice. Venues with evidence of winter floor events: Peterborough Memorial Centre, "
     "Slush Puppie Place (Kingston), Tribute Communities Centre and the Children's Arena (Oshawa), Pickering Casino Resort arena, Casino Rama.",
     "- **1,000+ crowd**: show 1 sold about 400 tickets with 400-600 people in the building. No small-town venue is proven at 1,000. "
     "Marmora's arena is the only small venue that publishes a 1,000-person dry floor ($600/day, May-Aug).",
     "- **Over 3 hours**: Niagara Falls (~4 h 29 min) and Alliston (~3 h 16 min). Not for 2027.",
     "- **Lead time**: OAC event permit at least 30 days before the show; rule-change exemptions 120 days before. January 2027 is very tight.",
     "- **Unknowns**: no big venue publishes rental costs; small-town floor capacities aren't published; 2027 OHL home schedules decide winter dates; "
     "casinos may not rent to an amateur promotion.", "",
     "## Towns, best first", ""]
for t in sorted(d["towns"], key=lambda t: (t.get("over_3h", False), -t["scores"]["overall"])):
    sc = t["scores"]
    L.append(f"### {t['town']}: {sc['overall']} overall" + ("  (OVER 3 HOURS)" if t.get("over_3h") else ""))
    L.append(f"Crowd {sc['crowd']} · venue {sc['venue']} · fighters {sc['fighters']} · competition {sc['competition']} · distance {sc['distance']}")
    L.append(f"- Verdict: {t['verdict']}")
    L.append(f"- Drive from Bancroft: {t['drive_from_bancroft']} · Population: {t.get('population', '')}")
    for v in t["venues"]:
        L.append(f"- **Venue: {v['name']}** ({v['type']}, score {v['scores']['overall']}). Capacity: {v['capacity_event']} "
                 f"Winter: {v['winter_ok']}. {v.get('winter_note', '')} Rate: {v['rental_rate']}. Combat history: {v.get('hosted_combat', '')} "
                 f"Contact: {v['booking_contact']}. Source: {v.get('source_url', '')}")
    gyms = "; ".join(g["name"] for g in t.get("gyms_nearby", []))
    if gyms: L.append(f"- Gyms nearby: {gyms}")
    ev = "; ".join(f"{e['name']} ({e['date']})" if e.get("date") else e["name"] for e in t.get("combat_events", []))
    if ev: L.append(f"- Combat events: {ev}")
    cl = "; ".join(f"{c['name']}: {c['typical_dates']}" for c in t.get("clashes", []))
    if cl: L.append(f"- Clashes: {cl}")
    L.append("")
L += ["## Other promotions", ""]
for p in d["promotions"]:
    L.append(f"- **{p['name']}**. Markets: {p.get('markets', '')}. Venues: {p.get('venues', '')}. How often: {p.get('frequency', '')}. Source: {p.get('source_url', '')}")
open(os.path.join(OUTDIR, "venues-2027.md"), "w").write("\n".join(L) + "\n")

# ---------------------------------------------------------------- money-model.md
wb = load_workbook(XL, data_only=True)
wf = load_workbook(XL)  # labels
I, C, H = wb["Inputs"], wb["Compare"], wb["Haliburton 2026 Check"]
lab = {}
for row in C.iter_rows(min_row=4, max_row=C.max_row):
    if row[0].value: lab.setdefault(row[0].value, []).append(row[0].row)
def crow(name, k=0): return lab[name][k]
NV = 6
cols = [chr(ord("B") + i) for i in range(NV)]
names = [C[f"{c}{crow('Venue')}"].value for c in cols]
def val(name, c, k=0): return C[f"{c}{crow(name, k)}"].value or 0
def money(x): return f"${x:,.0f}" if x >= 0 else f"−${-x:,.0f}"
ilab = {I.cell(r, 1).value: r for r in range(1, I.max_row + 1) if I.cell(r, 1).value}
def g(name): return I.cell(ilab[name], 2).value
avg = I.cell(ilab["Average price per paid ticket"], 2).value
net = val("Ticket money after fees and HST (per paid ticket)", "B")
M = ["# CCCL money model (same math as the Event Profit Predictor)", "",
     "Use the numbers on the Predictor's Inputs tab if you can open it. These defaults use Cody's real 2026 prices and costs (October 2026). "
     "Anything not from Cody or a published source is an ESTIMATE.", "",
     "## The first show's real result", "",
     f"Cody's numbers for the Haliburton show (400 tickets, $18,000 sponsors, $29,000 costs) add up to about {money(round(H['B' + str([r for r in range(1, H.max_row + 1) if H.cell(r, 1).value == 'PROFIT THESE NUMBERS ADD UP TO'][0])].value, -3))} profit, "
     "and that profit went into the gym and other projects (Kyle, Oct 2026), which is why Cody said it broke even. The profits below are the show's own "
     "profit, before any money goes to the gym. The exact figure will come from the Eventbrite payout report.", "",
     "## Adults-only (19+) shows", "",
     "Cody is open to adults-only shows in 2027 if the money works. The Predictor's '19+ Show Check' tab compares a family show, a family show with a "
     "fenced 19+ bar area, and an adults-only show. It now uses Mat McTrash's offer (Cody, Oct 20, 2026): CCCL gets 50% of bar sales where Mat McTrash runs the bar. "
     "With that and the other estimates (crowd drops 20% without kids and families, $12 a head at a family show's 19+ area, $25 a head at an adults-only show), "
     "a family show with a fenced 19+ bar area makes the most at every venue, and adults-only makes the least: each paid ticket is worth about $73 after fees. "
     "Adults-only only beats the family show with a 19+ area if CCCL gets about $21 of bar money per adult (at 50%, adults spending about $41-42 each), or the crowd drops by about 10% or less. Other venues' bar terms are still unknown. "
     "Alcohol sponsors need a mostly-adult audience, so they only fit adults-only shows or a separate 19+ area. A for-profit promoter usually can't hold the "
     "public-event liquor permit itself (AGCO), so ask each venue who runs the bar and what share CCCL gets. Not legal advice.", "",
     "## Inputs (defaults)",
     f"- Crowd as % of sellable capacity: low {g('Low crowd: % of sellable capacity that shows up'):.0%}, expected {g('Expected crowd: % of sellable capacity'):.0%}, "
     f"high {g('High crowd: % of sellable capacity'):.0%} (2026: about 500 people in a building of about 1,000)",
     f"- People who didn't buy a ticket: {g('People who didn' + chr(39) + 't buy a ticket: % of the crowd'):.0%} of the crowd (2026: 400 tickets, about 500 people)",
     f"- Cody's minimum crowd: {g('Smallest crowd worth doing a show for (people)'):,} people",
     f"- Ticket tiers (2026 prices, from Cody): general admission $60 (78% of paid), VIP seat $100 (10%, estimate), VIP table $1,500 seating 8-10 (the model uses 8 = $187.50 a seat, 12%) "
     f"→ average ${avg:,.2f} per paid ticket",
     f"- Eventbrite fees CCCL pays: {g('Ticketing fees CCCL pays (% of ticket sales)'):.1%} (Eventbrite Canada: 3.5% + $1.29 per ticket plus 2.9% processing, blended over the 2026 ticket mix). HST taken out of ticket prices: "
     f"{g('HST inside the ticket price that goes to the government (%)'):.0%} unless CCCL is HST-registered with tax-included prices.",
     "- Bar and merch profit: $0 (the arena or a local group ran the bar in 2026). Livestream income: $0 (it didn't make money in 2026).",
     f"- Money per paid ticket after fees: ${net:,.2f}",
     "- Contingency: 5% of costs. Stress test: crowd 30% below expected and half the sponsorship falls through.", "",
     "## Costs (Cody's 2026 list, $29,000)",
     "Arena $1,500 · cage $3,000 · officials $3,000 · medical $2,000 · insurance $2,000 · security $1,500 · lights/sound/stream $1,000 · "
     "everything else $15,000 (split in the Predictor as: OAC permit $600, marketing $2,000, athlete costs $1,500, seating $1,000, crew food $600, "
     "the rest $9,300). The three city arenas get higher security, lights/sound, seating, athlete and crew costs (estimates), and their rent is an "
     "estimate until they quote.", "",
     "## Formulas",
     "- Paid tickets = round(capacity × crowd% × (1 − non-paying%))",
     "- Money per paid ticket = average price × (1 − fee% − HST%) + (bar + merch) ÷ (1 − non-paying%)",
     "- Break-even paid tickets = (total cost − sponsorship − livestream income) ÷ money per paid ticket, rounded up",
     "- Sponsorship needed on a low night = total cost − livestream − low-night paid tickets × money per paid ticket",
     "- Profit = ticket money after fees + bar/merch × crowd + sponsorship + livestream − total cost", "",
     "## Reference results (defaults, expected crowd = 50% of capacity)", "",
     "| Venue | Capacity | Sponsorship assumed | Total cost | Break-even paid tickets | Expected crowd | Profit at expected crowd | Sponsorship needed on a low night |",
     "|---|---|---|---|---|---|---|---|"]
for c, n in zip(cols, names):
    if not val("Sellable capacity (people)", c): continue
    M.append(f"| {n} | {val('Sellable capacity (people)', c):,} | {money(val('Sponsorship you expect (from Inputs)', c))} | {money(val('TOTAL EVENT COST', c))} | "
             f"{val('Paid tickets needed (with expected sponsors)', c):,} | {val('Crowd (people in the building)', c, 1):,} | "
             f"{money(val('PROFIT / (LOSS)', c, 1))} | {money(val('Sponsorship needed to break even on a LOW night', c))} |")
M += ["", "Venue notes: Haliburton uses the real 2026 arena bill. Bancroft and Haliburton capacities are estimates. Marmora's 1,000 and $600/day are published. "
      "Peterborough, Kingston and Oshawa are set to 2,000 seats with estimated rent: get quotes. Any venue whose expected crowd is under 1,000 is below Cody's minimum."]
open(os.path.join(OUTDIR, "money-model.md"), "w").write("\n".join(M) + "\n")
print("wrote venues-2027.md and money-model.md")

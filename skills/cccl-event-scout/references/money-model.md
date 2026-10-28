# CCCL money model (same math as the Event Profit Predictor)

Use the numbers on the Predictor's Inputs tab if you can open it. These defaults use Cody's real 2026 prices and costs (October 2026). Anything not from Cody or a published source is an ESTIMATE.

## The first show's real result

Cody's numbers for the Haliburton show (400 tickets, $18,000 sponsors, $29,000 costs) add up to about $18,000 profit, and that profit went into the gym and other projects (Kyle, Oct 2026), which is why Cody said it broke even. The profits below are the show's own profit, before any money goes to the gym. The exact figure will come from the Eventbrite payout report.

## Adults-only (19+) shows

Cody is open to adults-only shows in 2027 if the money works. The Predictor's '19+ Show Check' tab compares a family show, a family show with a fenced 19+ bar area, and an adults-only show. It now uses Mat McTrash's offer (Cody, Oct 20, 2026): CCCL gets 50% of bar sales where Mat McTrash runs the bar. With that and the other estimates (crowd drops 20% without kids and families, $12 a head at a family show's 19+ area, $25 a head at an adults-only show), a family show with a fenced 19+ bar area makes the most at every venue, and adults-only makes the least: each paid ticket is worth about $73 after fees. Adults-only only beats the family show with a 19+ area if CCCL gets about $21 of bar money per adult (at 50%, adults spending about $41-42 each), or the crowd drops by about 10% or less. Other venues' bar terms are still unknown. Alcohol sponsors need a mostly-adult audience, so they only fit adults-only shows or a separate 19+ area. A for-profit promoter usually can't hold the public-event liquor permit itself (AGCO), so ask each venue who runs the bar and what share CCCL gets. Not legal advice.

## Inputs (defaults)
- Crowd as % of sellable capacity: low 35%, expected 50%, high 80% (2026: about 500 people in a building of about 1,000)
- People who didn't buy a ticket: 20% of the crowd (2026: 400 tickets, about 500 people)
- Cody's minimum crowd: 1,000 people
- Ticket tiers (2026 prices, from Cody): general admission $60 (78% of paid), VIP seat $100 (10%, estimate), VIP table $1,500 seating 8-10 (the model uses 8 = $187.50 a seat, 12%) → average $79.30 per paid ticket
- Eventbrite fees CCCL pays: 7.9% (Eventbrite Canada: 3.5% + $1.29 per ticket plus 2.9% processing, blended over the 2026 ticket mix). HST taken out of ticket prices: 0% unless CCCL is HST-registered with tax-included prices.
- Bar and merch profit: $0 (the arena or a local group ran the bar in 2026). Livestream income: $0 (it didn't make money in 2026).
- Money per paid ticket after fees: $73.04
- Contingency: 5% of costs. Stress test: crowd 30% below expected and half the sponsorship falls through.

## Costs (Cody's 2026 list, $29,000)
Arena $1,500 · cage $3,000 · officials $3,000 · medical $2,000 · insurance $2,000 · security $1,500 · lights/sound/stream $1,000 · everything else $15,000 (split in the Predictor as: OAC permit $600, marketing $2,000, athlete costs $1,500, seating $1,000, crew food $600, the rest $9,300). The three city arenas get higher security, lights/sound, seating, athlete and crew costs (estimates), and their rent is an estimate until they quote.

## Formulas
- Paid tickets = round(capacity × crowd% × (1 − non-paying%))
- Money per paid ticket = average price × (1 − fee% − HST%) + (bar + merch) ÷ (1 − non-paying%)
- Break-even paid tickets = (total cost − sponsorship − livestream income) ÷ money per paid ticket, rounded up
- Sponsorship needed on a low night = total cost − livestream − low-night paid tickets × money per paid ticket
- Profit = ticket money after fees + bar/merch × crowd + sponsorship + livestream − total cost

## Reference results (defaults, expected crowd = 50% of capacity)

| Venue | Capacity | Sponsorship assumed | Total cost | Break-even paid tickets | Expected crowd | Profit at expected crowd | Sponsorship needed on a low night |
|---|---|---|---|---|---|---|---|
| A.J. LaRue Arena · Haliburton | 1,000 | $18,000 | $30,450 | 171 | 500 | $16,764 | $10,000 |
| North Hastings Community Centre · Bancroft | 800 | $12,000 | $30,954 | 260 | 400 | $4,417 | $14,594 |
| Dr. Hamilton Crawford Memorial Arena · Marmora | 1,000 | $6,000 | $31,164 | 345 | 500 | $4,050 | $10,714 |
| Peterborough Memorial Centre · Peterborough | 2,000 | $10,000 | $48,300 | 525 | 1,000 | $20,128 | $7,400 |
| Slush Puppie Place · Kingston | 2,000 | $6,000 | $52,290 | 634 | 1,000 | $12,138 | $11,390 |
| Tribute Communities Centre · Oshawa | 2,000 | $6,000 | $52,290 | 634 | 1,000 | $12,138 | $11,390 |

Venue notes: Haliburton uses the real 2026 arena bill. Bancroft and Haliburton capacities are estimates. Marmora's 1,000 and $600/day are published. Peterborough, Kingston and Oshawa are set to 2,000 seats with estimated rent: get quotes. Any venue whose expected crowd is under 1,000 is below Cody's minimum.

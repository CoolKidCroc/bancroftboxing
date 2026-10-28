"""Write the social skill's show-campaign.md from the CAMP list in build_marketing.py (single source of truth)."""
import re, sys
s = open("build_marketing.py").read()
a = s.index("CAMP = ["); b = s.index("\n]\n", a) + 2
CAMP = eval(s[a + len("CAMP = "):b])
L = ["# CCCL show campaign (weeks before the show)", "",
     "The same plan as the Campaign tab of the CCCL Marketing Kit. Numbers are weeks before the show; after the show they are days. Count back from the show date once a show is booked. Every post needs Cody's OK (Approvals tab in Constrictor HQ).",
     "CCCL posts go on CCCL's own Facebook and Instagram, then get shared to Bancroft Boxing & Grappling's pages.", "",
     "| Weeks before | Channel | What to do | Graphic | Owner |", "|---|---|---|---|---|"]
for w, ch, what, asset, owner in CAMP:
    when = "Show day" if w == 0 else ("Day after" if w == -1 else (f"{-w} days after" if w < 0 else str(w)))
    L.append(f"| {when} | {ch} | {what} | {asset} | {owner} |")
L += ["", "## Rules for every campaign post",
      "- Real numbers only: 400 tickets and 15 sponsors at the first show are real. Never invent crowd, viewer or follower numbers.",
      "- 'Selling fast' or 'last tables' only when the Ticket Tracker says so.",
      "- Fighters: adults who agreed only; official results only; respect for every fighter, no trash talk.",
      "- Youth: no photos or video without a parent's written OK, never named in press, never in ads.",
      "- Every link in a post is that source's own Eventbrite tracking link (Ticket Sources tab), so we can see what sells.",
      "- Sponsors get only what their package includes; check the Packages and Fulfilment tabs.",
      "- No alcohol logos or drinks on anything youth-facing."]
open(sys.argv[1], "w").write("\n".join(L) + "\n")
print("wrote", sys.argv[1], len(CAMP), "rows")

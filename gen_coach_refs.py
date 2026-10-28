"""Write the bbg-coach skill's references from coaching_data.py and research/coaching.json (single source of truth with the sheet).
usage: python3 gen_coach_refs.py [research/coaching.json]"""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from coaching_data import DRILLS, CLASSES, CONTACT, HOUSE, EQUIPMENT, NO_EQUIPMENT
RES = json.load(open(sys.argv[1])) if len(sys.argv) > 1 and os.path.exists(sys.argv[1]) else None
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skills/bbg-coach/references"); os.makedirs(D, exist_ok=True)

L = ["# BBG classes (Gymdesk schedule and Cody's answers, October 2026)", "",
     "Contact allowed is the most a plan may use. Youngest age is the age every drill must suit, except in a class that splits into age groups: there, a drill above the youngest age is fine as long as younger kids do a drill for their age at the same time. Change these only on Cody's say-so.", "",
     "| Class | When | Minutes | Ages | Youngest | Contact allowed | Focus | Coach | Splits into age groups? |", "|---|---|---|---|---|---|---|---|---|"]
for c in CLASSES:
    L.append(f"| {c[0]} | {c[1]} | {c[2]} | {c[3]} | {c[4]} | {c[5].replace(';', ', ')} | {c[6]} | {c[7]} | {c[8]} |")
L += ["", "Yoga (Tue 6:15-7:00 PM) has its own instructor and isn't planned here.", "",
      "## Equipment at BBG (Cody, October 2026)", "", "Plans only use this gear: " + ", ".join(EQUIPMENT) + ".", "",
      "Not at BBG: " + ", ".join(NO_EQUIPMENT).lower() + ". Slip and bob-and-weave lines use a light rope or string tied between two posts; jump rope is done rope-less."]
open(os.path.join(D, "classes.md"), "w").write("\n".join(L) + "\n")

L = ["# BBG drill library", "", f"{len(DRILLS)} drills, the same as the Drills tab of the BBG Coaching Library. Contact levels: " + " · ".join(CONTACT) + ".", ""]
cur = None
for d in DRILLS:
    if d[1] != cur:
        cur = d[1]; L += [f"## {cur}", ""]
    L += [f"### {d[0]}", f"- Ages: {d[2]} · Level: {d[3]} · Contact: {d[4]} · Equipment: {d[5]} · {d[6]} min",
          f"- How: {d[7]}", f"- Coaching points: {d[8]}", f"- Safety: {d[9]}", ""]
open(os.path.join(D, "drills.md"), "w").write("\n".join(L))

L = ["# Safety for BBG classes", "", "## BBG house rules (Cody's; never loosen these)", ""]
for t, who, rule, src in HOUSE: L.append(f"- **{t}** ({who}): {rule}")
if RES:
    L += ["", "## What the sport bodies and Ontario say", "", f"Researched {RES.get('generated', 'October 2026')}. Confirm anything that matters with the source; this isn't legal or medical advice.", ""]
    topics = {}
    for it in RES.get("items", []): topics.setdefault(it.get("topic", "Other"), []).append(it)
    for t, its in topics.items():
        L += [f"### {t}", ""]
        for it in its:
            conf = "" if it.get("confidence") in ("high", "medium") else " *(low confidence: check)*"
            L.append(f"- ({it.get('age_group', 'all')}) {it.get('guidance', '')}{conf} [{it.get('source_name', 'source')}]({it.get('source_url', '')})")
        L.append("")
    if RES.get("class_notes"):
        L += ["## Notes by class", ""]
        for k, v in RES["class_notes"].items():
            L.append(f"**{k}**"); L += [f"- {x}" for x in v]; L.append("")
    if RES.get("open_questions"):
        L += ["## For Cody to confirm", ""] + [f"- {q}" for q in RES["open_questions"]] + [""]
    if RES.get("answered_oct20"):
        L += ["## Answered by Cody (Oct 20, 2026)", ""] + [f"- {q}" for q in RES["answered_oct20"]] + [""]
open(os.path.join(D, "safety.md"), "w").write("\n".join(L) + "\n")
print("wrote", os.listdir(D))

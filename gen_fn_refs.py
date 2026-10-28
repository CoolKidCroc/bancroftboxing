"""Write cccl-fight-night references from research JSON.
usage: python3 gen_fn_refs.py research/oac_fight_night.json research/media_gyms.json skills/cccl-fight-night/references"""
import json, sys, os
OAC, MG, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

if os.path.exists(OAC):
    d = json.load(open(OAC))
    L = ["# Ontario rules that shape CCCL fight nights", "",
         f"Researched {d.get('generated', 'October 2026')} from official sources where possible. A plain-English summary for planning, not legal advice. "
         "Rules change: confirm anything that matters with the Ontario Athletics Commission (AthleticsComm@ontario.ca) before each show.", ""]
    if d.get("weight_classes"):
        L += ["## Weight classes", "", "| Class | Limit (lb) | Limit (kg) |", "|---|---|---|"]
        for w in d["weight_classes"]:
            L.append(f"| {w.get('name','')} | {w.get('limit_lb','')} | {w.get('limit_kg','')} |")
        L.append("")
    topics = {}
    for it in d.get("items", []):
        topics.setdefault(it.get("topic", "Other"), []).append(it)
    for t, items in topics.items():
        L += [f"## {t}", ""]
        for it in items:
            src = it.get("source_url", "")
            tag = " (secondary source)" if it.get("source_type") == "secondary" else ""
            qte = f" Quote: \"{it['quote']}\"" if it.get("quote") else ""
            L.append(f"- {it.get('rule','')}{qte} Source: {src}{tag}. Confidence: {it.get('confidence','')}.")
        L.append("")
    if d.get("open_questions"):
        L += ["## Ask the Commission", ""] + [f"- {q}" for q in d["open_questions"]] + [""]
    open(os.path.join(OUT, "ontario-rules.md"), "w").write("\n".join(L) + "\n")
    print("ontario-rules.md", len(d.get("items", [])), "items")

if os.path.exists(MG):
    d = json.load(open(MG))
    L = ["# Media and gyms for CCCL (researched October 2026)", "",
         "Research only: nobody has been contacted. Contacts are each outlet's or gym's own published route. Re-check before using. "
         "The same lists are on the Press and Recruitment tabs of the Fight Night Kit.", "", "## Media", ""]
    for m in d.get("media", []):
        L.append(f"- **{m.get('name','')}** ({m.get('type','')}, {m.get('area','')}). Send news: {m.get('how_to_submit','')}. "
                 f"{m.get('notes','') or ''} Source: {m.get('source_url','')}. Confidence: {m.get('confidence','')}.")
    L += ["", "## Gyms that could send fighters", ""]
    for g in d.get("gyms", []):
        L.append(f"- **{g.get('name','')}** ({g.get('town','')}): {g.get('disciplines','')}. Competition team: {g.get('has_competition_team','')}. "
                 f"Contact: {g.get('contact','')}." + (f" Website: {g['website']}." if g.get('website') else "") + (f" {g['notes']}" if g.get('notes') else "") + f" Source: {g.get('source_url','')}. Confidence: {g.get('confidence','')}.")
    if d.get("notes"):
        L += ["", "## Notes", "", d["notes"]]
    open(os.path.join(OUT, "media-and-gyms.md"), "w").write("\n".join(L) + "\n")
    print("media-and-gyms.md", len(d.get("media", [])), "media", len(d.get("gyms", [])), "gyms")

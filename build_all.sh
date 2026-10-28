#!/usr/bin/env bash
# Rebuilds every file in the Bancroft Boxing folder from this source.
# usage: bash build_all.sh OUTDIR          (OUTDIR mirrors the Bancroft Boxing folder layout)
# Needs: Python 3 + openpyxl; Node 18+ with the "docx" and "playwright" packages installed globally
#        (npm i -g docx playwright, then npx playwright install chromium); "npm install" once inside canva/ for the fonts.
# Optional: set RECALC to the xlsx skill's recalc.py to fill in formula values before saving (Google Sheets recalculates on open anyway).
set -e
cd "$(dirname "$0")"
O="${1:?give an output folder, e.g. bash build_all.sh ../out}"
export NODE_PATH="$(npm root -g)"
mkdir -p "$O/Command Centre" "$O/BBG Coaching" "$O/BBG Members & Leads" "$O/CCCL/Sponsorships" "$O/CCCL/Marketing/Canva starter designs" \
  "$O/CCCL/Fight Night" "$O/Handoff/Cheat sheet (save to your phone)" "$O/The Constrictor (Claude skills)" "$O/Second Chance" "$O/Athlete Brand" "$O/Brand"

# ---- sheets
python3 build_predictor.py "$O/CCCL/CCCL Event Profit Predictor.xlsx" research/town_scout.json research/venues_2027.json
python3 build_pipeline.py research/sponsor_prospects.json "$O/CCCL/Sponsorships/CCCL Sponsorship Pipeline.xlsx" research/big_sponsors.json
python3 build_marketing.py "$O/CCCL/Marketing/CCCL Marketing Kit.xlsx" research/marketing_tools.json
python3 build_fightnight.py "$O/CCCL/Fight Night/CCCL Fight Night Kit.xlsx" research/oac_fight_night.json research/media_gyms.json
python3 build_leads.py "$O/BBG Members & Leads/BBG Leads & Trials.xlsx"
python3 build_coaching.py "$O/BBG Coaching/BBG Coaching Library.xlsx" research/coaching.json
python3 build_passport.py "$O/BBG Coaching/BBG Skills Passport.xlsx"
python3 build_funding.py "$O/Command Centre/The Constrictor Funding Tracker.xlsx" research/funding.json
python3 build_hq.py "$O/Command Centre/Constrictor HQ.xlsx"
python3 build_contacts.py "$O/Command Centre/BBG & CCCL Contacts.xlsx" research/venues_2027.json research/town_scout.json research/media_gyms.json
python3 build_money.py "$O/Command Centre/BBG & CCCL Money Summary.xlsx"
python3 build_tests.py "$O/Handoff/The Constrictor Skill Tests.xlsx"
python3 build_second_chance.py "$O/Second Chance/Second Chance Tracker.xlsx" research/second_chance_plan.json
if [ -n "$RECALC" ]; then
  find "$O" -name "*.xlsx" -print0 | while IFS= read -r -d '' f; do python3 "$RECALC" "$f" 300 >/dev/null && echo "recalc ok: $(basename "$f")"; done
fi

# ---- skill reference files made from the research and sheets
python3 gen_scout_refs.py research/venues_2027.json "$O/CCCL/CCCL Event Profit Predictor.xlsx" skills/cccl-event-scout/references
python3 gen_coach_refs.py research/coaching.json
python3 gen_fn_refs.py research/oac_fight_night.json research/media_gyms.json skills/cccl-fight-night/references
python3 gen_campaign_md.py skills/bbg-cccl-social/references/show-campaign.md

# ---- Word documents
L=onepager/cccl-logo-400.png; BL=handoff/bbg-round-400.png
node plan.js "$O/00 START HERE - The Constrictor AI Plan.docx" logo_small.png
node staff.js "$O/BBG & CCCL Staff Playbook.docx"
node install.js "$O/The Constrictor (Claude skills)/How to add The Constrictor skills to Claude.docx"
node handoff/guide.js "$O/Handoff/The Constrictor Handoff Guide.docx" "$BL" "$L"
node workstation.js "$O/Handoff/The Constrictor Workstation Plan.docx" "$BL" "$L"
node templates.js "$O/CCCL/Sponsorships/CCCL Sponsor Pitch Templates.docx"
node templates_fn.js "$O/CCCL/Fight Night/CCCL Fight Night Templates.docx" "$L"
node fighter_pkg.js "$O/CCCL/Fight Night/CCCL Fighter Package.docx" "$L"
node csj.js "$O/Command Centre/CSJ 2027 Application Draft.docx" "$BL"
node bbg_employer.js "$O/Command Centre/BBG First Employee Basics.docx" "$BL"
node second_chance.js "$O/Second Chance/Second Chance Starter Plan.docx" research/second_chance_plan.json
node sc_docs.js "$O/Second Chance/Second Chance Policy Drafts.docx" "$O/Second Chance/Second Chance First Letters.docx" skills/second-chance/references
node athlete.js "$O/Athlete Brand/The Constrictor Athlete Brand Plan.docx" "$BL" "$L" skills/bbg-cccl-social/references/athlete-brand.md

# ---- designs and PDFs (Playwright)
node athlete_designs.js "$O/Athlete Brand"
node canva/gen_designs2.js "$O/CCCL/Marketing/Canva starter designs" "$L"
rm -rf "${O:?}/CCCL/Marketing/Canva starter designs/_bg"   # working backgrounds, not part of the folder
node canva/gen_brandkit.js "$O/Brand/CCCL Brand Kit.pdf" "$L" research/marketing_tools.json
node onepager/pdf.js "$(pwd)/onepager/sponsor-sheet.html" "$O/CCCL/Sponsorships/CCCL 2027 Sponsor Packages.pdf"
node handoff/cheat.js "$O/Handoff/Cheat sheet (save to your phone)" "$L" "$BL"
cp blueprint/constrictor-blueprint.html "$O/The Constrictor AI Blueprint.html"
cp brand/*.png "$O/Brand/"

# ---- skills (one zip each, ready to upload to Claude)
cd skills
for s in */; do s="${s%/}"; rm -f "$O/The Constrictor (Claude skills)/$s.zip"
  zip -r -X -q "$O/The Constrictor (Claude skills)/$s.zip" "$s" -x '*.DS_Store' '*__pycache__*'; echo "zip $s"; done
echo "Before copying into the Bancroft Boxing folder, check the changed files open and look right."

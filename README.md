# The Constrictor AI: build source

Everything that makes the files in the Bancroft Boxing folder: the scripts, the research they use and the 9 skills. **Kyle's working copy. Never upload this folder to Cody's Google Drive.**

Last full rebuild: October 27, 2026 (every file rebuilt cleanly from this folder).

## The easy way: have Claude do it

Open Claude on this PC (the desktop app, so the chat can reach this folder) and say:

> Read README.md in Bancroft Boxing/_Kyle (don't upload)/Build source. Then [say what changed, e.g. "Cody raised drop-in to $25" or "apply Cody's follow-up answers"].

Cody's follow-up answers are saved by the form at https://claude.ai/artifact/3S4n3dbRBz1kt5cCfTdgXL (its database: collection `answers`, document `followups`). Kyle's Finish Plan (one folder up), on the Waiting on Cody tab, says where each answer goes.

Claude should edit the source here, run `build_all.sh` into a fresh output folder, check the files it changed, then copy only those into the Bancroft Boxing folder. Any skill it changed needs re-testing (Handoff / The Constrictor Skill Tests), and Cody re-uploads it: delete the old skill in Claude first, then upload the new zip.

## Running it yourself

You need Python 3 with `openpyxl`, Node 18+ with `docx` and `playwright` installed globally (`npm i -g docx playwright`, then `npx playwright install chromium`), and `npm install` run once inside `canva/` for the fonts. Then:

```
bash build_all.sh ../out
```

`../out` gets the same layout as the Bancroft Boxing folder. Copy across only the files you meant to change. Windows: run it in WSL or Git Bash.

## Which script makes which file

| File in the Bancroft Boxing folder | Made by | Uses |
|---|---|---|
| 00 START HERE - The Constrictor AI Plan.docx | plan.js | — |
| BBG & CCCL Staff Playbook.docx | staff.js | — |
| The Constrictor AI Blueprint.html | blueprint/constrictor-blueprint.html (copied as is; also published as the Blueprint link) | — |
| Athlete Brand: plan + 3 designs | athlete.js, athlete_designs.js | writes skills/bbg-cccl-social/references/athlete-brand.md |
| BBG Coaching Library.xlsx | build_coaching.py | coaching_data.py, research/coaching.json |
| BBG Skills Passport.xlsx | build_passport.py | writes skills/bbg-coach/references/passport.md |
| BBG Leads & Trials.xlsx | build_leads.py | — |
| CCCL Event Profit Predictor.xlsx | build_predictor.py | research/town_scout.json, venues_2027.json |
| CCCL Sponsorship Pipeline.xlsx | build_pipeline.py | research/sponsor_prospects.json, big_sponsors.json |
| CCCL Marketing Kit.xlsx | build_marketing.py | research/marketing_tools.json |
| CCCL Fight Night Kit.xlsx | build_fightnight.py | research/oac_fight_night.json, media_gyms.json |
| Fight Night Templates, Fighter Package (.docx) | templates_fn.js, fighter_pkg.js | — |
| Sponsor Pitch Templates.docx | templates.js | — |
| CCCL 2027 Sponsor Packages.pdf | onepager/sponsor-sheet.html via onepager/pdf.js | — |
| Canva starter designs (16), CCCL Brand Kit.pdf | canva/gen_designs2.js, canva/gen_brandkit.js | canva/thumbs, research/marketing_tools.json |
| Constrictor HQ.xlsx | build_hq.py | reads skills/bbg-cccl-social/references/athlete-brand.md |
| BBG & CCCL Contacts.xlsx | build_contacts.py | research/venues_2027.json, town_scout.json, media_gyms.json |
| BBG & CCCL Money Summary.xlsx | build_money.py | — |
| The Constrictor Funding Tracker.xlsx | build_funding.py | research/funding.json (also writes the funding skill's references) |
| CSJ 2027 Application Draft, BBG First Employee Basics | csj.js, bbg_employer.js | research/first_employee_facts.json |
| Second Chance: Starter Plan, Tracker, Policy Drafts, First Letters | second_chance.js, build_second_chance.py, sc_docs.js | research/second_chance_plan.json (also write the second-chance skill's references) |
| Handoff Guide, Workstation Plan, Skill Tests, cheat sheet | handoff/guide.js, workstation.js, build_tests.py, handoff/cheat.js | — |
| How to add The Constrictor skills (.docx) | install.js | — |
| The 9 skill zips | skills/* (zipped by build_all.sh) | some reference files are written by the scripts above: gen_scout_refs.py, gen_coach_refs.py, gen_fn_refs.py, gen_campaign_md.py |

`docstyle.js` is the shared Word look. The round-1 and round-2 question forms for Cody are in `questions/`. `research/` holds the sourced research behind every number; the `.md` files there give the sources and dates.

## Common changes

| What changed | Edit | Then |
|---|---|---|
| Prices, class times, ages | skills/bbg-front-desk/references/gym-facts.md AND the copy in skills/bbg-cccl-social/references; CLASSES in coaching_data.py; FACTS in staff.js | Rebuild; re-test the front-desk, social and coach skills |
| Sponsor packages | Packages in build_pipeline.py; skills/cccl-sponsor-desk (packages.md, cccl-facts.md, pitch-templates.md, the SKILL.md line); templates.js; onepager/sponsor-sheet.html | Rebuild; re-test sponsor desk |
| A show is booked | Nothing in the source. Cody copies the Fight Night Kit and Marketing Kit in Drive and fills Show Setup | — |
| Cody answers a question | The file named in Kyle's Finish Plan (Waiting on Cody tab) | Rebuild that file |
| New research | The matching research/*.json, keeping the source link and date | Rebuild what uses it |

## Rules every change must keep

- Drafts only: Claude never sends, posts or submits. People send.
- First names only in sheets. No health details, kids' last names, card numbers or personal phone numbers. Medical status is Yes / No / Pending only.
- Kids' photos only with a parent's OK. Nobody under 18 in an ad or named in a press release.
- Second Chance is kept separate: its own Gmail, Drive folder and bank account. Never record offences, convictions, health details or addresses; first name plus a code (SC-001) only.
- Only "Paid" counts as sponsor revenue. Skip cannabis and vape sponsors; flag alcohol, gambling and energy drinks, and keep them away from youth.
- No prize to a fighter worth $300 or more in total.
- Cody and the officials decide every bout.
- Real numbers only: no made-up figures, and say where every number came from.
- Cody's personal email never goes in a shared file.
- Research, not legal, tax or medical advice.
- Plain, simple language: Cody and his staff are the readers.

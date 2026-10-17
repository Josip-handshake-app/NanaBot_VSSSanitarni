# Zagreb VSS job alerts

Daily check for higher-education (`VSS`) public jobs in Zagreb, narrowed to sanitary engineering on every source except the state portal. The script reads the city boards, the state portal, Zagrebački holding, and the national employers listed below, emails new matches through [Resend](https://resend.com), and records what it already sent in `sent_jobs.json`.

## Sources

- City offices: https://zagreb.hr/objavljeni-natjecaji-oglasi/6501
- Culture institutions: https://zagreb.hr/javni-natjecaji-za-zaposljavanje-u-ustanovama-u-ku/202452
- Kindergartens: https://vrtici.zagreb.hr/natjecaji-za-zaposljavanje/122
- State portal: https://selekcija.gov.hr/natjecaji/objavljeni-natjecaji
- Zagrebački holding: https://www.zgh.hr/karijere/javni-natjecaji-za-zaposljavanje/5699
- HEP grupa: https://eprijave.hep.hr/
- Hrvatske ceste: https://hrvatske-ceste.hr/hr/stranice/o-nama/zaposljavanja
- Hrvatske vode: https://www.voda.hr/hr/zaposljavanje
- HŽ Putnički prijevoz: https://www.hzpp.hr/hr/tvrtka-2/zaposljavanje
- HŽ Infrastruktura: https://www.hzinfra.hr/category/natjecaj-za-zaposljavanje/
- HŽ Cargo: https://www.hzcargo.hr/zaposljavanje/
- Hrvatske autoceste: https://www.hac.hr/o-nama/zaposljavanje
- Hrvatske šume: https://www.hrsume.hr/vijesti/

On selekcija.gov.hr a job is kept when professional qualification contains `VSS` and the workplace contains Zagreb. On every other source the notice must ask for `VSS`, `visoka stručna sprema`, `magistar`, or `diplomski studij`, name a sanitary engineer (`sanitarn` together with `inženjer` or `inzenjer`), and place the job in Zagreb. The first successful run emails every current match. Later runs email only ids that are not yet in `sent_jobs.json`.

## GitHub Actions

Add these repository secrets:

- `RESEND_API_KEY` — Resend API key
- `RECEIVER_EMAIL` — address that should receive the alert
- `RESEND_FROM_EMAIL` — verified sender, for example `NanaBot <jobs@yourdomain.com>`

The workflow [`.github/workflows/daily-jobs.yml`](.github/workflows/daily-jobs.yml) runs at 11:37 and 16:37 Europe/Zagreb. If the first run already saved the new ids, the second run does not send another email.

After a successful send, the workflow commits `sent_jobs.json` as `github-actions[bot]`. A failed send leaves the file unchanged, so the same jobs are tried again next time.

You can start a run manually from the Actions tab with **Run workflow**.

## Local run

```powershell
python -m pip install -r requirements.txt
$env:RESEND_API_KEY = "re_..."
$env:RECEIVER_EMAIL = "sister@example.com"
$env:RESEND_FROM_EMAIL = "NanaBot <jobs@yourdomain.com>"
python scrape_jobs.py
```

Copy [`.env.example`](.env.example) only as a reminder. The script reads `os.environ`. It does not load a `.env` file.

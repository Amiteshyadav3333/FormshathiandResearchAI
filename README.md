# FormSaathi Desktop Workspace

Local-first desktop workspace for shop sales reports, source-linked online research, and reviewed job-form entries. It includes an animated helper and a Hindi conversational panel.

## Start

```bash
python3 -m venv venv
source venv/bin/activate          # Windows: venv\\Scripts\\activate
pip install -r requirements.txt
python server.py
```

Open <http://127.0.0.1:8000> on the same computer. Do not expose Flask's development server to the public internet.

## Desktop app

- **macOS:** Open `dist/FormSaathi.app` after building with `./build_mac.sh`.
- **Windows:** Run `build_windows.bat` on a Windows PC; it creates `dist\FormSaathi\FormSaathi.exe`.
- The `.github/workflows/build-desktop.yml` workflow can build distributable app folders for macOS and Windows from a tagged release or manual workflow run.

The Windows executable must be compiled on Windows (or using the included Windows build workflow). Internet access is required for online research. The rest of the workspace runs locally on the computer.

## Sales workbook

1. Upload a CSV, TSV, XLSX, XLS or JSON export. The app looks for product/item, quantity, unit price or total sales, date, and unit/total cost columns.
2. Review the on-screen product summary, totals and skipped rows. Choose the currency for display and add a shop/company name if desired.
3. Download a formatted Excel workbook with **Overview**, **Product Summary**, **Sales Detail**, and **Monthly Trend** sheets.

`Product` and `Quantity` are required. Sales are calculated as quantity × unit price when no total is present. Profit is estimated only when cost data exists. Returns, tax, fees, and overhead are not inferred. Only the first sheet of an uploaded Excel file is read; up to 5,000 uploaded rows are processed.

## Jobs and forms

- Upload a resume in PDF, DOCX or TXT, or a tabular data file for other forms.
- Add the online form URL and open it in another tab. Use the field-reader bookmarklet to copy its field names into FormSaathi.
- Check and correct each suggested match. Uncertain fields stay empty; ask the Hindi assistant what a field means.
- Review the target HTTPS site and selected values, then confirm the destination checkbox. Generate the generic fill bookmarklet and copy the approved values separately. At the matching site, run the bookmarklet and paste the copied data into its prompt. The bookmarklet contains no personal values and refuses to run on a different site. Verify each entry, handle validation messages, and submit the form yourself.

The app does not log into job sites, submit applications, answer screening questions from assumptions, bypass CAPTCHAs, or fill passwords/OTPs/payment fields. It only prefills reviewed values into the host explicitly approved by the user; it never submits the application. Resume extraction is best-effort; check the extracted name, email, phone, and skills.

## Hindi assistant

The in-app assistant answers basic sales questions from the current report, explains a few common job fields, and asks for more detail when it cannot tell which task you mean. It uses local, rule-based guidance and does not send records or resumes to an external AI service. Online research uses Bing search and returns source titles, snippets, links, and search time; download these into a source workbook and verify each linked page before using its claims. Search snippets are not verified datasets, and the app does not currently create a complete, validated MLA asset list from search alone.

`agent.py` is a separate optional voice-agent script and is not started by the desktop app. When run, it sends microphone audio/transcripts to its configured LiveKit/Deepgram/Inference/Synthesia services (and to the configured Spring endpoint when enabled); it also prints voice turns to its console. Keep `.env.local` private and avoid sharing sensitive information through that optional voice path.

## Privacy and limitations

- Uploaded documents are processed in memory and are not saved to disk by the app. Data stays in the active browser page until it is closed or refreshed.
- Search requests send the query to Bing. Do not include sensitive personal information in search queries. Returned snippets and linked pages can be incomplete or out of date; check source dates and original documents.
- The local app binds only to loopback and checks the browser request origin. Uploaded CSV/Excel/Word archives, PDF page counts, and row counts have size/complexity caps; generated Excel/CSV escapes formula-like text.
- Uploads are capped at 12 MB per request; tabular files are capped at 5,000 rows, compressed Office archives at 100 MB expanded size, and resume PDFs at 100 pages. These limits protect a single-computer app from accidental or hostile oversized files.
- Form fill uses a browser bookmarklet and may not work on custom controls, shadow DOM pages, or sites that block bookmarklets. Complete those fields manually.
- Field matching is based on column names and labels. It cannot verify source data, job eligibility, or application accuracy.
- This is a single-user local prototype. Use authentication, a production WSGI server, HTTPS, and security review before any shared or public deployment.

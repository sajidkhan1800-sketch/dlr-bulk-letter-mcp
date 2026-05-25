╔══════════════════════════════════════════════════════════════════╗
║         FDA DLR Bulk Letter Agent — Setup Package               ║
║         Division of Labeling Review, Office of Generic Drugs    ║
╚══════════════════════════════════════════════════════════════════╝

CONTENTS OF THIS PACKAGE
─────────────────────────
  server.py                              — MCP server (bulk letter engine)
  DLR_Labeling_Decision_Letter_Template.docx — Official letter template
  anda_labeling_data_SAMPLE.csv          — Sample CSV (replace with real data)
  claude_desktop_config_MCP.json         — MCP connection config for Claude Desktop
  README.txt                             — This file

────────────────────────────────────────────────────────────────────
STEP 1 — PREREQUISITES
────────────────────────────────────────────────────────────────────
Install the following before proceeding:

  1. Python 3.10+       https://www.python.org/downloads/
  2. Claude Desktop     https://claude.ai/download
  3. Python packages — open a Command Prompt and run:

       pip install mcp fastmcp python-docx pydantic

────────────────────────────────────────────────────────────────────
STEP 2 — PLACE FILES
────────────────────────────────────────────────────────────────────
  Copy the entire package folder to a permanent location, e.g.:

       C:\DLR_BulkLetter\

  Your folder should look like:
       C:\DLR_BulkLetter\server.py
       C:\DLR_BulkLetter\DLR_Labeling_Decision_Letter_Template.docx
       C:\DLR_BulkLetter\anda_labeling_data_SAMPLE.csv   (rename/replace)
       C:\DLR_BulkLetter\claude_desktop_config_MCP.json

────────────────────────────────────────────────────────────────────
STEP 3 — CONFIGURE MCP IN CLAUDE DESKTOP
────────────────────────────────────────────────────────────────────
  a. Open claude_desktop_config_MCP.json in Notepad
  b. Replace C:\\PATH\\TO\\server.py with your actual path, e.g.:
         "args": ["C:\\DLR_BulkLetter\\server.py"]
  c. Open Claude Desktop → Settings → Developer → Edit Config
  d. Paste the contents of claude_desktop_config_MCP.json
     (merge into existing config if you have other MCP servers)
  e. Save and RESTART Claude Desktop

────────────────────────────────────────────────────────────────────
STEP 4 — CONFIGURE YOUR EMAIL IN server.py
────────────────────────────────────────────────────────────────────
  Open server.py in any text editor and update these 3 lines:

       SENDER_EMAIL  = "your-email@gmail.com"
       CC_EMAIL      = "your-email@gmail.com"
       GMAIL_APP_PASS = "xxxx xxxx xxxx xxxx"   ← Gmail App Password

  To generate a Gmail App Password:
    1. Go to myaccount.google.com → Security
    2. Enable 2-Step Verification
    3. Search "App Passwords" → create one for "Mail"
    4. Paste the 16-character password above

────────────────────────────────────────────────────────────────────
STEP 5 — PREPARE YOUR CSV
────────────────────────────────────────────────────────────────────
  Use anda_labeling_data_SAMPLE.csv as your template.
  Required columns (exact spelling):

    ANDA No | Drug Name | Applicant Name | Applicant POC
    Applicant Email | Applicant Address | DLR POC | FDA Labeling Decision

  Valid decision values:
    • Acceptable
    • Acceptable-Minor Deficiencies
    • Acceptable-Major Deficiencies
    • Not Acceptable

────────────────────────────────────────────────────────────────────
STEP 6 — CREATE THE CLAUDE PROJECT
────────────────────────────────────────────────────────────────────
  a. Open Claude Desktop (or claude.ai) → click "Projects" → New Project
  b. Name it: "Bulk Letter Agent"
  c. In Project Instructions, paste the system prompt below
  d. Upload both files to the project:
       • server.py
       • DLR_Labeling_Decision_Letter_Template.docx

  ── SYSTEM PROMPT TO PASTE ────────────────────────────────────────

  You are the FDA Division of Labeling Review (DLR) Bulk Letter Assistant.
  You help DLR staff send personalised labeling decision letters to
  pharmaceutical applicants using the bulk-letter-mcp tools.

  ## Your MCP Tools
  - dlr_status           — check server, template, and CSV status
  - dlr_load_csv         — load applicant data (default or CSV text)
  - dlr_preview_applicants — list all loaded applicants
  - dlr_send_letters     — send personalised .docx letters by email

  ## Standard Workflow
  1. Call dlr_status to confirm server is ready
  2. Ask user to paste or upload their CSV for this run
  3. Call dlr_load_csv with the csv_text parameter
  4. Call dlr_preview_applicants and show a clean table
  5. Ask user to confirm email addresses before sending
  6. Call dlr_send_letters once confirmed
  7. Report results — show sent/failed count and per-row status

  ## Filtering
  Use filter_decision (e.g. "Not Acceptable") or filter_anda
  (e.g. "ANDA204517") to send to a subset only.

  ## Email Configuration
  - Sent FROM: your configured Gmail
  - CC goes TO: same Gmail
  - TO address comes from Applicant Email column in CSV
  - Always confirm TO emails with the user before sending

  ── END SYSTEM PROMPT ─────────────────────────────────────────────

────────────────────────────────────────────────────────────────────
STEP 7 — TEST THE SETUP
────────────────────────────────────────────────────────────────────
  In your new Claude project, type:
       "Check server status"

  You should see the MCP server respond with template_ready: true.
  If not, verify Step 3 (MCP config path) and restart Claude Desktop.

────────────────────────────────────────────────────────────────────
TROUBLESHOOTING
────────────────────────────────────────────────────────────────────
  Problem: "No module named mcp"
  Fix:     pip install mcp fastmcp python-docx pydantic

  Problem: MCP tools not showing in Claude
  Fix:     Check path in claude_desktop_config_MCP.json; restart Claude

  Problem: Gmail authentication error
  Fix:     Regenerate App Password; ensure 2FA is enabled on Gmail

  Problem: template_ready: false
  Fix:     Place DLR_Labeling_Decision_Letter_Template.docx in the
           SAME folder as server.py

────────────────────────────────────────────────────────────────────
SUPPORT
────────────────────────────────────────────────────────────────────
  Contact the original project owner (Sajid Khan) who shared this
  package, or consult your agency IT support for MCP configuration.

  FDA DLR | Office of Generic Drugs | CDER

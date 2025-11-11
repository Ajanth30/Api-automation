# Dynamic API Test Generator

Automates API regression testing from Excel definitions by:

- Reading test cases from a shared Excel file
- Generating a Postman collection on the fly
- Executing the collection with Newman
- Writing back pass/fail status and producing reports

Ideal for use in CI (e.g., Jenkins) so every build pulls the latest GitLab changes, regenerates tests, and emails results.

---

## Requirements

- Python 3.10 or newer
- Node.js 16+ (for Newman)
- `npm install -g newman`
- Git client (for Jenkins checkout/pull)
- Access to the Excel workbook referenced in `services_config.yaml`
- Optional: SMTP account for email notifications

Install Python dependencies:

```bash
pip install -r requirements.txt
```

---

## Project Structure

```
.
├── src/
│   ├── main.py                  # Entry point
│   ├── excel_postman_generator.py
│   ├── newman_runner.py
│   ├── auth_client.py
│   └── emailer.py
├── services_config.yaml         # Environment + email configuration
├── requirements.txt
└── README.md
```

Sample output artifacts (`*_results.xlsx`, `*_postman_collection.json`, `newman_results.json`) are produced during runs and should be git-ignored.

---

## Configuration (`services_config.yaml`)

Key settings:

- `excel_path`: Full path to the master test workbook.
- `collection_name`: Friendly name for the generated Postman collection.
- `gateway_base_url`: Optional override for all request hosts.
- `auth`: Optional block to fetch a bearer token before running tests.
  - `base_url`, `endpoint`, `method`, `headers`, `body`
  - `token_path`: Dot-path to extract the token from the JSON response
  - `header_name` / `header_prefix`: How to inject the token in requests
- `email`: Optional SMTP settings to send results.
  - `recipients`: List of email addresses
  - `from`: Sender address
  - `smtp.host`, `smtp.port`, `smtp.username`, `smtp.password`
  - `smtp.use_tls` / `smtp.use_ssl`

If Jenkins injects credentials as environment variables, reference them in YAML using `"${ENV_VAR_NAME}"`.

---

## Running Locally

```bash
python -m src.main
```

The script will:

- Validate/authenticate (if configured)
- Generate `<collection_name>_postman_collection.json`
- Execute Newman via `newman_runner.py`
- Update a copy of the Excel file with `ActualStatus`/`Status`
- Print the paths of generated artifacts

---

## Jenkins Integration Overview

1. Install prerequisites on the Jenkins node (Python, Git, Node.js, Newman).
2. Checkout the GitLab repository in the job.
3. Run `pip install -r requirements.txt`.
4. Execute `python -m src.main` (optionally inside a virtualenv).
5. Use Jenkins credentials binding so `services_config.yaml` can access SMTP or auth secrets.
6. Archive generated reports or publish them via email.

Refer to Jenkins job examples in your CI folder or ask the infra team for the shared pipeline if available.

---

## Email Notifications

With the `email` section configured, the `emailer.send_results_email` helper attaches:

- Generated Postman collection
- Excel results workbook (with pass/fail coloring)

Console logs include `✉️ Results email sent successfully.` on success. Failures print the underlying SMTP exception.

---

## Troubleshooting

- **Authentication errors**: ensure `auth.base_url`, `endpoint`, and `token_path` match the API. Inspect console logs for `AuthError`.
- **Newman not found**: check that `newman` (or `newman.cmd` on Windows) is on the PATH for the user running Jenkins.
- **Excel locked**: the script writes to `<excel_path>_results.xlsx`. Make sure Excel isn’t holding the file open.
- **Emails blocked**: confirm firewall rules allow SMTP traffic and credentials are correct.

For verbose Newman diagnostics, temporarily edit `src/newman_runner.py` to add `--verbose` to the CLI command.

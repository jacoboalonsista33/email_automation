# Email Automation

This repository provides a small script to send an initial email and up to three follow-ups based on an Excel spreadsheet.

## Usage
1. Prepare an Excel file called `contacts.xlsx` in the repository root with the following columns:
   - `Nombre`: recipient name.
   - `Empresa`: company name.
   - `Email`: email address.

2. Ensure you have Python installed with the packages from `requirements.txt`.

3. Define the SMTP credentials as environment variables:
   - `SMTP_SERVER`
   - `SMTP_PORT`
   - `SMTP_USER`
   - `SMTP_PASSWORD`
   - `FROM_ADDRESS`

4. Run the script:

```bash
python send_emails.py
```

The script sends the first email with `presentation.pdf` attached and then sends three follow-ups every five minutes.

# Twilio SMS Bulk Messaging Application

Send customized SMS messages to a group of recipients from Excel using Twilio Lookup validation and a message template.

A Python application that sends customized SMS messages to recipients using the Twilio Programmable Messaging API. Recipient data is read from an Excel file, phone numbers are validated using the Twilio Lookup service, and results are written back to the Excel file.

## Features

✅ **Bulk SMS Sending** - Send messages to multiple recipients from an Excel file
✅ **Phone Number Validation** - Validates phone numbers using Twilio Lookup API
✅ **Template Support** - Customize message templates with variable substitution
✅ **Error Handling** - Detailed error codes for troubleshooting
✅ **Excel Integration** - Read recipients and write results to Excel files
✅ **Comprehensive Logging** - Real-time progress monitoring

## Project Structure

```
Twilio_Sending_SMS/
├── main.py                    # Main application entry point
├── config.py                  # Configuration loader
├── message_template.txt       # SMS message template
├── Sample Data.xlsx           # Recipients data file (local only)
├── Sample Data.xlsx.sample    # Example recipient workbook for publishing
├── .env                       # Twilio credentials (do not commit)
├── .env.example               # Template for .env file
├── requirements.txt           # Python dependencies
├── REQUIREMENTS.md            # Detailed requirements documentation
└── README.md                  # This file
```

## Quick Start

1. Open the project folder:
   ```bash
   cd /Users/cdeng/Documents/Projects/Twilio_Sending_SMS
   ```
2. Activate the virtual environment:
   ```bash
   source venv/bin/activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the app:
   ```bash
   python main.py
   ```

Optional: run in dry-run mode to validate without sending SMS:
```bash
cd /Users/cdeng/Documents/Projects/Twilio_Sending_SMS
source venv/bin/activate
python main.py --dry-run
```

Use `Sample Data.xlsx.sample` as the example recipient workbook, and keep the real `Sample Data.xlsx` file local and untracked, just like `.env`.

## Prerequisites

- Python 3.8 or higher
- A Twilio account with:
  - Account SID
  - Auth Token
  - A verified phone number to send from
- Active Twilio Lookup service (for phone validation)

## Installation

### 1. Clone/Download the Project
```bash
cd /Users/cdeng/Documents/Projects/Twilio_Sending_SMS
```

### 2. Create Virtual Environment (Recommended)
```bash
python3 -m venv venv
source venv/bin/activate  # On macOS/Linux
# or
venv\Scripts\activate  # On Windows
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Twilio Credentials
The `.env` file should already contain your Twilio credentials:
```
TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxx
TWILIO_AUTH_TOKEN=xxxxxxxxxxxxxx
FROM_NUMBER=+1XXXXXXXXXX
```

## Preparing Your Data

### Excel File Format

Your Excel file (`Sample Data.xlsx`) should have the following structure:

| Column 1 | Column 2 | Column 3 | Column 4 |
|----------|----------|----------|----------|
| Name | Phone Number | Product Info | Error Code |
| John Doe | 5551234567 | Premium Plan | |
| Jane Smith | +16505551234 | Basic Plan | |
| Bob Johnson | 14155558901 | Enterprise Plan | |

**Column Requirements:**
- **Column 1 (Name)**: Recipient's name (required)
- **Column 2 (Phone Number)**: Phone number in various formats:
  - 10 digits: `5551234567` (will auto-add leading 1)
  - 11 digits starting with 1: `15551234567`
  - With formatting: `(555) 123-4567` (formatting removed)
  - With +: `+15551234567`
- **Column 3 (Product Info)**: Product or service name (required)
- **Column 4 (Error Code)**: Leave empty - will be filled by the application

### Message Template

Edit `message_template.txt` to customize your message:

```
Hello {name}, your {product_info} is expired, please update your payment information.
```

**Available Variables:**
- `{name}` - Recipient's name (from Column 1)
- `{product_info}` - Product info (from Column 3)

Example templates:
```
Hi {name}, reminder: your {product_info} subscription expires soon!

Dear {name}, your {product_info} requires immediate attention.

{name}, it's time to renew your {product_info}. Click here to update.
```

## Running the Application

### Start the environment
```bash
cd /Users/cdeng/Documents/Projects/Twilio_Sending_SMS
source venv/bin/activate
```

### Basic Execution
```bash
python main.py
```

### Dry Run Mode
```bash
python main.py --dry-run
```
This validates Excel loading, template rendering, and phone lookup without sending SMS or updating the file.

### Custom File Paths
```bash
python main.py --recipients-file "my_contacts.xlsx" --template-file "custom_template.txt"
```
These flags override the `RECIPIENTS_FILE` and `TEMPLATE_FILE` values from `.env`.

### Override From Number
```bash
python main.py --from-number "+15551234567"
```
This flag overrides the `FROM_NUMBER` value from `.env`.

### Full Example
```bash
python main.py --dry-run --recipients-file "my_contacts.xlsx" --template-file "custom_template.txt" --from-number "+15551234567"
```

### Example Output
```
2026-06-04 10:30:45,123 - INFO - Starting Twilio SMS Bulk Messaging Application
2026-06-04 10:30:45,125 - INFO - From Number: +16504886006
2026-06-04 10:30:45,126 - INFO - Recipients File: Sample Data.xlsx
2026-06-04 10:30:45,200 - INFO - Read 3 recipients from Excel file

2026-06-04 10:30:45,300 - INFO - Processing: John Doe (5551234567)
2026-06-04 10:30:45,800 - INFO - Phone validated: +15551234567
2026-06-04 10:30:46,500 - INFO - ✓ Message sent to John Doe
...

==================================================
EXECUTION SUMMARY
==================================================
Total Recipients:     3
Successfully Sent:    2
Failed:               1
  - Invalid Phone:    0
  - Lookup Failed:    0
  - Send Failed:      1
==================================================
```

## Error Codes

After running, Column 4 of your Excel file will contain one of these error codes:

| Code | Meaning | Action |
|------|---------|--------|
| `SUCCESS` | Message sent successfully | ✓ Complete |
| `INVALID_PHONE` | Phone number format is invalid | Check phone number format |
| `LOOKUP_FAILED` | Twilio Lookup service error | Verify number or check Twilio Lookup status |
| `SEND_FAILED` | Twilio API failed to send message | Check Twilio account status/balance |
| `MISSING_DATA` | Required data (name or phone) is missing | Fill in missing fields in Excel |
| `INVALID_TEMPLATE` | Message template substitution failed | Check template variables |

## Troubleshooting

### "TWILIO_ACCOUNT_SID not found in .env file"
- Verify `.env` file exists in the project directory
- Check that credentials are in the correct format
- Ensure no extra spaces or quotes around values

### "Recipients file not found"
- Verify `Sample Data.xlsx` exists in the project directory
- Check file name spelling (case-sensitive on some systems)

### Phone number validation failures
- Ensure phone numbers are in valid US format (10 or 11 digits)
- Check that Twilio Lookup service is enabled on your account
- Verify numbers are not blocked or invalid

### "Send failed" errors
- Check Twilio account balance
- Verify FROM_NUMBER is a verified Twilio number
- Check that recipients are in a country where SMS is allowed
- Review Twilio logs for specific error details

### Application hangs or slow
- Twilio Lookup API adds latency - this is normal
- Check internet connection
- Verify Twilio account isn't rate-limited

## Advanced Usage

### Updating the Template
Edit `message_template.txt` with your custom message:
```bash
# macOS/Linux
nano message_template.txt

# Or use any text editor
```

### Customizing Configuration
You can modify configuration without editing code:
```bash
# Edit .env to change file paths
TEMPLATE_FILE=custom_template.txt
RECIPIENTS_FILE=my_contacts.xlsx
```

### Running Multiple Times
You can safely run the application multiple times. The Excel file will be updated each time with the latest status.

## Important Notes

⚠️ **Security**
- Never commit `.env` file to version control
- Keep `TWILIO_AUTH_TOKEN` confidential
- Use `.env.example` as a template reference

⚠️ **Cost Implications**
- Each SMS sent incurs a charge based on your Twilio plan
- Phone number validation (Lookup) may also have associated costs
- Test with a small number of recipients first

⚠️ **Rate Limiting**
- Twilio may rate-limit bulk sends
- The application processes one recipient at a time by default
- For large volumes (1000+), consider implementing batching

## Support & Troubleshooting

For more information:
- [Twilio Documentation](https://www.twilio.com/docs)
- [Twilio Lookup API](https://www.twilio.com/docs/lookup)
- [Twilio Python SDK](https://www.twilio.com/docs/libraries/python)

Check [REQUIREMENTS.md](REQUIREMENTS.md) for detailed technical requirements and future enhancements.

## Version

**Version**: 1.0  
**Last Updated**: June 4, 2026

---

**Happy messaging!** 🚀

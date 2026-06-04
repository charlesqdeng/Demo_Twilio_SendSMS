# Twilio SMS Bulk Messaging Application - Requirements

## Project Overview
A local Python application that sends SMS messages to a group of recipients using the Twilio Programmable Messaging API. The application reads recipient data from an Excel file, validates phone numbers, and sends personalized SMS messages based on a template.

---

## Functional Requirements

### 1. Configuration & Authentication
- **Source**: Read Twilio account credentials from `.env` file
- **Credentials Required**:
  - `TWILIO_ACCOUNT_SID`: Twilio account identifier
  - `TWILIO_AUTH_TOKEN`: Twilio authentication token
  - `FROM_NUMBER`: Twilio phone number to send SMS from (format: +1XXXXXXXXXX)

### 2. Recipient Data
- **Source**: Local Excel file (`Sample Data.xlsx`)
- **Required Columns**:
  - Column 1: Recipient Name
  - Column 2: Phone Number
  - Column 3: Product Info
  - Column 4: Error Code (output field - populated by application)

### 3. Message Template
- **Source**: Local template file
- **Default Template**: "Hello {name}, your {product info} is expired, please update your payment information"
- **Variables Supported**:
  - `{name}`: Recipient's name
  - `{product_info}`: Product information for the recipient

### 4. Phone Number Validation
- **Service**: Twilio Lookup API
- **Validation Rules**:
  - Valid formats:
    - 11 digits starting with 1 (e.g., 1XXXXXXXXXX)
    - 10 digits (application will automatically add leading 1)
  - Invalid formats are rejected with error code written to Excel

### 5. SMS Sending
- **Service**: Twilio Programmable Messaging API
- **Process**:
  - Read each recipient record
  - Validate phone number
  - Substitute template variables with recipient data
  - Send SMS message
  - Log result (success or error code) to Excel file (Column 4)

### 6. Error Handling & Reporting
- **Error Codes Written to Excel Column 4**:
  - `SUCCESS`: Message sent successfully
  - `INVALID_PHONE`: Phone number format validation failed
  - `LOOKUP_FAILED`: Twilio Lookup service returned an error
  - `SEND_FAILED`: Twilio API failed to send the message
  - `MISSING_DATA`: Required recipient data is missing
  - `INVALID_TEMPLATE`: Template substitution failed

### 7. Output
- **Update Excel File**: Error codes and status written to Column 4 of each row
- **Console Logging**: Display progress and summary of sent/failed messages

---

## Technical Requirements

### Technology Stack
- **Language**: Python 3.8+
- **SMS Service**: Twilio API
- **Data Format**: Excel (.xlsx)
- **Configuration**: Environment variables (.env)

### Python Dependencies
- `twilio`: Twilio Python SDK
- `openpyxl`: Excel file reading and writing
- `python-dotenv`: Environment variable management

### File Structure
```
Twilio_Sending_SMS/
├── .env                          # Twilio credentials (not versioned)
├── .env.example                  # Template for environment variables
├── REQUIREMENTS.md               # This file
├── message_template.txt          # SMS message template
├── Sample Data.xlsx              # Recipients data
├── main.py                       # Main application
├── requirements.txt              # Python dependencies
└── config.py                     # Configuration loader
```

---

## Configuration Files

### .env Format
```
TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxx
TWILIO_AUTH_TOKEN=xxxxxxxxxxxxxx
FROM_NUMBER=+1XXXXXXXXXX
TEMPLATE_FILE=message_template.txt
RECIPIENTS_FILE=Sample Data.xlsx
```

### message_template.txt Format
```
Hello {name}, your {product_info} is expired, please update your payment information.
```

---

## Process Flow

1. **Initialization**
   - Load Twilio credentials from `.env`
   - Load message template from file
   - Open Excel file with recipient data

2. **For Each Recipient**
   - Validate phone number using Twilio Lookup API
   - Substitute template variables
   - Send SMS via Twilio Programmable Messaging API
   - Record result/error code in Excel Column 4
   - Log to console

3. **Completion**
   - Save updated Excel file
   - Display summary (total sent, failed, errors)

---

## Error Handling Strategy

- **Phone Validation**: Use Twilio Lookup to verify format and detect invalid numbers
- **Graceful Failures**: Continue processing remaining recipients if one fails
- **Error Codes**: Record specific error codes for troubleshooting
- **Logging**: Maintain console output for real-time monitoring

---

## Future Enhancements (To Be Considered)

- [ ] Support for multiple message templates
- [ ] Scheduling messages for later delivery
- [ ] Retry mechanism for failed sends
- [ ] Detailed error logs with timestamps
- [ ] Batch processing with progress reporting
- [ ] Support for additional recipient data fields
- [ ] Delivery status tracking
- [ ] Message personalization options

---

## Security Notes

- **Never commit `.env` file** - Always use `.env.example` as template
- **Protect credentials** - Keep TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN secure
- **Excel file handling** - Be cautious with recipient phone number data

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-06-04 | Initial requirements document |


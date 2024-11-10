# Outlook-phishing-detection
VBA code for analyzing Outlook emails for phishing indicators, such as suspicious attachments, headers, and URLs.

## Usage
1. Download the email as a `.msg` file.
2. Run the `AnalyzeDownloadedEmail` function in Outlook VBA, passing the file path of the `.msg` file.

## Code Overview
- **AnalyzeDownloadedEmail**: Main function to load and analyze an email.
- **GetEmailHeaders**: Extracts headers from the email.
- **CheckEmailBodyForPhishingLinks**: Scans for suspicious URLs.
- **CheckForSuspiciousAttachments**: Checks attachments for harmful file types.

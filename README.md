# 📅 Outlook Free Slot Generator

A lightweight tool to analyze your Outlook calendars (supporting multiple accounts) and automatically generate a summary of your availability. Perfect for scheduling meetings without sharing your full calendar access.

The tool calculates free slots based on your working hours, lunch breaks, and specific preferences defined in a configuration file.

## 🚀 Key Features
* **Multi-Calendar Support**: Reads availability from multiple configured email accounts simultaneously.
* **Dual Output Formats**: Generates availability in both a **styled HTML table** and a **plain text** format.
* **Fully Customizable**: Configure working hours, colors, translations, and duration slots via a simple JSON file.
* **One-Click Execution**: Includes `.bat` files to bypass PowerShell execution policies effortlessly—no admin rights required.

## 📁 Project Structure
* `config.json`: The central configuration hub for schedules, emails, and localization.
* `outlook-free-slot-generator-table.ps1`: The logic for generating the graphical HTML table.
* `outlook-free-slot-generator-text.ps1`: The logic for generating the plain text summary.
* `outlook-free-slot-generator-table.bat`: A wrapper to run the table generator instantly.
* `outlook-free-slot-generator-text.bat`: A wrapper to run the text generator instantly.

## ⚙️ Configuration
Tailor the tool to your needs by editing `config.json`. Below is an example of the English configuration:

```json
{
    "CalendarsToRead": ["your.email@company.com"],
    "WorkingHours": { "Start": 9, "End": 18 },
    "LunchBreak": { "Start": 13, "End": 14 },
    "Preferences": {
        "DaysForward": 5,
        "MinSlotDurationMinutes": 30,
        "TableHeaderColor": "#2563eb"
    },
    "Localization": {
        "Culture": "en-US",
        "MailSubject": "Availability for the upcoming days"
        // ... see config file for full localization options
    }
}

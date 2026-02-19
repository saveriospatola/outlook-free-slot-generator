# 📅 Outlook Free Slot Generator

A specialized utility to scan your Outlook calendars (supporting multiple accounts) and automatically generate a new email containing your availability. This tool helps you schedule meetings efficiently without exposing your entire calendar.

The generator accounts for your working hours, lunch breaks, and personal preferences, all defined in a simple configuration file.

## 🚀 Key Features
* **Unified Script Engine**: A single core script (`outlook-free-slot-generator.ps1`) manages all logic via parameters.
* **Multi-Calendar Scanning**: Aggregate availability from several email accounts at once by listing them in the config.
* **Advanced Visual Customization**: Support for **Color Themes** and **Font** selection (family and size) directly from the configuration.
* **Dual Output Modes**: Choose between a **styled HTML table** for professional emails or **clean text** for quick updates.
* **Full Automation**: The script **automatically creates a new Outlook email** with the subject and body already populated.
* **Portable & Secure**: Uses `.bat` wrappers to bypass execution policies—runs on Windows machines with Outlook installed without needing Admin rights.

---

## 🎨 Customization & Themes

You can now radically change the visual appearance of the HTML table by editing the `config.json` file.

### Changing the Color Theme
In the `Preferences` section, update the `"SelectedTheme"` value by choosing from the predefined options in `ColorThemes`:
* `Grigio` (Professional Gray - Default)
* `Blu` (Blue)
* `Verde` (Green)
* `Viola` (Purple)
* `Arancio` (Orange)

### Personalizing the Font
Also within the `Preferences` section, you can define the typography to match your style:
* **FontName**: Enter a list of preferred fonts (e.g., `"Segoe UI, Calibri, Arial, sans-serif"`).
* **FontSize**: Define the text size (e.g., `"11pt"`).

---

## ⚙️ Configuration (config.json)

Below is an example of the section dedicated to visual styling:

```json
{
  "Preferences": {
    "DaysForward": 5,
    "MinSlotDurationMinutes": 30,
    "SelectedTheme": "Blu",
    "FontName": "Segoe UI, Calibri, sans-serif",
    "FontSize": "11pt"
  },
  "ColorThemes": {
    "Grigio": { ... },
    "Blu": { 
      "HeaderBg": "#1e40af", 
      "HeaderText": "#ffffff", 
      "BadgeBg": "#eff6ff",
      "BadgeText": "#1e40af",
      "BadgeBorder": "#dbeafe",
      "RowAlt": "#f8fafc"
    }
  }
}

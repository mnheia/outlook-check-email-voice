Copyright (c) 2018, Mnheia <mnheia@gmail.com>

# outlook-check-email-voice
A VBS script that checks the number of unread messages in the Outlook inbox and reads the result aloud.

The script supports English, German, Russian, Spanish, French and Chinese voices.

# Example
Edit the language at the top of the script:

```
Language = "en"       ' en, de, ru, es, fr, zh
```

The script can be run through Task Scheduler.

If the requested language voice is not installed, the script falls back to an English/default SAPI voice.

# Requirements
- Microsoft Windows
- Microsoft Outlook
- A compatible SAPI voice for the selected language

# Bugs
Please report any bugs or feature requests through the web interface at https://github.com/mnheia/outlook-check-email-voice/issues

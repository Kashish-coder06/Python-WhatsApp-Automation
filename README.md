# Bulk WhatsApp Sender using Python

A Python GUI application that sends personalized WhatsApp messages to multiple contacts using data from an Excel file.

## Features

* Send bulk WhatsApp messages
* Read contacts from Excel (.xlsx)
* Personalized messages using contact names
* Simple Tkinter GUI
* Activity log for message tracking
* Automatic country code support
* WhatsApp Web integration

## Technologies Used

* Python
* Tkinter
* Pandas
* PyWhatKit
* OpenPyXL

## Excel File Format

Create an Excel file with the following columns:

| Name   | Phone       | From    |
| ------ | ----------- | ------- |
| Hamza | 03452152156 | Kashish |
| Ali | 03352501546 | Kashish |

## Installation

Install required libraries:

```bash
pip install pywhatkit pandas openpyxl
```

## Run the Project

```bash
python bulk_whatsapp_sender.py
```

## Message Template Example

```text
Assalamualaikum {Name},

Hope you are doing well.

Regards,
{From}
```

## How It Works

1. Select an Excel file.
2. Enter the message template.
3. Click "Start Sending".
4. The application reads contacts from Excel.
5. It replaces placeholders with actual values.
6. WhatsApp Web opens automatically and sends messages.

## Author

Kashish Kumari


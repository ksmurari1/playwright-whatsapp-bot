# WhatsApp Automation Bot

A Playwright-based WhatsApp Web automation bot built with Python for personalized message automation, message extraction, screenshots, and reporting.

## Overview

This project automates a controlled WhatsApp Web workflow using Playwright and Python.

### Core Flow

```text
contacts.xlsx
     |
     v
Python + Playwright
     |
     v
WhatsApp Web
     |
     +--> Search contact
     +--> Open chat
     +--> Personalize message
     +--> Send message
     +--> Capture screenshot
     +--> Extract last 3 messages
     |
     v
JSON + Excel Reports

Sample input and generated reports are kept locally to avoid publishing personal WhatsApp/contact data
# desktopWksSpredsheet23

## Overview
This is a Python PyQt desktop automation tool that automates outbound calls using a spreadsheet of phone numbers and mapped audio files.  
It simulates dialing via `pyautogui`, plays the corresponding audio message, and logs results back into the spreadsheet.

## Features
- Load `.ods` spreadsheet of phone numbers + audio files
- Simulated dialing using mouse/keyboard automation
- Multiple audio backends: VLC, sounddevice, winsound (Windows), or system default
- Configurable click coordinates for dial/hang buttons
- Call timeouts, delays, and automatic hangup
- GUI with Start/Stop controls and real-time logging
- Spreadsheet status updates with timestamps

## Requirements
- Python 3.9 or later  
- VLC installed (if using VLC backend)  

## Installation
1. Install dependencies:
   ```bash
   pip install -r requirements.txt

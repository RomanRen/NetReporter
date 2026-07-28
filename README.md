# NetReporter
# NetReporter — AI-Powered Network Diagnostics Utility

A Python desktop utility that monitors your internet connection in real time, logs the results, and uses Google's Gemini API to generate a plain-language diagnosis of your connection quality — plus automated Excel reporting.

## Features

- **Continuous Monitoring** — samples download speed, upload speed, and latency every 10 seconds in a background thread
- **AI-Powered Diagnostics** — sends live network metrics to the Gemini API and returns a one-line, human-readable diagnosis and suggested fix (in Hebrew)
- **Drop Detection** — detects full connection drops via an HTTP-based reachability check, independent of the speed test
- **Automatic Fallback** — falls back to a lightweight HTTP-based throughput estimate if the primary speed test library is unavailable or fails
- **CSV Logging** — every sample is appended to a local CSV log (`network_raw_log.csv`) with timestamp, download/upload speed, ping, and status
- **Automated Excel Reporting** — aggregates the last 4 hours of samples into a summary row (max/min/average speed, total drops) and appends it to an Excel report (`network_report.xlsx`) on a recurring schedule, with a manual "Generate Report" button as well
- **Light/Dark Theme Toggle** — simple desktop GUI built with Tkinter

## Tech Stack

- **Language:** Python
- **GUI:** Tkinter
- **AI:** Google Gemini API (`google-genai`)
- **Data / Reporting:** pandas, openpyxl
- **Speed Testing:** `speedtest-cli`, with a custom HTTP-based fallback
- **Packaging:** PyInstaller (bundled as a standalone macOS app)

## Getting Started

**Requirements:** Python 3.9+

```bash
# Clone the repository
git clone https://github.com/RomanRen/NetReporter.git
cd NetReporter

# Install dependencies
pip install google-genai pandas openpyxl speedtest-cli
```

### Configuration

The app reads your Gemini API key from an environment variable:

```bash
export GEMINI_API_KEY="your-api-key-here"
```

Get a key from [Google AI Studio](https://aistudio.google.com/).

### Run

```bash
python netreporter.py
```

Click **"Start Monitoring"** to begin. Live results and AI-generated diagnostics stream into the on-screen console. Click **"Generate Excel Report"** at any time to produce a summary report on demand — it also opens automatically.

### Output Files

Both files are written to your Desktop:

| File | Contents |
|---|---|
| `network_raw_log.csv` | Every individual sample: timestamp, download/upload Mbps, ping (ms), status |
| `network_report.xlsx` | Periodic aggregated summaries: max/min/average download speed and total drops over each window |

## Building a Standalone App

A PyInstaller spec file (`netreporter.spec`) is included to package the app into a standalone executable:

```bash
pyinstaller netreporter.spec
```

## License

This project is for portfolio and educational purposes.

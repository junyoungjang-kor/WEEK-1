# Hanwha BNCP Weekly Security Report Generator

Hanwha BNCP 이라크 현장 주간 보안 보고서 자동 생성 도구

## Features
- Korean → English AI translation (Claude API)
- Auto-fill routine daily entries (daily checks, submissions)
- SSG shift change tracking (every 4 days)
- Auto-format finance entries (IQD currency)
- Client Feedback structured formatting (Issue/Summary/Actions)
- GUI interface (tkinter)

## Usage
```
python weekly_report_generator.py
```

## Requirements
- Python 3.12+
- anthropic (Claude API)
- tkinter (built-in)

## Setup
1. Run the application
2. Enter your Anthropic API key in the top bar
3. Click "Save & Test" to verify
4. Fill in weekly data and generate report

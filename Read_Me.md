# Meraki LCS Report - User Guide

## Prerequisites
- Python 3.x installed
- Meraki API key set as environment variable `MERAKI_API_KEY`
- Template PowerPoint file (default: `template.pptx` in project directory)

## Installation

### Setting up your environment
```bash
# Clone repository (if applicable)
git clone <repository-url>

# Install required dependencies
pip install -r requirements.txt
```

or install dependencies individually:

```bash
pip install meraki python-pptx requests beautifulsoup4 numpy pandas scikit-learn python-dateutil
```

### Setting up your Meraki API key
```bash
# For Linux/Mac
export MERAKI_API_KEY='your_api_key_here'

# For Windows (Command Prompt)
set MERAKI_API_KEY=your_api_key_here

# For Windows (PowerShell)
$env:MERAKI_API_KEY="your_api_key_here"
```

## Usage
```bash
python main.py -o <organization_id> [options]
```

### Required Arguments
- `-o [ORG_IDS]`: Space-separated list of Meraki organization IDs

### Common Options
- `-d, --days [DAYS]`: Number of days to look back for client data (1-31, default: 14)
- `--output [PATH]`: Custom output path for PowerPoint file (default: "meraki_report.pptx")
- `--template [PATH]`: Custom PowerPoint template path (default: "template.pptx")
- `--slides [LIST]`: Comma-separated list of slide numbers to generate (default: all)
- `--debug`: Enable verbose debugging output
- `--keep-all-slides`: Don't remove slides for missing device types

### Product Adoption Flags
- `--secure-connect`: Indicate organization has Secure Connect deployed
- `--umbrella`: Indicate organization has Umbrella deployed
- `--thousand-eyes`: Indicate organization has Thousand Eyes deployed
- `--spaces`: Indicate organization has Spaces deployed
- `--xdr`: Indicate organization has XDR deployed

### Debug Module Flags
- `--debug-clients`: Run only the clients dashboard slide
- `--debug-mx`: Run only the MX firmware restrictions slide
- `--debug-adoption`: Run only the product adoption slide
- `--debug-executive-summary`: Run only the executive summary slide

## Output
The program generates a comprehensive PowerPoint presentation with slides including:

1. Title slide with organization name(s)
2. Executive summary
3. Dashboard summary (networks, inventory, active nodes, client stats)
4. Firmware status for all device types (MX, MS, MR, MV, MG)
5. Firmware compliance reporting
6. End of Life product summary and details
7. Product adoption overview
8. Predictive lifecycle management

The report automatically adapts to show only relevant slides based on your organization's device inventory.

## Report Methodology

### Executive Summary - Network Health
Network health starts with a score of 100 and deducts points based on:
- Critical devices: -30 points for ≥25%, -20 points for ≥15%, -10 points for ≥5%, -5 points for >0%
- Warning devices: -15 points for ≥40%, -10 points for ≥25%, -5 points for ≥10%
- End of Support devices: -25 points for ≥20%, -15 points for ≥10%, -10 points for ≥5%, -5 points for >0%
- End of Sale devices: -10 points for ≥30%, -5 points for ≥15%
- Critical firmware devices: -20 points for ≥25%, -15 points for ≥15%, -10 points for ≥5%, -5 points for >0%
- Missing core products (MX, MS, MR): -5 points each
- Missing advanced products (Secure Connect, Umbrella): -2 points each

Final scores translate to ratings: 90-100 (Excellent), 80-89 (Good), 70-79 (Satisfactory), 
60-69 (Fair), 40-59 (Needs Attention), Below 40 (Critical Issues).

### Executive Summary - Device Health
Devices are categorized as:
- **Critical**: End-of-support reached OR firmware categorized as critical
- **Warning**: Approaching end-of-support (within 1 year) OR end-of-sale reached OR firmware categorized as warning
- **Good**: None of the above conditions

### Firmware Compliance
- **Critical**: Device is on a significantly outdated firmware version (major version behind latest)
- **Warning**: Device is not on the latest firmware but not critically outdated
- **Good**: Device running latest firmware version

### Predictive Lifecycle Risk Assessment
Risk scores (0-100) are calculated based on:
- **End of Support timing**: 
  - Already EOL: +70 points
  - Within 6 months: +60 points
  - Within 1 year: +40 points
  - Within 2 years: +20 points
- **End of Sale timing**:
  - Already past end-of-sale: +25 points
  - Within 6 months: +20 points
  - Within 1 year: +15 points

Risk categories:
- **High Risk**: Score ≥ 70 (prioritized for immediate replacement)
- **Medium Risk**: Score 40-69 (scheduled for near-term replacement)
- **Low Risk**: Score < 40 (no urgent replacement needed)

## Example Commands
```bash
# Basic report for a single organization
python main.py -o 123456

# Report for multiple organizations with 30-day client history
python main.py -o 123456 789012 -d 30

# Generate only specific slides with custom output
python main.py -o 123456 --slides 1,2,3,6 --output custom_report.pptx
```

## Troubleshooting
- If you encounter API rate limiting, the program will automatically adjust and retry.
- For detailed debugging information, use the `--debug` flag.
- If slides are missing, check that your organization has the corresponding device types.
- Verify your API key has appropriate permissions for all organizations.
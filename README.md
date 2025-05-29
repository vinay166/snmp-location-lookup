# SNMP Location Lookup Tool

This tool reads device names from an Excel file, queries the LibreNMS API for each device, updates the Excel file with the API response data, and checks if the location field follows a compliant format.

## Features

- **Multi-Sheet Processing**: Processes all sheets in the Excel file
- **Progress Tracking**: Shows real-time progress during execution in the command line
- **Device Lookup**: 
  - Reads device names from an Excel file
  - Appends domain suffix (default: `.sac.ragingwire.net`) to hostnames if not already present
  - Queries the LibreNMS API for each device
- **Data Enrichment**: Updates the Excel file with the following information from the API:
  - hostname
  - ip
  - sysDescr
  - hardware
  - os
  - version
  - last_polled
  - location
- **Location Compliance Checking**:
  - Adds an "Expected_Location" column that shows the expected location format built from Excel columns
  - Adds a "Compliant" column that indicates whether the location field from the API matches the expected location
- **Status Tracking**:
  - Adds a "Status" column that shows if the device was found in LibreNMS
  - Clearly marks devices not found in LibreNMS
  - Performs DNS lookup for devices not found in LibreNMS
  - Adds "DNS_IP" and "DNS_Status" columns showing DNS lookup results
- **Formatted Output**:
  - Color-coded cells for compliant (green) and non-compliant (red) locations
  - Yellow highlighting for devices not found in LibreNMS
  - Auto-adjusted column widths for better readability
  - Formatted headers for better visibility
  - Summary sheet with statistics for all processed sheets
- **Backup Creation**: Automatically creates a backup of the original Excel file before making changes

## Requirements

- Python 3.6+
- Required Python packages (install via `pip install -r requirements.txt`):
  - pandas
  - openpyxl
  - requests

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python snmp_location_lookup.py --excel "IDF MDF Audit March 2025 (003).xlsx" --api-url "https://10.1.0.183" --api-token "your-api-token"
```

### Command Line Arguments

- `--excel`: Path to the Excel file (required)
- `--api-url`: LibreNMS API URL (required)
- `--api-token`: LibreNMS API token (required)
- `--location-format`: Format template with column references (e.g., `$B.$C$E.$F`) (default: `$B.$C.$D.$E`)
- `--domain-suffix`: Domain suffix to append to hostnames (default: `.sac.ragingwire.net`)
- `--device-column`: Zero-based index of the column containing device names (default: 0, which is column A)

### How to Specify Location Format

The location format uses column references from your Excel file to build the expected location string. For example:

- `$B.$C.$D.$E` means: Take values from columns B, C, D, and E, and join them with periods
- `$B.$C$E.$F` means: Take values from columns B, C, E, and F, joining with periods except between C and E

#### Column Reference Guide

| Excel Column | Reference |
|--------------|----------|
| A            | $A       |
| B            | $B       |
| C            | $C       |
| ...          | ...      |
| Z            | $Z       |

### Examples

#### Basic Example

```bash
python snmp_location_lookup.py --excel "IDF MDF Audit March 2025 (003).xlsx" --api-url "https://10.1.0.183" --api-token "56edba407b43647ec53db30320e64303"
```

#### Custom Location Format

```bash
python snmp_location_lookup.py --excel "IDF MDF Audit March 2025 (003).xlsx" --api-url "https://10.1.0.183" --api-token "56edba407b43647ec53db30320e64303" --location-format "$B.$C$E.$F"
```

#### Specify Device Column

If your device names are in column C (index 2) instead of column A:

```bash
python snmp_location_lookup.py --excel "IDF MDF Audit March 2025 (003).xlsx" --api-url "https://10.1.0.183" --api-token "56edba407b43647ec53db30320e64303" --device-column 2
```

#### Custom Domain Suffix

```bash
python snmp_location_lookup.py --excel "IDF MDF Audit March 2025 (003).xlsx" --api-url "https://10.1.0.183" --api-token "56edba407b43647ec53db30320e64303" --domain-suffix ".example.com"
```

## Output Format

### Excel Output

After running the script, the Excel file will be updated with the following additional columns:

- `hostname`: The hostname from the LibreNMS API
- `ip`: The IP address of the device
- `sysDescr`: System description from SNMP
- `hardware`: Hardware model information
- `os`: Operating system information
- `version`: Software version
- `last_polled`: When the device was last polled by LibreNMS
- `location`: The actual location string from the LibreNMS API
- `Expected_Location`: The expected location string built from your Excel columns
- `Compliant`: "Yes" if the actual location matches the expected location, "No" otherwise
- `Status`: "Found" if the device was found in LibreNMS, "Not found in LibreNMS" otherwise
- `DNS_IP`: IP address from DNS lookup (only for devices not found in LibreNMS)
- `DNS_Status`: Status of DNS lookup ("Found in DNS", "Not found in DNS", or error message)

#### Summary Sheet

The script also creates a dedicated "Summary" sheet in the Excel file with the following information:

- Sheet Name: Name of each processed sheet
- Total Devices: Number of devices in each sheet
- Devices Found: Number of devices found in LibreNMS
- Devices Not Found: Number of devices not found in LibreNMS
- Compliant Locations: Number of devices with compliant locations
- Non-Compliant Locations: Number of devices with non-compliant locations
- Processed Date: Date and time when the sheet was processed

A "TOTAL" row is added at the bottom of the summary sheet with the sum of all statistics.

A backup of your original Excel file will be created with a `.bak` extension before any changes are made.

### Command Line Summary

The script provides a detailed summary in the command line after processing each sheet:

```
Summary for sheet 'Sheet1':

  Total devices processed: 50
  Devices found in LibreNMS: 45
  Devices not found in LibreNMS: 5
  Devices with compliant location: 40
  Devices with non-compliant location: 5
```

This summary helps you quickly identify how many devices were processed, found, and have compliant locations without having to manually count them in the Excel file.

## Notes

- The script assumes the first column in the Excel file contains device names
- SSL certificate verification is disabled by default when making API requests
- The script will update the existing Excel file in-place
- Column references in the location format are 0-indexed (A=0, B=1, etc.)
- If a column reference is out of range, it will be ignored with a warning

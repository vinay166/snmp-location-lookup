#!/usr/bin/env python3
"""
Create a sample Excel file for demonstrating the SNMP Location Lookup script
"""

import pandas as pd

# Create a sample DataFrame
data = {
    'Device Name': [
        'CA2-RDC-ISP-EDGE-01',
        'CA2-RDC-CORE-SW-01',
        'CA2-RDC-ACCESS-SW-01',
        'CA2-RDC-MGMT-SW-01',
        'CA3-RDC-ISP-EDGE-01'
    ],
    'Location': ['', '', '', '', ''],
    'Rack': ['', '', '', '', ''],
    'Notes': ['', '', '', '', '']
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
df.to_excel('IDF MDF Audit March 2025 (003).xlsx', index=False)

print("Sample Excel file created: IDF MDF Audit March 2025 (003).xlsx")

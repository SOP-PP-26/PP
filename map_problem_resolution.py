import pandas as pd

# Load the Excel file
#file_path = r'C:\Users\SekarAkh\OneDrive - Electrolux\Documents\Akhila Work\PP\tickets.xlsx'  # Update this path if needed
file_path = r'./tickets.xlsx'
df = pd.read_excel(file_path)

# Define resolution mapping (keywords → resolution)
resolution_map = {
    'Plumbing issue': ['leak', 'drip', 'pipe'],
    'Electrical problem': ['switch', 'power', 'electric', 'voltage'],
    'Mechanical issue': ['noise', 'jam', 'grind'],
    'Physical damage': ['broken', 'crack', 'dent'],
    'Performance issue': ['slow', 'lag', 'freeze'],
    'Software issue': ['error', 'crash', 'bug'],
    'Overheating': ['heat', 'hot', 'burn'],
    'Cooling issue': ['cold', 'cool', 'ice']
}

# Function to find resolution based on keywords
def map_resolution_with_confidence(description):
    description = str(description).lower()
    best_resolution = 'Unknown issue'
    max_matches = 0
    confidence = 0

    for resolution, keywords in resolution_map.items():
        match_count = sum(1 for keyword in keywords if keyword in description)
        if match_count > max_matches:
            max_matches = match_count
            best_resolution = resolution
            confidence = int((match_count / len(keywords)) * 100)

    return best_resolution, confidence

for index, row in df.iterrows():
    ticket = row['Ticket']
    pnc = row['PNC']
    resolution, confidence = map_resolution_with_confidence(pnc)
    print(f"Ticket #{ticket}: {pnc}")
    print(f"→ Resolution: {resolution} (Confidence: {confidence}%)\n")

    # Create a list to store the output rows
output_rows = []

# Process each ticket and collect results
for index, row in df.iterrows():
    ticket = row['Ticket']
    problem = row['Description']
    resolution, confidence = map_resolution_with_confidence(problem)
    
    # Print to screen
    print(f"Ticket #{ticket}: {problem}")
    print(f"→ Resolution: {resolution} (Confidence: {confidence}%)\n")
    
    # Add to output list
    output_rows.append({
        'Ticket': ticket,
        'Description': problem,
        'Resolution': resolution,
        'Confidence (%)': confidence
    })

# Create a new DataFrame from the results
output_df = pd.DataFrame(output_rows)

# Save to CSV
#output_file = r'C:\Users\SekarAkh\OneDrive - Electrolux\Documents\Akhila Work\PP\resolved_tickets.csv'  # Update path if needed
output_file = r'./tickets.xlsx'
output_df.to_csv(output_file, index=False)

print(f"✅ Results saved to: {output_file}")
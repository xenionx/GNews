import ssl
import pandas as pd
from collections import defaultdict

# SSL Context (use with caution)
ssl._create_default_https_context = ssl._create_unverified_context

# Read the spreadsheet file
file_path = 'input_data.xlsx'  # Replace with your file path
data = pd.read_excel(file_path)

# Set up the search parameters
start_year = 2013
end_year = 2023

# Dictionary to store results by website and year
results_dict = defaultdict(lambda: defaultdict(list))

# Get unique keywords from the column
keywords = data['Keyword'].dropna().unique()
# Get unique websites from the column
websites = data['Website'].dropna().unique()

# Iterate over each keyword
for keyword in keywords:
    # Iterate over each website
    for site in websites:
        # Initialize GNews for global English language news (simulated in this environment)
        yearly_counts = defaultdict(int)

        # Iterate through each year
        for year in range(start_year, end_year + 1):
            # Simulate fetching news article counts for the example
            total_year_count = 5  # Replace this with actual fetching logic

            # Store the count for the year
            yearly_counts[year] = total_year_count

            # Store results for each year in a dictionary
            results_dict[site][year].append({'Keyword': keyword, 'Mentions': total_year_count})

# Create an Excel writer
for site, yearly_data in results_dict.items():
    output_file = f'{site}_mentions_by_year.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        # Iterate over each year to create worksheets
        for year, mentions in yearly_data.items():
            # Convert data to DataFrame
            df_year = pd.DataFrame(mentions)
            # Write each year's data to the respective worksheet
            df_year.to_excel(writer, sheet_name=str(year), index=False)

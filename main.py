import os
import ssl
import pandas as pd
from gnews import GNews
from collections import defaultdict
from datetime import datetime  # Import for timestamp

# SSL Context (use with caution)
ssl._create_default_https_context = ssl._create_unverified_context

# Read the spreadsheet file
file_path = 'input_data.xlsx'  # Replace with your file path
data = pd.read_excel(file_path)

# Set up the search parameters
start_year = 2013
end_year = 2023

# Create 'Output' directory if it doesn't exist
output_dir = 'Output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

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
        # Initialize GNews for global English language news
        google_news = GNews(language='en', max_results=500)

        # Dictionary to store the count of articles per year
        yearly_counts = defaultdict(int)

        # Iterate through each year
        for year in range(start_year, end_year + 1):
            # Define two 6-month intervals
            intervals = [
                ((year, 1, 1), (year, 6, 30)),   # First 6 months
                ((year, 7, 1), (year, 12, 31))   # Last 6 months
            ]

            # Initialize count for the current year
            total_year_count = 0

            # Iterate over each 6-month interval
            for start_date, end_date in intervals:
                google_news.start_date = start_date
                google_news.end_date = end_date

                try:
                    # Get news articles
                    articles = google_news.get_news(f"intitle:{keyword} site:{site}")
                    # Count articles mentioning the keyword
                    keyword_articles = [article for article in articles if keyword.lower() in article['title'].lower() or keyword.lower() in article['description'].lower()]

                    # Accumulate the count for the year
                    total_year_count += len(keyword_articles)

                except Exception as e:
                    print(f"An error occurred for period {start_date} to {end_date} with site {site}: {e}")

            # Store the combined count for the full year
            yearly_counts[year] = total_year_count

            # Print combined output for the year
            print(f"Retrieved {total_year_count} articles for {year} with keyword '{keyword}' from {site}")

            # Store results for each year in a dictionary
            results_dict[site][year].append({'Keyword': keyword, 'Mentions': total_year_count})

# Create an Excel writer for each website
for site, yearly_data in results_dict.items():
    # Create a timestamp to make the filename unique
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(output_dir, f'{site}_mentions_by_year_{timestamp}.xlsx')

    with pd.ExcelWriter(output_file) as writer:
        # Iterate over each year to create worksheets
        for year, mentions in yearly_data.items():
            # Convert data to DataFrame
            df_year = pd.DataFrame(mentions)
            # Write each year's data to the respective worksheet
            df_year.to_excel(writer, sheet_name=str(year), index=False)

print("Excel files created successfully for each website inside the 'Output' folder.")

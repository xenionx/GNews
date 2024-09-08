import ssl
import pandas as pd
from gnews import GNews
from collections import defaultdict

# SSL Context (use with caution)
ssl._create_default_https_context = ssl._create_unverified_context

# Read the spreadsheet file
file_path = 'input_data.xlsx'  # Replace with your file path
data = pd.read_excel(file_path)

# Set up the search parameters
start_year = 2013
end_year = 2023

# List to store results
results = []

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

        # Sum the counts for all years
        total_count = sum(yearly_counts.values())

        # Add the results to the list
        results.append({'Keyword': keyword, 'Website': site, 'Mentions': total_count})

# Create a DataFrame from the results
df_results = pd.DataFrame(results)

# Sort the DataFrame by 'Mentions' in descending order
df_results_sorted = df_results.sort_values(by='Mentions', ascending=False)

# Print the sorted DataFrame
print(df_results_sorted)

# Optionally, save the DataFrame to an Excel or CSV file
df_results_sorted.to_excel('keyword_mentions_sorted.xlsx', index=False)

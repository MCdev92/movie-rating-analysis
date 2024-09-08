import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Load the data
# Column names for u.data
columns = ['user_id', 'movie_id', 'rating', 'timestamp']
ratings = pd.read_csv(r'ml-100k/u.data', sep='\t', names=columns)

# Column names for u.item (movies)
movie_cols = ['movie_id', 'title', 'release_date', 'video_release_date', 'IMDb_URL', 'unknown', 'Action', 
              'Adventure', 'Animation', 'Children', 'Comedy', 'Crime', 'Documentary', 'Drama', 'Fantasy', 
              'Film-Noir', 'Horror', 'Musical', 'Mystery', 'Romance', 'Sci-Fi', 'Thriller', 'War', 'Western']
movies = pd.read_csv(r'ml-100k/u.item', sep='|', names=movie_cols, encoding='latin-1')

# Merge ratings with movie titles
data = pd.merge(ratings, movies[['movie_id', 'title']], on='movie_id')

# Show the first few rows to understand the dataset
print("First few rows of the merged dataset:")
print(data.head())

# Check for missing data
print("\nChecking for missing data:")
print(data.isnull().sum())

# Top 10 movies with the highest average ratings
top_rated_movies = data.groupby('title')['rating'].mean().sort_values(ascending=False).head(10)
print("\nTop 10 Highest Rated Movies:")
print(top_rated_movies)

# Plot rating distribution
plt.figure(figsize=(8,6))
sns.histplot(data['rating'], bins=5, kde=False)
plt.title('Distribution of Movie Ratings')
plt.xlabel('Rating')
plt.ylabel('Count')
plt.show()

# Bar chart of top-rated movies
plt.figure(figsize=(10,6))
top_rated_movies.plot(kind='bar', color='skyblue')
plt.title('Top 10 Highest Rated Movies')
plt.xlabel('Movie Title')
plt.ylabel('Average Rating')
plt.xticks(rotation=45)
plt.show()

# ------------------------------------
# Calculate Histogram Data (for Excel)
# ------------------------------------
# Calculate the histogram bins and counts using numpy
counts, bin_edges = np.histogram(data['rating'], bins=5)

# Prepare the histogram data for Excel
histogram_data = pd.DataFrame({
    'Rating Range (Bins)': [f'{bin_edges[i]:.1f} - {bin_edges[i+1]:.1f}' for i in range(len(bin_edges)-1)],
    'Count': counts
})

print("\nHistogram Data:")
print(histogram_data)

# ------------------------------------
# Prepare Top 10 Highest Rated Movies Data for Excel
# ------------------------------------
# Convert the top_rated_movies Series into a DataFrame
top_rated_movies_df = top_rated_movies.reset_index()
top_rated_movies_df.columns = ['Movie Title', 'Average Rating']

# ------------------------------------
# Export Merged Data, Histogram Data, and Top-Rated Movies to Excel
# ------------------------------------
with pd.ExcelWriter(r'data_files/movie_ratings_analysis_with_histogram_and_top_rated.xlsx', engine='openpyxl') as writer:
    # Export the merged dataset
    data.to_excel(writer, sheet_name='Merged Data', index=False)
    
    # Export the histogram data
    histogram_data.to_excel(writer, sheet_name='Histogram Data', index=False)
    
    # Export the top-rated movies
    top_rated_movies_df.to_excel(writer, sheet_name='Top Rated Movies', index=False)

print("Data exported to data_files/movie_ratings_analysis_with_histogram_and_top_rated.xlsx")

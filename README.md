# Movie-Collection-Analytics
First of all, this is a personal project and therefore I am aware that it has a lot of things to improve.

The script peliculas.py connects to the TMDB api to get information about the movie that the user wants (id, title, score, budget, box office and duration, plus a category that depends on the user).
All this information is stored in a csv file which is then converted into an excel file with two sheets; the first one contains the list of all stored movies, this table is edited with conditional formatting based on the following parameters:

From the score:
Red: score less than 5 
Yellow: score between 5 (included) and 7
Green: score higher than 7 (included)

From profit:
Red: negative benefit
Green: positive benefit

The charts are intended for when you have at least one film from each decade. Right now the csv has 1009 films in total.

The other sheet contains the following graphs
- General:
  - Doughnut diagram of pending and viewed movies.
  - Radar diagram comparing the duration, score, budget and box office of both categories.
- By decades
  - Bar chart with the number of movies stored in each decade. Also by category
  - Pie chart with the above information as percentages.
- By years
  - Number of films stored for each year
  - Average score of films for each year

- Score histogram

- Films for which box office and budget data are available:
  - Scatterplot of both data in 3D.
  - Scatterplot of both data as a function of decade.
  - Scatterplot with the mean and median of budget and box office indicated.
  - Scatterplot with trend (linear regression)
  - Scatterplot of each decade with its trend line (linear regression)

- Treemap of the 10 highest-rated films and the 10 highest-grossing films

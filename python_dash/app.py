import dash
import dash_core_components as dcc
import dash_html_components as html
import pandas as pd
from dash.dependencies import Output, Input

# https://realpython.com/python-dash/
data = pd.read_csv("collection.csv")

years = []
for element in data['YEAR']:
    if element not in years:
        years.append(int(element))
sorted_years = sorted(years, reverse=False)

data_grouped_by_year = data.groupby('YEAR').mean()
mean_score_year = []
mean_gross_year = []
mean_budget_year = []
for element in data_grouped_by_year['SCORE']:
    mean_score_year.append(element)
for element in data_grouped_by_year['GROSS']:
    mean_gross_year.append(element)
for element in data_grouped_by_year['BUDGET']:
    mean_budget_year.append(element)

data_grouped_by_year = data.groupby('YEAR').count()
number_movies = []
for element in data_grouped_by_year['SCORE']:
    number_movies.append(element)

external_stylesheets = [
    {
        "href": "https://fonts.googleapis.com/css2?"
                "family=Lato:wght@400;700&display=swap",
        "rel": "stylesheet",
    },
]

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
app.title = "Movie Analytics"
app.layout = html.Div(
    children=[
        html.Div(
            children=[
                html.P(children="🥑", className="header-emoji"),
                html.H1(
                    children="Movie Analytics: Understand Your ¿Favourite? Movies!", className="header-title"
                ),
                html.P(
                    children="Analyze the behavior of movie ratings between 1980 and 2030",
                    className="header-description",
                ),
            ],
            className="header",
        ),
        html.Div(
            children=[
                html.Div(
                    children=[
                        html.Div(children="Category", className="menu-title"),
                        dcc.Dropdown(
                            id="type-filter",
                            options=[
                                {"label": avocado_type, "value": avocado_type}
                                for avocado_type in data.CATEGORY.unique()
                            ],
                            value="Vista",
                            clearable=False,
                            searchable=False,
                            className="dropdown",
                        ),
                    ],
                ),
                html.Div(
                    children=[
                        html.Div(
                            children="Date Range",
                            className="menu-title"
                        ),
                        dcc.DatePickerRange(
                            id="date-range",
                            start_date=str(data.YEAR.min()),
                            end_date=str(data.YEAR.max()),
                        ),
                    ]
                ),
            ],
            className="menu",
        ),
        html.Div(
            children=[
                html.Div(
                    children=dcc.Graph(
                        id="rating-chart",
                        config={"displayModeBar": False},
                        figure={
                            "data": [
                                {
                                    "x": sorted_years,
                                    "y": mean_score_year,
                                    "type": "lines",
                                    "hovertemplate": "%{y:.2f}"
                                                     "<extra></extra>",
                                },
                            ],
                            "layout": {
                                "title": {
                                    "text": "Average Rating of Movies",
                                    "x": 0.05,
                                    "xanchor": "left",
                                },
                                "xaxis": {"fixedrange": True},
                                "yaxis": {
                                    "fixedrange": True,
                                },
                                "colorway": ["#17B897"],
                            },
                        },
                    ),
                    className="card",
                ),
                html.Div(
                    children=dcc.Graph(
                        id="gross-chart",
                        config={"displayModeBar": False},
                        figure={
                            "data": [
                                {
                                    "x": sorted_years,
                                    "y": mean_gross_year,
                                    "type": "lines",
                                    "hovertemplate": "%{y:.2f}"
                                                     "<extra></extra>",
                                },
{
                                    "x": sorted_years,
                                    "y": mean_budget_year,
                                    "type": "lines",
                                    "hovertemplate": "%{y:.2f}"
                                                     "<extra></extra>",
                                },
                            ],
                            "layout": {
                                "title": {
                                    "text": "Average Gross vs Budget of Movies",
                                    "x": 0.05,
                                    "xanchor": "left",
                                },
                                "xaxis": {"fixedrange": True},
                                "yaxis": {"fixedrange": True,
                                },
                                "colorway": ["#17B897"],
                            },
                        },
                    ),
                    className="card",
                ),
                html.Div(
                    children=dcc.Graph(
                        id="volume-chart",
                        config={"displayModeBar": False},
                        figure={
                            "data": [
                                {
                                    "x": sorted_years,
                                    "y": number_movies,
                                    "type": "bar",
                                },
                            ],
                            "layout": {
                                "title": {
                                    "text": "Movies stored",
                                    "x": 0.05,
                                    "xanchor": "left",
                                },
                                "xaxis": {"fixedrange": True},
                                "yaxis": {"fixedrange": True},
                                "colorway": ["#E12D39"],
                            },
                        },
                    ),
                    className="card",
                ),
            ],
            className="wrapper",
        ),

    ]

)


@app.callback(
    [Output("rating-chart", "figure"), Output("gross-chart", "figure"), Output("volume-chart", "figure")],
    [
        Input("type-filter", "value"),
        Input("date-range", "start_date"),
        Input("date-range", "end_date"),
    ],
)
def update_charts(avocado_type, start_date, end_date):
    mask = (
            (data.CATEGORY == avocado_type)
            & (data.YEAR >= int(start_date[0:4]))
            & (data.YEAR <= int(end_date[0:4]))
    )

    filtered_data = data.loc[mask, :]
    print(filtered_data)

    years_filtered = []
    for element1 in filtered_data['YEAR']:
        if element1 not in years_filtered:
            years_filtered.append(int(element1))
    sorted_years_filtered = sorted(years_filtered, reverse=False)

    data_filtered_grouped_by_year = filtered_data.groupby('YEAR').mean()
    mean_score_year_filtered = []
    for element2 in data_filtered_grouped_by_year['SCORE']:
        mean_score_year_filtered.append(element2)

    data_filtered_grouped_by_year = filtered_data.groupby('YEAR').mean()
    mean_gross_year_filtered = []
    for element4 in data_filtered_grouped_by_year['GROSS']:
        mean_gross_year_filtered.append(element4)

    data_filtered_grouped_by_year = filtered_data.groupby('YEAR').mean()
    mean_budget_year_filtered = []
    for element5 in data_filtered_grouped_by_year['BUDGET']:
        mean_budget_year_filtered.append(element5)

    data_grouped_by_year_filtered = filtered_data.groupby('YEAR').count()
    number_movies_filtered = []
    for element3 in data_grouped_by_year_filtered['SCORE']:
        number_movies_filtered.append(element3)

    rating_chart_figure = {
        "data": [
            {
                "x": sorted_years_filtered,
                "y": mean_score_year_filtered,
                "type": "lines",
                "hovertemplate": "%{y:.2f}<extra></extra>",
            },
        ],
        "layout": {
            "title": {
                "text": "Average Rating of Movies",
                "x": 0.05,
                "xanchor": "left",
            },
            "xaxis": {"fixedrange": True},
            "yaxis": {"fixedrange": True},
            "colorway": ["#17B897"],
        },
    }

    gross_chart_figure = {
        "data": [
            {
                "x": sorted_years_filtered,
                "y": mean_gross_year_filtered,
                "type": "lines",
                'name': 'Mean Gross',
                "hovertemplate": "%{y:.2f}<extra></extra>",
            },
            {
                "x": sorted_years_filtered,
                "y": mean_budget_year_filtered,
                "type": "lines",
                'name': 'Mean Budget',
                "hovertemplate": "%{y:.2f}<extra></extra>",
            },
        ],
        "layout": {
            "title": {
                "text": "Average Gross vs Budget of Movies",
                "x": 0.05,
                "xanchor": "left",
            },
            "xaxis": {"fixedrange": True},
            "yaxis": {"fixedrange": True},
            "colorway": ["#4661af", "#a361af"],
        },
    }

    volume_chart_figure = {
        "data": [
            {
                "x": sorted_years_filtered,
                "y": number_movies_filtered,
                "type": "bar",
            },
        ],
        "layout": {
            "title": {
                "text": "Movies Stored",
                "x": 0.05,
                "xanchor": "left"
            },
            "xaxis": {"fixedrange": True},
            "yaxis": {"fixedrange": True},
            "colorway": ["#E12D39"],
        },
    }
    return rating_chart_figure, gross_chart_figure, volume_chart_figure


if __name__ == "__main__":
    app.run_server(debug=True)

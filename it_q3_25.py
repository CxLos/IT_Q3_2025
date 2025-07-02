# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
from datetime import datetime
import folium
import os
import sys
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component

# Google Web Credentials
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
# data_path = 'data/IT_Responses.xlsx'
# file_path = os.path.join(script_dir, data_path)
# data = pd.read_excel(file_path)
# df = data.copy()

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1wNUS59k4D6mSq-ciF6PDkcIcrmr9uu867lqP4fM6VyA/edit?resourcekey=&gid=1758812507#gid=1758812507"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
worksheet = sheet.get_worksheet(0)  # ✅ This grabs the first worksheet
data = pd.DataFrame(worksheet.get_all_records())
df = data.copy()

# Trim leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()

# Filtered df where 'Date of Activity:' is between January and March
df['Date of Activity'] = pd.to_datetime(df['Date of Activity'], errors='coerce')
df = df[(df['Date of Activity'].dt.month >= 1) & (df['Date of Activity'].dt.month <= 3)]
df['Month'] = df['Date of Activity'].dt.month_name()

df_1 = df[df['Month'] == 'January']
df_2 = df[df['Month'] == 'February']
df_3 = df[df['Month'] == 'March']

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# print(df_m.head())
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns)
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())
# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

it_columns =[
    'Timestamp', 
    'Date of Activity',
    'Which form are you filling out?',
    'Person submitting this form:',
    'Activity Duration (hours):',
    'How much time did you spend on these tasks? (minutes)',
    'Please provide brief description of activity', 
    'Email Address',
    'Any recent or planned changes to BMHC lead services or programs?',
    'Month'
]

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

# ============================== Data Preprocessing ========================== #

df.rename(
    columns={
        "Activity Duration (hours):": "Hours",
        "Which form are you filling out?": "Form Type",
        "Person submitting this form:": "Person",
    }, 
inplace=True)

# Get the reporting quarter:
def get_custom_quarter(date_obj):
    month = date_obj.month
    if month in [10, 11, 12]:
        return "Q1"  # October–December
    elif month in [1, 2, 3]:
        return "Q2"  # January–March
    elif month in [4, 5, 6]:
        return "Q3"  # April–June
    elif month in [7, 8, 9]:
        return "Q4"  # July–September

# Reporting Quarter (use last month of the quarter)
report_date = datetime(2025, 3, 1)  # Example report date for Q2 (Jan–Mar)
month = report_date.month

current_quarter = get_custom_quarter(report_date)
print(f"Reporting Quarter: {current_quarter}")

# Extract year from report_date:
report_year = report_date.year

# ============================== IT Events ========================== #

it_events = len(df)
# print('IT Events:', it_events)

# ============================== IT Hours ========================== #

hours_unique =[
    4, 1, 5, 2, 0.5, 3, '', 80, 1.5
]

# print("IT Hours unique before:", df['Hours'].unique().tolist())

df['Hours'] = (df['Hours']
    .astype(str)
    .str.strip()
    .replace({
        '': 0,
        })
    )   

df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
df['Hours'] = df['Hours'].fillna(0)
# print("IT Hours unique after:", df['Hours'].unique().tolist())

# Adjust the quarter calculation for custom quarters
if month in [10, 11, 12]:
    quarter = 1  # Q1: October–December
elif month in [1, 2, 3]:
    quarter = 2  # Q2: January–March
elif month in [4, 5, 6]:
    quarter = 3  # Q3: April–June
elif month in [7, 8, 9]:
    quarter = 4  # Q4: July–September
    
    # Calculate start and end month indices for the quarter
all_months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]
start_month_idx = (quarter - 1) * 3
month_order = all_months[start_month_idx:start_month_idx + 3]

# Define a mapping for months to their corresponding quarter
quarter_months = {
    1: ['October', 'November', 'December'],  # Q1
    2: ['January', 'February', 'March'],    # Q2
    3: ['April', 'May', 'June'],            # Q3
    4: ['July', 'August', 'September']      # Q4
}

# Get the months for the current quarter
months_in_quarter = quarter_months[quarter]

# Calculate total hours for each month in the current quarter
hours = []
for month in months_in_quarter:
    hours_in_month = df[df['Month'] == month]['Hours'].sum()
    hours_in_month = round(hours_in_month)
    hours.append(hours_in_month)
    # print(f'IT hours in {month}:', hours_in_month, 'hours')

it_hours = df.groupby('Hours').size().reset_index(name='Count')
it_hours = df['Hours'].sum()
it_hours = round(it_hours)
# print('Q2 IT hours:', it_hours, 'hours')

# Create DataFrame for IT Hours
df_hours = pd.DataFrame({
    'Month': months_in_quarter,
    'Hours': hours
})

# Bar chart for IT Hours
hours_fig = px.bar(
    df_hours,
    x='Month',
    y='Hours',
    color="Month",
    text='Hours',
    title= f'{current_quarter} IT Hours by Month',
    labels={
        'Hours': 'Number of Hours',
        'Month': 'Month'
    },
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Hours',
    height=900,  # Adjust graph height
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickfont=dict(size=18),  # Adjust font size for the month labels
        tickangle=-25,  # Rotate x-axis labels for better readability
        title=dict(
            text=None,
            font=dict(size=20),  # Font size for the title
        ),
    ),
    yaxis=dict(
        title=dict(
            text='Number of Hours',
            font=dict(size=22),  # Font size for the title
        ),
    ),
    bargap=0.08,  # Reduce the space between bars
).update_traces(
    texttemplate='%{text}',  # Display the count value above bars
    textfont=dict(size=25),  # Increase text size in each bar
    textposition='auto',  # Automatically position text above bars
    textangle=0, # Ensure text labels are horizontal
    hovertemplate=(  # Custom hover template
        '<b>Hours</b>: %{y}<extra></extra>'  
    ),
).add_annotation(
    x='January',  # Specify the x-axis value
    # y=df_hours.loc[df_hours['Month'] == 'January', 'Hours'].values[0] + 100,  # Position slightly above the bar
    text='',  # Annotation text
    showarrow=False,  # Hide the arrow
    font=dict(size=30, color='red'),  # Customize font size and color
    align='center',  # Center-align the text
)

# Pie Chart IT Hours
hours_pie = px.pie(
    df_hours,
    names='Month',
    values='Hours',
    color='Month',
    height=800
).update_layout(
    title=dict(
        x=0.5,
        text=f'{current_quarter} IT Hours by Month',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
    margin=dict(
        l=0,  # Left margin
        r=0,  # Right margin
        t=100,  # Top margin
        b=0   # Bottom margin
    )  # Add margins around the chart
).update_traces(
    rotation=180,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=25),  # Increase text size in each bar
    textinfo='value+percent',
    # texttemplate='<br>%{percent:.0%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# -------------------- Person completing this form DF ------------------- #

# Group the data by 'Month' and 'Person completing this form:.4' and count occurrences
df_person_counts = (
    df.groupby(['Month', 'Person'], sort=False)
    .size()
    .reset_index(name='Count')
)

# print('Person Counts: \n', df_person_counts.head())

# Assign categorical ordering to the 'Month' column
df_person_counts['Month'] = pd.Categorical(
    df_person_counts['Month'],
    categories=months_in_quarter,
    ordered=True
)

# print("Mont Order:", month_order)

# Create the grouped bar chart
person_fig = px.bar(
    df_person_counts,
    x='Month',
    y='Count',
    color='Person',
    barmode='group',
    text='Count',
    title=f'{current_quarter} Form Submissions by Person',
    labels={
        'Count': 'Number of Submissions',
        'Month': 'Month',
        'Person': 'Person'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_person_counts['Month'].unique(),
        tickangle=-35
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    hovermode='x unified'
).update_traces(
        textfont=dict(size=25),  # Increase text size in each bar
    textposition='outside',
    hovertemplate='<br><b>Count: </b>%{y}<br>',
    customdata=df_person_counts['Person'].values.tolist()
)

# Group the data by "Person completing this form:" to count occurrences
df_pf = df.groupby('Person').size().reset_index(name="Count")

person_pie = px.pie(
    df_pf,
    names='Person',
    values='Count',
    color='Person',
    height=800
).update_layout(
    title=dict(
        x=0.5,
        text= f'{current_quarter} People Completing Forms',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
    margin=dict(
        t=150,  # Adjust the top margin (increase to add more padding)
        l=20,   # Optional: left margin
        r=20,   # Optional: right margin
        b=20    # Optional: bottom margin
    )
).update_traces(
    rotation=0,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=25),  # Increase text size in each bar
    textinfo='value+percent',
    insidetextorientation='horizontal',  # Horizontal text orientation
    texttemplate='%{value}<br>%{percent:.0%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# # ========================== DataFrame Table ========================== #

# df_data2 data table
data_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df[col] for col in df.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

data_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=800,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server   

app.layout = html.Div(
  children=[ 
    html.Div(
        className='divv', 
        children=[ 
          html.H1(
              f'BMHC IT Report {current_quarter} {report_year}', 
              className='title'),
          html.H2( 
              '01/01/2025 - 03/31/2025', 
              className='title2'),
          html.Div(
              className='btn-box', 
              children=[
                  html.A(
                    'Repo',
                    href=f'https://github.com/CxLos/IT_{current_quarter}_2025',
                    className='btn'),
    ]),
  ]),    

# Data Table
# html.Div(
#     className='row0',
#     children=[
#         html.Div(
#             className='table',
#             children=[
#                 html.H1(
#                     className='table-title',
#                     children='Data Table'
#                 )
#             ]
#         ),
#         html.Div(
#             className='table2', 
#             children=[
#                 dcc.Graph(
#                     className='data',
#                     figure=marcom_table
#                 )
#             ]
#         )
#     ]
# ),

# ROW 1
html.Div(
    className='row0',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_quarter} IT Events']
            ),
            html.Div(
                className='circle1',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high3',
                            children=[it_events]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high2',
                children=[f'{current_quarter} IT Hours']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[it_hours]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=hours_fig
                )
            ]
        )
    ]
),

# # ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=hours_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=person_fig
                )
            ]
        )
    ]
),

# # ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=person_pie
                )
            ]
        )
    ]
),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph1',
            children=[
                html.Div(
                    className='table',
                    children=[
                        html.H1(
                            className='table-title',
                            children='Data Events Table'
                        )
                    ]
                ),
                html.Div(
                    className='table2',
                    children=[
                        dcc.Graph(
                            className='data',
                            figure=data_table
                        )
                    ]
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
            dcc.Graph(
                style={'height': '800px', 'width': '1000px'}  # Set height and width
            )
            ],
        )
    ]
),
])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# =================================== Updated Database ================================= #

# updated_path = f'data/IT_{current_quarter}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)
# df.to_excel(data_path, index=False)
# print(f"DataFrame saved to {data_path}")

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx
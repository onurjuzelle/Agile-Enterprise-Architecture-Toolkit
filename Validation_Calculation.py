#file_path = '/Users/onurguzel/Downloads/Agile Enterprise Architecture Maturity Model Validation Survey (Responses).xlsx'
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Load the Excel file
file_path = '/Users/onurguzel/Downloads/Agile Enterprise Architecture Maturity Model Validation Survey (Responses).xlsx'
df = pd.read_excel(file_path)

# Clean the column names by stripping leading/trailing spaces
df.columns = df.columns.str.strip()

# Map Likert scale responses to numerical values
likert_scale = {
    'Strongly Disagree': 1,
    'Disagree': 2,
    'Neutral': 3,
    'Agree': 4,
    'Strongly Agree': 5
}

# Apply the mapping to the dataframe
df_numerical = df.replace(likert_scale)

# Define the categories and their respective questions
categories = {
    "Accuracy": [
        "[The maturity model can accurately reflect the organizations current state of Agile Enterprise Architecture]",
        "[The model accurately distinguishes between different levels of maturity ]",
        "[The model reflects real world scenarios and challenges faced in the industry]",
        "[The criteria used in the model resonates with my understanding of Agile Enterprise Architecture]",
        "[The model provides valuable guidance for developing our Enterprise Architecture to become more Agile]"
    ],
    "Relevance": [
        "[The criteria for the levels are relevant for our organizations Enterprise Architecture Practices]",
        "[The model clearly represents Agile values and mindset incorporated into Enterprise Architecture]"
    ],
    "Generalizability": [
        "[The model is generalizable to different types of organizations (e.g small, medium, large))]",
        "[The model is applicable across various sectors (e.g Technology, Finance, Public))]"
    ],
    "Usability": [
        "[The model use actionable insights for improving the agility in Enterprise Architecture]",
        "[The model facilitates constructive discussions about maturity within the organization]",
        "[The description of maturity levels are clear and precise]",
        "[The model is east to understand and use]"
    ],
    "Overall Satisfaction": [
        "[I am satisfied with the overall structure of the maturity model]",
        "[The model meet my expectations of assessing Agile Enterprise Architecture Maturity]",
        "[The model provides valuable guidance for developing our Enterprise Architecture to become more Agile]",
        "[I would recommend this model to other organizations for assessing their Agile Enterprise Architecture Maturity]"
    ]
}

# Remove the leading/trailing spaces from the category questions
categories = {cat: [q.strip() for q in questions] for cat, questions in categories.items()}

# Calculate the average scores for each category
average_scores = {category: df_numerical[questions].mean(axis=1).mean() for category, questions in categories.items()}

# Convert the dictionary to a DataFrame for plotting
average_scores_df = pd.DataFrame(list(average_scores.items()), columns=['Category', 'Average Score'])

# Plot the bar chart using Plotly
fig = px.bar(
    average_scores_df, 
    x='Category', 
    y='Average Score', 
    color='Average Score',
    color_continuous_scale=px.colors.sequential.Viridis
)

fig.update_layout(
    title='Average Scores by Category',
    xaxis_title='Category',
    yaxis_title='Average Score',
    paper_bgcolor='darkgrey',
    plot_bgcolor='darkgrey',
    yaxis=dict(range=[0, 5]),
    xaxis=dict(tickangle=-45)
)

fig.show()

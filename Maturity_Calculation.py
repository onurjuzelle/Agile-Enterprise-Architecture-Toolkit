import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import re
import plotly.express as px
import plotly.graph_objects as go

# Define aspects of Agile Enterprise Architecture
aspects = [
    "Strategy to Execution", "Stakeholder Management", "Thought Leadership", 
    "Architecture Development", "Architecture Control", "Architecture Governance", 
    "Technology Portfolio Management", "Enterprise Model Management", "Agile Enterprise Architecture Culture"
]

# Map columns to aspects based on their descriptions
aspect_column_map = {
    "Please select the statement that you think is the closest description of your organizations current state of Strategy to Execution": "Strategy to Execution",
    "Please select the statement that you think is the closest description of your organizations current state of Stakeholder Management": "Stakeholder Management",
    "Please select the statement that you think is the closest description of your organizations current state of Thought Leadership": "Thought Leadership",
    "Please select the statement that you think is the closest description of your organizations current state of Architectural Development": "Architecture Development",
    "Please select the statement that you think is the closest description of your organizations current state of Architectural Control": "Architecture Control",
    "Please select the statement that you think is the closest description of your organizations current state of Architectural Governance": "Architecture Governance",
    "Please select the statement that you think is the closest description of your organizations current state of Technology Portfolio Management": "Technology Portfolio Management",
    "Please select the statement that you think is the closest description of your organizations current state of Enterprise Model Management": "Enterprise Model Management",
    "Please select the statement that you think is the closest description of your organizations current state of Agile Enterprise Architecture Culture": "Agile Enterprise Architecture Culture"
}

# Strip whitespace from keys in the aspect_column_map
aspect_column_map = {k.strip(): v for k, v in aspect_column_map.items()}

# Initialize main application window
root = tk.Tk()
root.title("Maturity Model Assessment")
root.geometry("1200x800")

def load_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        df = pd.read_excel(file_path)
        # Preprocess the data
        df = preprocess_data(df)
        # Trim whitespace from column names
        df.columns = df.columns.str.strip()
        process_data(df)

def preprocess_data(df):
    # Remove every character after the number in columns G to O (index 6 to 14)
    for col in df.columns[6:15]:
        df[col] = df[col].astype(str).str.extract(r'(\d+)')
    
    # Drop columns B, C, D, E (index 1, 2, 3, 4)
    df.drop(df.columns[[1, 2, 3, 4]], axis=1, inplace=True)
    
    # Standardize the role names by removing special characters and extra spaces
    df['What is your current title?'] = df['What is your current title?'].str.replace(r'[^\w\s]', '', regex=True).str.strip()
    
    # Manually handle known role names
    df['What is your current title?'] = df['What is your current title?'].replace({
        "Architecture Models Insights Expert": "Architecture Models & Insights Expert"
    })
    
    return df

def process_data(df):
    # Define the role column and aspect columns based on their descriptions
    role_column = "What is your current title?"
    df = df.rename(columns={col.strip(): aspect_column_map.get(col.strip(), col) for col in df.columns})

    # Debug: print the columns of the dataframe
    print("Columns in the dataframe:", df.columns)

    # Check if all required columns are present
    missing_columns = [aspect for aspect in aspects if aspect not in df.columns]
    if missing_columns:
        messagebox.showerror("Error", f"Missing columns in the Excel file: {missing_columns}")
        return

    # Include the 'Role' column for role-based processing
    df_aspects = df[aspects + [role_column]]
    df_aspects = df_aspects.rename(columns={role_column: "Role"})

    # Extract numeric scores from aspect columns
    def extract_numeric_value(value):
        match = re.search(r'\d+', str(value))
        return float(match.group()) if match else np.nan

    for aspect in aspects:
        df_aspects[aspect] = df_aspects[aspect].apply(extract_numeric_value)

    # Show the processed dataframe in the GUI
    #display_dataframe(df_aspects)

    # Calculate average scores for each aspect
    avg_scores = df_aspects[aspects].mean().to_dict()

    # Calculate role-based average scores
    roles = [
        "Team Lead",
        "Lead Enterprise Architect",
        "Enterprise Architect",
        "Architecture Models and Insights Expert",
        "Architecture Governance Expert"
    ]
    role_avg_scores = {}
    for role in roles:
        if role in df_aspects['Role'].unique():
            role_avg_scores[role] = df_aspects[df_aspects['Role'] == role][aspects].mean().to_dict()
        else:
            role_avg_scores[role] = {aspect: np.nan for aspect in aspects}

    display_radar_chart(avg_scores)
    display_avg_scores_chart("Average Score per Aspect", avg_scores)
    
    for role in roles:
        display_avg_scores_chart(f"Average Score per Aspect: {role}", role_avg_scores[role])

    display_combined_radar_chart(avg_scores, role_avg_scores)

def display_dataframe(df):
    # Create a new window for displaying the dataframe
    dataframe_window = tk.Toplevel(root)
    dataframe_window.title("Processed DataFrame")
    dataframe_window.geometry("800x600")

    text_widget = tk.Text(dataframe_window)
    text_widget.pack(expand=True, fill=tk.BOTH)

    text_widget.insert(tk.END, df.to_string())

def determine_maturity_level(score):
    if score <= 1.5:
        return "Foundational"
    elif score <= 2.5:
        return "Developed"
    elif score <= 3.5:
        return "Optimized"
    elif score <= 4.5:
        return "Integrated"
    else:
        return "Advanced"

def display_radar_chart(avg_scores):
    if not avg_scores:
        messagebox.showerror("Error", "No valid scores to display.")
        return

    labels = list(avg_scores.keys())
    stats = list(avg_scores.values())
    
    # Use Plotly to create the radar chart with the same color scheme as the bar chart
    fig = go.Figure(data=go.Scatterpolar(
        r=stats + [stats[0]],
        theta=labels + [labels[0]],
        fill='toself',
        marker=dict(color=px.colors.sequential.Viridis[4])
    ))

    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 5]
            ),
        ),
        showlegend=False,
        title='Overall Average Scores per Aspect'
    )

    # Add the scores to the radar chart
    for i, score in enumerate(stats):
        fig.add_trace(go.Scatterpolar(
            r=[score, score],
            theta=[labels[i], labels[i]],
            mode='markers+text',
            text=[f'{score:.1f}'],
            textposition='top center',
            marker=dict(color='black')
        ))

    # Show the radar chart
    fig.show()

def display_combined_radar_chart(avg_scores, role_avg_scores):
    fig = go.Figure()

    # Add the overall average scores
    labels = list(avg_scores.keys())
    stats = list(avg_scores.values())
    fig.add_trace(go.Scatterpolar(
        r=stats + [stats[0]],
        theta=labels + [labels[0]],
        fill='toself',
        name='Overall Average'
    ))

    # Add the scores for each role
    for role, scores in role_avg_scores.items():
        labels = list(scores.keys())
        stats = list(scores.values())

        fig.add_trace(go.Scatterpolar(
            r=stats + [stats[0]],
            theta=labels + [labels[0]],
            fill='toself',
            name=role
        ))

    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 5]
            ),
        ),
        showlegend=True,
        title='Role-based Comparison of Average Scores per Aspect'
    )

    fig.show()

def display_avg_scores_chart(title, scores):
    # Create a new window for the average scores bar chart
    avg_scores_window = tk.Toplevel(root)
    avg_scores_window.title(title)
    avg_scores_window.geometry("800x600")

    # Sort scores by value
    sorted_scores = dict(sorted(scores.items(), key=lambda item: item[1]))

    fig = px.bar(
        x=list(sorted_scores.keys()), 
        y=list(sorted_scores.values()),
        labels={'x': 'Aspect', 'y': 'Average Score'},
        title=title,
        color=list(sorted_scores.values()),
        color_continuous_scale=px.colors.sequential.Viridis
    )

    fig.update_layout(
        paper_bgcolor='darkgrey', 
        plot_bgcolor='darkgrey',
        yaxis=dict(range=[0, 5])  # Set the y-axis range to [0, 5]
    )
    fig.update_xaxes(title='Aspect')
    fig.update_yaxes(title='Average Score')

    # Display the bar chart
    fig.show()


# Add a button to load the Excel file
load_button = tk.Button(root, text="Load Excel File", command=load_excel, height=2)
load_button.pack(pady=20)

root.mainloop()

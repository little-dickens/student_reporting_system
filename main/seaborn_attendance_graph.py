import sqlite3
import seaborn as sns
import plotly.express as px
import pandas as pd
import matplotlib.pyplot as plt
import mplcursors
from utils import *
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_FILE_PATH = PROJECT_ROOT / "database" / "master.db"

# Connect to the SQLite database
conn = sqlite3.connect(DB_FILE_PATH)
cursor = conn.cursor()

query = """

    SELECT
        a.attendance_pk,
        a.student_id,
        a.class_date,
        a.site,
        a.class_hours,
        s.last_name,
        s.first_name
    From
        Attendance a
    LEFT JOIN
        Student s
    ON
        a.student_id = s.student_id
"""

cursor.execute(query)
data = cursor.fetchall()

# Convert to DataFrame
df = pd.DataFrame(data, columns=[
    'attendance_pk', 'student_id', 'class_date', 
    'class_site', 'hours', 'last_name', 'first_name'
])

# Convert class_date to datetime format for grouping by month
df['class_date'] = pd.to_datetime(df['class_date'])

# Convert hours to discrete attendance count (assuming each 2.5 hours is one attendance)
df['attendance_count'] = (df['hours'] / 2.5).round().astype(int)

# Group by month and site, and sum attendance counts
weekly_aggregate = df.groupby([df['class_date'].dt.to_period('W'), 'class_site']).agg(
    total_attendance=('attendance_count', 'sum')
).reset_index()

# Filter out records with zero attendance
weekly_aggregate = weekly_aggregate[weekly_aggregate['total_attendance'] > 0]

# Optional: convert the period back to a timestamp for easier reading
weekly_aggregate['class_date'] = weekly_aggregate['class_date'].dt.to_timestamp()

# Plotting with Seaborn
plt.figure(figsize=(10, 6))
sns.lineplot(data=weekly_aggregate, x='class_date', y='total_attendance', hue='class_site', marker='o')

# Adding labels to each point
for line_data in weekly_aggregate.groupby('class_site'):
    site, data = line_data
    for i in range(data.shape[0]):
        plt.text(data['class_date'].iloc[i], 
                 data['total_attendance'].iloc[i] + 0.5,  # Adjust for spacing above the marker
                 f"{data['total_attendance'].iloc[i]}", 
                 horizontalalignment='center', 
                 size='small', 
                 color='black')

# Adding titles and labels
plt.title('Weekly Attendance by Site')
plt.xlabel('Week')
plt.ylabel('Total Attendance')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

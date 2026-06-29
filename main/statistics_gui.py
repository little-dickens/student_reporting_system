import sqlite3
import seaborn as sns
import pandas as pd
import matplotlib.pyplot as plt
from utils import *
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy.stats import linregress
from datetime import datetime
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator
import numpy as np
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_FILE_PATH = PROJECT_ROOT / "database" / "master.db"

def plot_attendance(site='all', frequency='Daily', date_range=('2025-06-30', '2026-07-01'), data_labels=False, prediction=False, add_regression=False, remove_outliers=False, fig=None, ax=None):
    """
    Plots attendance data based on specified parameters and returns a matplotlib figure.
    
    Parameters:
    - site (str or list): Attendance site ['NRY', 'NOL', 'BBE', 'SEU'] or 'all' for all sites). Default is 'all'.
    - frequency (str): Aggregation frequency ('Daily', 'Weekly', 'Monthly', 'Fiscal Year', or 'All'). Default is 'Weekly'.
    - date_range (tuple): Start and end date as ('YYYY-MM-DD', 'YYYY-MM-DD'). Default is the current fiscal year.
    - data_labels (bool): Whether to display labels on each data point. Default is False.
    - prediction (bool): Placeholder for future prediction functionality.
    - add_regression (bool): Whether to add a regression line for each site. Default is False.
    - remove_outliers (bool): Whether to remove outliers with `total_attendance` values of 0. Default is False.
    """
    
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
        FROM
            Attendance a
        LEFT JOIN
            Student s ON a.student_id = s.student_id
    """

    cursor.execute(query)
    data = cursor.fetchall()
    conn.close()

    # Convert to DataFrame
    df = pd.DataFrame(data, columns=[
        'attendance_pk', 'student_id', 'class_date', 'class_site', 'hours', 'last_name', 'first_name'
    ])

    # Convert class_date to datetime format
    df['class_date'] = pd.to_datetime(df['class_date'])

    # Filter by site if not 'all'
    if site != 'all':
        if isinstance(site, list):
            df = df[df['class_site'].isin(site)]
        else:
            df = df[df['class_site'] == site]

    # Filter by date range
    start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    df = df[(df['class_date'] >= start_date) & (df['class_date'] <= end_date)]

    # Convert hours to attendance count
    df['attendance_count'] = (df['hours'] / 2.5).round().astype(int)

    # Define frequency grouping
    if frequency == 'Daily':
        df['date_group'] = df['class_date']
    elif frequency == 'Weekly':
        df['date_group'] = df['class_date'].dt.to_period('W').dt.to_timestamp()
    elif frequency == 'Monthly':
        df['date_group'] = df['class_date'].dt.to_period('M').dt.to_timestamp()
    elif frequency == 'Fiscal Year':
        df['date_group'] = df['class_date'].apply(lambda x: x.replace(year=x.year if x.month < 7 else x.year + 1))
    else:
        df['date_group'] = df['class_date'].dt.to_period('A').dt.to_timestamp()

    # Group by the specified frequency and site
    aggregated_data = df.groupby(['date_group', 'class_site']).agg(
        total_attendance=('attendance_count', 'sum')
    ).reset_index()

    # Optionally remove outliers (values of 0)
    if remove_outliers:
        aggregated_data = aggregated_data[aggregated_data['total_attendance'] > 0]

    # Define unique sites for consistent color assignment
    unique_sites = aggregated_data['class_site'].unique()
    color_palette = sns.color_palette("tab10", len(unique_sites))
    color_map = {site: color for site, color in zip(unique_sites, color_palette)}

    # Use existing figure and axis if provided
    if fig is None or ax is None:
        fig, ax = plt.subplots(figsize=(10, 6))

    # Plotting with Seaborn, with explicit hue order
    sns.lineplot(data=aggregated_data, x='date_group', y='total_attendance', hue='class_site', marker='o', palette=color_map, ax=ax)

    # Set y-axis to integer values only
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))

    # Define vertical padding for data labels
    if frequency == "Daily":
        if site == 'BBE':
            vertical_padding = .1
        else:
            vertical_padding = .25
    elif frequency == "Weekly":
        vertical_padding = .5
    else:
        vertical_padding = 1.5

# Add data labels if specified
    if data_labels:
        for site_name, site_data in aggregated_data.groupby('class_site'):
            previous_label = None  # Initialize variable to track previous label

            for i in range(site_data.shape[0]):
                current_label = site_data['total_attendance'].iloc[i]

                # Only write the label if it is different from the previous one
                if previous_label is None or current_label != previous_label:
                    ax.text(site_data['date_group'].iloc[i], 
                            current_label + vertical_padding,  # Adjust for spacing above the marker
                            f"{current_label}", 
                            horizontalalignment='center', 
                            size='small', 
                            color='black')

                previous_label = current_label  # Update previous label


    # Custom legend entries for combined line and regression equation
    custom_legend_entries = []

    # Add regression lines if specified
    if add_regression:
        for site in unique_sites:
            site_data = aggregated_data[aggregated_data['class_site'] == site]
            
            # Skip regression if no valid data points after removing outliers
            if site_data['total_attendance'].count() <= 1:
                continue

            x = site_data['date_group'].apply(lambda date: date.toordinal()).values
            y = site_data['total_attendance'].values

            # Perform linear regression
            slope, intercept, _, _, _ = linregress(x, y)
            y_regression = intercept + slope * x

            # Plot regression line with the matching color from color_map
            ax.plot(site_data['date_group'], y_regression, linestyle='--', color=color_map[site])

            # Create a single legend entry with both line and regression equation
            equation = f"{site} (y = {slope:.2f}x + {intercept:.2f})"
            custom_legend_entries.append((plt.Line2D([0], [0], color=color_map[site], marker='o', linestyle='-'), equation))

    # Apply the custom legend, replacing the default
    if add_regression is True:
        ax.legend(handles=[entry[0] for entry in custom_legend_entries], 
                labels=[entry[1] for entry in custom_legend_entries],
                title="Sites & Regression")
    else:
        ax.legend(handles=[entry[0] for entry in custom_legend_entries], 
                labels=[entry[1] for entry in custom_legend_entries],
                title="Sites")        

    # Titles and labels
    ax.set_title(f'{frequency} Attendance by Site')
    ax.set_xlabel(frequency)
    ax.set_ylabel('Total Attendance')
    ax.tick_params(axis='x', rotation=45)

    # Set x-axis formatting for Monthly frequency
    if frequency == 'Monthly':
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b '%y"))

    plt.tight_layout()

    # Calculate statistics
    attendance_values = aggregated_data['total_attendance']
    stats = {
        "x̄": attendance_values.mean(),
        "M": attendance_values.median(),
        "Mo": attendance_values.mode().iloc[0] if not attendance_values.mode().empty else np.nan,
        "s": attendance_values.std(),
        "s\u00B2": attendance_values.var(),
        "min": attendance_values.min(),
        "max": attendance_values.max(),
        "R": attendance_values.max() - attendance_values.min(),
        "IQR": attendance_values.quantile(0.75) - attendance_values.quantile(0.25),
        "10th": attendance_values.quantile(0.1),
        "25th": attendance_values.quantile(0.25),
        "75th": attendance_values.quantile(0.75),
        "90th": attendance_values.quantile(0.9),
        "Σ": attendance_values.sum(),
    }
    
    # Return the figure object for embedding in the GUI
    return fig, stats

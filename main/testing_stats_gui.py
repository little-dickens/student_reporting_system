import sqlite3
import seaborn as sns
import pandas as pd
import matplotlib.pyplot as plt
from utils import *
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
from matplotlib.ticker import MultipleLocator
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_FILE_PATH = PROJECT_ROOT / "database" / "master.db"

def plot_testing(
        site='all',
        frequency='Fiscal Year',
        date_range=('2024-07-01', '2025-06-30'), 
        data_labels=True,
        combined_scaled=False,
        reading_scaled=False,
        writing_scaled=False,
        msg_total=False,
        msg_percent=False,
        combined_change_percent=False,
        reading_change_percent=False,
        writing_change_percent=False,
        fig=None, 
        ax=None
    ):

    # Connect to the SQLite database
    conn = sqlite3.connect(DB_FILE_PATH)
    cursor = conn.cursor()

    query = """
        SELECT
            site,
            date,
            pre_or_post,
            level,
            form,
            reading_raw,
            reading_scaled,
            reading_nrs,
            writing_raw,
            writing_scaled,
            writing_nrs,
            writing_folio_answers,
            combined,
            got_msg
        FROM
            Testing
    """

    cursor.execute(query)
    data = cursor.fetchall()
    conn.close()

    # Use existing figure and axis if provided
    if fig is None or ax is None:
        fig, ax = plt.subplots(figsize=(10, 6))

    # Convert to DataFrame
    df = pd.DataFrame(data, columns=[
        'site', 'date', 'pre_or_post', 'level', 'form', 'reading_raw', 'reading_scaled', 'reading_nrs', 'writing_raw', 'writing_scaled', 'writing_nrs', 'writing_folio_answers', 'combined', 'got_msg'
    ])

    # Convert test date to datetime format
    df['date'] = pd.to_datetime(df['date'])

    # Step up + nrs levels
    def convert_values(value):
        if isinstance(value, str) and value.endswith('+'):
            return int(value[:-1]) + 1
        return value

    df['reading_nrs'] = df['reading_nrs'].apply(convert_values)
    df['writing_nrs'] = df['writing_nrs'].apply(convert_values)

    # Filter by date range
    start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    df = df[(df['date'] >= start_date) & (df['date'] <= end_date)]

    # Filter by site if not 'all'
    if site != 'all':
        if isinstance(site, list):
            df = df[df['site'].isin(site)]
        else:
            df = df[df['site'] == site]

    # Introduce 'Total' site that accumulates data from selected sites
    df_total = df.copy()
    df_total['site'] = 'Total'
    df = pd.concat([df, df_total], ignore_index=True)

    # Filter out locator tests
    df = df[df['level'].isin([1, 2, 3, 4])]

    # Map got_msg values to 'Yes'/'No'
    df['got_msg'] = df['got_msg'].fillna(0).astype(int)
    df['got_msg_label'] = df['got_msg'].map({0: 'No', 1: 'Yes'})

    # Prepare color palette
    unique_sites = df['site'].unique()
    site_palette = sns.color_palette("tab10", len(unique_sites))
    site_color_map = {site: color for site, color in zip(unique_sites, site_palette)}
    got_msg_palette = {'No': 'red', 'Yes': 'green'}  # Palette for got_msg_label

    # Plotting
    if combined_scaled or reading_scaled or writing_scaled:
        # Determine which column to plot
        if combined_scaled:
            value_col = 'combined'
            x_label = 'Combined Scaled'
        elif reading_scaled:
            value_col = 'reading_scaled'
            x_label = 'Reading Scaled'
        elif writing_scaled:
            value_col = 'writing_scaled'
            x_label = 'Writing Scaled'

        sns.boxplot(
            data=df, 
            x=value_col, 
            y='site', 
            palette=site_color_map, 
            width=0.6,
            dodge=False,
            ax=ax
        )

        # Customize axis labels and title
        ax.set_xlabel(x_label)
        ax.set_ylabel('Site')
        ax.set_title(f'Box Plot of {x_label} by Site')

    elif msg_total or msg_percent:
        # Filter for Post tests only
        df_msg = df[df['pre_or_post'] == 'Post']

        # Ensure both 'Yes' and 'No' categories are present
        got_msg_order = ['No', 'Yes']

        if msg_total:
            # Plot total counts
            sns.countplot(
                data=df_msg, 
                x='site', 
                hue='got_msg_label', 
                palette=got_msg_palette, 
                order=sorted(unique_sites),
                hue_order=got_msg_order,
                ax=ax
            )
            y_label = 'Total Count'
            title = 'MSG Total by Site'

        elif msg_percent:
            # Calculate percentages
            total_counts = df_msg.groupby(['site'])['got_msg_label'].count().reset_index(name='total')
            msg_counts = df_msg.groupby(['site', 'got_msg_label'])['got_msg_label'].count().reset_index(name='count')
            merged_df = pd.merge(msg_counts, total_counts, on='site')
            merged_df['percent'] = (merged_df['count'] / merged_df['total']) * 100

            # Ensure all combinations are present
            all_combinations = pd.MultiIndex.from_product([unique_sites, got_msg_order], names=['site', 'got_msg_label']).to_frame(index=False)
            merged_df = pd.merge(all_combinations, merged_df, on=['site', 'got_msg_label'], how='left')
            merged_df['percent'] = merged_df['percent'].fillna(0)

            # Plot percentages
            sns.barplot(
                data=merged_df, 
                x='site', 
                y='percent', 
                hue='got_msg_label', 
                palette=got_msg_palette, 
                order=sorted(unique_sites),
                hue_order=got_msg_order,
                ax=ax
            )
            y_label = 'Percentage (%)'
            title = 'MSG Percentage by Site'

            # Show percentages on top of bars if data_labels is True
            if data_labels:
                for container in ax.containers:
                    ax.bar_label(container, fmt='%.1f%%', label_type='edge')
                # Iterate over the bars and the data
                for i, (bar, count) in enumerate(zip(ax.patches, merged_df['count'])):
                    print(i, bar, count)
                    # The bars are plotted in order of sites and hue levels
                    height = bar.get_height()
                    ax.text(
                        bar.get_x() + bar.get_width()/2,
                        height / 2,
                        f'{int(count)}',
                        ha='center', va='center', color='white', fontsize=10
                    )
        # Customize axis labels and title
        ax.set_xlabel('Site')
        ax.set_ylabel(y_label)
        ax.set_title(title)
        ax.legend(title='MSG Received', loc='upper right')

    print(merged_df)
    print(merged_df['count'])

    # Adjust plot aesthetics
    # ax.grid(True, linestyle='--', alpha=0.7)

    plt.tight_layout()

    # Return the figure object for embedding in Tkinter
    return fig


# Display the plot
plt.show()
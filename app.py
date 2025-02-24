import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from collections import Counter

st.header("Turnaround Time Analysis")

# File upload
uploaded_file = st.file_uploader("Upload an Excel file:", type=["xlsx"])

if uploaded_file is not None:
    raw_data = pd.read_excel(uploaded_file)
    
    # Creating a cleaned DataFrame
    turnaround_data = pd.DataFrame()
    turnaround_data["Transaction ID"] = raw_data.iloc[:, 0]
    turnaround_data["Sale Date"] = pd.to_datetime(raw_data.iloc[:, 1], errors='coerce')
    turnaround_data["Date of Execution (DOE)"] = pd.to_datetime(raw_data.iloc[:, 2], errors='coerce')
    turnaround_data["COI Upload Date"] = pd.to_datetime(raw_data.iloc[:, 3], errors='coerce')

    # Extract year and month for filtering
    turnaround_data['Sale Year'] = turnaround_data["Sale Date"].dt.year
    turnaround_data['Sale Month'] = turnaround_data["Sale Date"].dt.month

    # Calculate turnaround times (TAT)
    turnaround_data["TAT Sale to DOE"] = (turnaround_data["Date of Execution (DOE)"] - turnaround_data["Sale Date"]).dt.days
    turnaround_data["TAT DOE to COI Upload"] = (turnaround_data["COI Upload Date"] - turnaround_data["Date of Execution (DOE)"]).dt.days

    # Year selection dropdown
    available_years = sorted(turnaround_data['Sale Year'].dropna().unique())
    selected_year = st.selectbox("Select Year", ["All"] + list(map(int, available_years)))

    # Filter data based on user selection
    filtered_data = turnaround_data.copy()
    if selected_year != "All":
        filtered_data = filtered_data[filtered_data["Sale Year"] == int(selected_year)]
    
    # Month selection dropdown
    available_months = sorted(filtered_data['Sale Month'].dropna().unique())
    selected_month = st.selectbox("Select Month", ["All"] + list(map(int, available_months)))

    if selected_month != "All":
        filtered_data = filtered_data[filtered_data["Sale Month"] == int(selected_month)]

    if filtered_data.empty:
        st.warning("No data available for the selected filters.")
    else:
        # Group and count TAT pairs
        tat_pairs_count = Counter(zip(filtered_data["TAT Sale to DOE"], filtered_data["TAT DOE to COI Upload"]))
        tat_summary = pd.DataFrame(tat_pairs_count.items(), columns=["TAT Pair", "Count"])
        tat_summary[["TAT Sale to DOE", "TAT DOE to COI Upload"]] = pd.DataFrame(tat_summary["TAT Pair"].tolist(), index=tat_summary.index)
        tat_summary.drop(columns=["TAT Pair"], inplace=True)
        tat_summary["Total TAT (Sale to COI)"] = tat_summary["TAT Sale to DOE"] + tat_summary["TAT DOE to COI Upload"]
        tat_summary = tat_summary.sort_values(by=["Total TAT (Sale to COI)"], ascending=True)

        # Plotting
        fig, ax = plt.subplots(figsize=(12, 16))
        y_positions = range(len(tat_summary))

        bars_sale_to_doe = ax.barh(y_positions, tat_summary["TAT Sale to DOE"], color='#876FD4', label="Sale to DOE")
        bars_doe_to_coi = ax.barh(y_positions, tat_summary["TAT DOE to COI Upload"], color='#B6A6E9', left=tat_summary["TAT Sale to DOE"], label="DOE to COI Upload")

        # Add labels inside the bars
        for bar, days in zip(bars_sale_to_doe, tat_summary["TAT Sale to DOE"]):
            ax.text(bar.get_width() / 2, bar.get_y() + bar.get_height() / 2, f"{days}", 
                    ha='center', va='center', color='white', fontsize=10, fontweight='bold')

        for bar, days, start in zip(bars_doe_to_coi, tat_summary["TAT DOE to COI Upload"], tat_summary["TAT Sale to DOE"]):
            ax.text(start + bar.get_width() / 2, bar.get_y() + bar.get_height() / 2, f"{days}", 
                    ha='center', va='center', color='white', fontsize=10, fontweight='bold')

        # Display count beside each bar group
        for y, total_tat, count in zip(y_positions, tat_summary["Total TAT (Sale to COI)"], tat_summary["Count"]):
            ax.text(total_tat + 1, y, f"({count})", ha='left', va='center', fontsize=12, fontweight='bold', color='black')

        # Formatting the plot
        ax.set_yticks([])
        ax.set_xlabel("Turnaround Time (Days)", fontsize=14, fontname='Arvo', fontstyle='italic')
        ax.set_ylabel("Turnaround Time Groups (Sale to DOE, DOE to COI)", fontsize=14, fontname='Arvo', fontstyle='italic')
        ax.set_title("Turnaround Time Analysis (Grouped & Stacked Horizontal Bar Chart)", fontsize=16, fontname='Georgia', fontweight='bold')
        ax.legend()
        ax.spines['left'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        st.pyplot(fig)

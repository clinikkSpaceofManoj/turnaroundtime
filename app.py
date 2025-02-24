import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from collections import Counter

st.header("Turn Around Time Analysis")

data = st.file_uploader("Provide the file: ", type=["xlsx"])

if data is not None:
    df = pd.read_excel(data)
    df["Sale date"] = pd.to_datetime(df["Sale date"], errors='coerce')
    df["DOE"] = pd.to_datetime(df["DOE"], errors='coerce')
    df["COI Uploaded Date"] = pd.to_datetime(df["COI Uploaded Date"], errors='coerce')
    df['year'] = df["Sale date"].dt.year
    df['month'] = df["Sale date"].dt.month
    
    df["TATSaleToDOE"] = (df["DOE"] - df["Sale date"]).dt.days
    df["TATDOEToCOI"] = (df["COI Uploaded Date"] - df["DOE"]).dt.days
    
    st.dataframe(df)
    
    years = sorted(df['year'].dropna().unique())
    months = sorted(df['month'].dropna().unique())
    
    selected_year = st.slider("Select Year", min_value=min(years), max_value=max(years), value=min(years))
    selected_month = st.slider("Select Month", min_value=min(months), max_value=max(months), value=min(months))
    
    if selected_year and selected_month:
        df = df[(df["month"] == selected_month) & (df["year"] == selected_year)]
    
    tatCounts = Counter(zip(df["TATSaleToDOE"], df["TATDOEToCOI"]))
    groupedData = pd.DataFrame(tatCounts.items(), columns=["TAT Pair", "Count"])
    groupedData[["TATSaleToDOE", "TATDOEToCOI"]] = pd.DataFrame(groupedData["TAT Pair"].to_list(), index=groupedData.index)
    groupedData = groupedData.drop(columns=["TAT Pair"])
    groupedData["TATSaleToCOI"] = groupedData["TATSaleToDOE"] + groupedData["TATDOEToCOI"]
    groupedData = groupedData.sort_values(by=["TATSaleToCOI"], ascending=True)
    
    st.dataframe(groupedData)
    
    fig, ax = plt.subplots(figsize=(12, 7))
    y_positions = range(len(groupedData))
    
    bars1 = ax.barh(y_positions, groupedData["TATSaleToDOE"], color='green', label="Sale to DOE")
    bars2 = ax.barh(y_positions, groupedData["TATDOEToCOI"], color='red', left=groupedData["TATSaleToDOE"], label="DOE to COI Upload")
    
    for bar, days in zip(bars1, groupedData["TATSaleToDOE"]):
        ax.text(bar.get_width() / 2, bar.get_y() + bar.get_height() / 2, f"{days}", ha='center', va='center', color='white', fontsize=10, fontweight='bold')
    
    for bar, days, start in zip(bars2, groupedData["TATDOEToCOI"], groupedData["TATSaleToDOE"]):
        ax.text(start + bar.get_width() / 2, bar.get_y() + bar.get_height() / 2, f"{days}", ha='center', va='center', color='white', fontsize=10, fontweight='bold')
    
    for y, total_tat, count in zip(y_positions, groupedData["TATSaleToCOI"], groupedData["Count"]):
        ax.text(total_tat + 1, y, f"({count})", ha='left', va='center', fontsize=12, fontweight='bold', color='black')
    
    ax.set_yticks(y_positions)
    ax.set_yticklabels([f"{row.TATSaleToDOE},{row.TATDOEToCOI}" for _, row in groupedData.iterrows()])
    ax.set_xlabel("Turnaround Time (Days)")
    ax.set_ylabel("Turnaround Time Groups (Sale to DOE, DOE to COI)")
    ax.set_title("Turnaround Time Analysis (Grouped & Stacked Horizontal Bar Chart)")
    ax.legend()
    
    st.pyplot(fig)

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Suppressing the warning
st.set_option('deprecation.showPyplotGlobalUse', False)

pwd = os.getcwd()

# Load data
@st.cache_data
def load_data():
    return pd.read_excel(os.path.join(pwd, "melted_output.xlsx"))

df = load_data()

# App title
st.title("Migrants Data Visualization")

# Visualization 1: Time Series Plot
st.header("Time Series Plot of Migrants from a Specific Country to States over the Years")
selected_states = st.multiselect("Select States", df["State"].unique().tolist(), default=df["State"].unique().tolist())
filtered_df = df[df["State"].isin(selected_states)]
fig1, ax1 = plt.subplots(figsize=(14, 6))
sns.lineplot(data=filtered_df, x="Year", y="Count of Migrants", hue="State", ax=ax1)
ax1.set_title("Trend of Migrants over the Years")
ax1.set_xticklabels(filtered_df["Year"].unique().tolist(), rotation=45)
st.pyplot(fig1)

# Visualization 2: Bar Chart
st.header("Migrants from Different Countries")
selected_year = st.selectbox("Select Year", df["Year"].unique().tolist())
year_filtered_df = df[df["Year"] == selected_year]
sorted_df = year_filtered_df.groupby("Country of birth(c)")["Count of Migrants"].sum().sort_values(ascending=False).head(10)
fig2, ax2 = plt.subplots(figsize=(14, 6))
sorted_df.plot(kind='bar', ax=ax2)
ax2.set_title("Migrants from Different Countries in " + selected_year)
st.pyplot(fig2)

# Visualization 3: Heatmap
st.header("Heatmap of Migrants Distribution Across States")
heatmap_data = df.pivot_table(values="Count of Migrants", index="State", columns="Year")
fig3, ax3 = plt.subplots(figsize=(14, 6))
sns.heatmap(heatmap_data, cmap="YlGnBu", ax=ax3)
st.pyplot(fig3)

# Visualization 4: Pie Chart
st.header("Proportion of Migrants from Different Countries")
selected_state = st.selectbox("Select State for Pie Chart", df["State"].unique().tolist())
state_filtered_df = df[df["State"] == selected_state]
pie_data = state_filtered_df.groupby("Country of birth(c)")["Count of Migrants"].sum().sort_values(ascending=False).head(10)
fig4, ax4 = plt.subplots(figsize=(10, 10))
ax4.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=140)
ax4.set_title("Proportion of Migrants from Different Countries to " + selected_state)
st.pyplot(fig4)
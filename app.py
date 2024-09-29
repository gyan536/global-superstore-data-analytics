# Importing necessary libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st

# Setting visualization styles
sns.set(style="whitegrid")

# Load the data
@st.cache
def load_data():
    file_path = 'global_superstore_2016.xlsx'  # Update path as necessary
    xls = pd.ExcelFile(file_path)
    orders_df = pd.read_excel(xls, sheet_name='Orders')
    return orders_df

# Load the data
orders_df = load_data()

# Data Cleaning
orders_df = orders_df.drop(columns=['Row ID', 'Order ID'])  # Drop irrelevant columns
orders_df['Order Date'] = pd.to_datetime(orders_df['Order Date'])
orders_df['Ship Date'] = pd.to_datetime(orders_df['Ship Date'])
orders_df['Order Month'] = orders_df['Order Date'].dt.month_name()
orders_df['Order Year'] = orders_df['Order Date'].dt.year
orders_df['Sales Growth'] = orders_df['Sales'].pct_change().fillna(0)

# Streamlit Title
st.title("Global Superstore Sales Dashboard")

# Data Overview
if st.checkbox("Show Data Overview"):
    st.subheader("Data Overview")
    st.write(orders_df.head())
    st.write(f"Dataset Shape: {orders_df.shape}")
    st.write(f"Missing Values:\n{orders_df.isnull().sum()}")

# Summary Statistics
if st.checkbox("Show Summary Statistics"):
    st.subheader("Summary Statistics")
    st.write(orders_df.describe())

# Total Sales and Profit
total_sales = orders_df['Sales'].sum()
total_profit = orders_df['Profit'].sum()
st.subheader("Total Sales and Profit")
st.write(f"Total Sales: ${total_sales:,.2f}")
st.write(f"Total Profit: ${total_profit:,.2f}")

# Sales by Category
st.subheader("Sales by Category")
category_sales = orders_df.groupby('Category')['Sales'].sum().reset_index()
fig1 = plt.figure(figsize=(10, 6))
sns.barplot(data=category_sales, x='Category', y='Sales', palette='viridis')
plt.title('Total Sales by Category')
plt.xticks(rotation=45)
st.pyplot(fig1)

# Profit by Category
st.subheader("Profit by Category")
category_profit = orders_df.groupby('Category')['Profit'].sum().reset_index()
fig2 = plt.figure(figsize=(10, 6))
sns.barplot(data=category_profit, x='Category', y='Profit', palette='viridis')
plt.title('Total Profit by Category')
plt.xticks(rotation=45)
st.pyplot(fig2)

# Sales by Region
st.subheader("Sales by Region")
region_sales = orders_df.groupby('Region')['Sales'].sum().reset_index()
fig3 = plt.figure(figsize=(10, 6))
sns.barplot(data=region_sales, x='Region', y='Sales', palette='plasma')
plt.title('Total Sales by Region')
plt.xticks(rotation=45)
st.pyplot(fig3)

# Monthly Sales Analysis
st.subheader("Monthly Sales Analysis")
monthly_sales = orders_df.groupby(['Order Year', 'Order Month'])['Sales'].sum().reset_index()
fig4 = plt.figure(figsize=(12, 6))
sns.lineplot(data=monthly_sales, x='Order Month', y='Sales', hue='Order Year', marker='o')
plt.title('Monthly Sales Over Years')
plt.xticks(rotation=45)
st.pyplot(fig4)

# Show Correlation Heatmap
if st.checkbox("Show Correlation Heatmap"):
    st.subheader("Correlation Heatmap")
    # Selecting only numeric columns for correlation analysis
    numeric_df = orders_df.select_dtypes(include=[np.number])
    
    if numeric_df.empty:
        st.warning("No numeric columns available for correlation analysis.")
    else:
        correlation_matrix = numeric_df.corr()
        fig5 = plt.figure(figsize=(10, 6))
        sns.heatmap(correlation_matrix, annot=True, fmt=".2f", cmap='coolwarm', square=True)
        plt.title("Correlation Matrix")
        st.pyplot(fig5)

# Saving the cleaned DataFrame
if st.button("Save Cleaned Data"):
    cleaned_file_path = 'cleaned_global_superstore_2016.xlsx'
    orders_df.to_excel(cleaned_file_path, index=False)
    st.success(f"Cleaned data saved to: {cleaned_file_path}")

# Footer
st.markdown("### Created with ❤️ using Streamlit")


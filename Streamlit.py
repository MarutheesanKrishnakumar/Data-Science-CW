import streamlit as st
import plotly.express as px
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')
import openpyxl

Cleaned_data = pd.read_excel('Global Superstore lite.xlsx')
st.set_page_config(page_title="Global Superstore", page_icon=":chart_with_upwards_trend:",layout="wide")

st.title(" :chart_with_upwards_trend: Minger Sales Analysis")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)


# Filtering date
col1, col2 = st.columns((2))
Cleaned_data["Order Date"] = pd.to_datetime(Cleaned_data["Order Date"])

# Getting the min and max date
start_date = Cleaned_data["Order Date"].min()
end_date = Cleaned_data["Order Date"].max()

with col1:
    date1 = st.date_input("Start Date", start_date)

with col2:
    date2 = st.date_input("End Date", end_date)

Cleaned_data = Cleaned_data[(Cleaned_data["Order Date"] >= pd.to_datetime(date1)) & (Cleaned_data["Order Date"] <= pd.to_datetime(date2))].copy()

SubCategory_df = Cleaned_data.groupby(by = ["Sub-Category"], as_index= False)["Sales"].sum()
with col1:
    st.subheader("Sales by Sub-Category")
    fig = px.bar(SubCategory_df, x = "Sub-Category", y = "Sales", text = ['${:,.2f}'.format(x) for x in SubCategory_df["Sales"]],
                 template = "seaborn")
    st.plotly_chart(fig,use_container_width=True, height = 200)

with col2:
    st.subheader("Sales by Region")
    fig = px.pie(Cleaned_data, values = "Sales", names = "Region", hole = 0.5)
    fig.update_traces(text = Cleaned_data["Region"], textposition = "outside")
    st.plotly_chart(fig,use_container_width=True)



chart1, chart2 = st.columns((2))
with chart1:
    st.subheader('Sales by Ship mode')
    fig = px.pie(Cleaned_data, values = "Profit", names = "Ship Mode", template = "plotly_dark")
    fig.update_traces(text = Cleaned_data["Ship Mode"], textposition = "inside")
    st.plotly_chart(fig,use_container_width=True)

with chart2:
    st.subheader('Category wise Sales')
    fig = px.pie(Cleaned_data, values = "Sales", names = "Category", template = "gridon")
    fig.update_traces(text = Cleaned_data["Category"], textposition = "inside")
    st.plotly_chart(fig,use_container_width=True)


Cleaned_data["Quarterly"] = Cleaned_data["Order Date"].dt.to_period("Q")

# Filter to include only January, April, July, and October quarters
filtered_df = Cleaned_data[Cleaned_data["Order Date"].dt.month.isin([1, 4, 7, 10])]

st.subheader('Quarterly Sales')

# Grouping Data by Quarters
linechart = pd.DataFrame(filtered_df.groupby(filtered_df["Quarterly"].dt.to_timestamp(freq='Q'))["Sales"].sum()).reset_index()

# Creating Line Chart
fig2 = px.line(linechart, x="Quarterly", y="Sales", labels={"Sales": "Amount"}, height=500, width=1000, template="gridon")

# Displaying the Chart
st.plotly_chart(fig2, use_container_width=True)

chart1, chart2 = st.columns((2))
with chart1:    
    # Box plot
    fig_box = px.box(Cleaned_data, x='Category', y='Profit', title='Profit Obtained by Category', width=700, height=800)
    st.plotly_chart(fig_box)

with chart2:
    # Group by Category and sum up Quantity
    df_sum = Cleaned_data.groupby('Sub-Category')['Quantity'].sum().reset_index()
    fig = px.bar(df_sum, x='Sub-Category', y='Quantity', title='Most Frequently Purchased Products',width=700, height=800)
    st.plotly_chart(fig)

chart1, chart2 = st.columns((2))
with chart1: 
    heatmap_data = Cleaned_data.groupby(['Region', 'Sub-Category']).agg({'Sales': 'sum'}).reset_index()
    fig_heatmap = px.imshow(heatmap_data.pivot(index="Region", columns="Sub-Category", values="Sales"), 
                        labels=dict(x="Sub-Category", y="Region", color="Sales"),
                        title='Sales Heatmap by Sub-Category and Region',
                        width=900, height=600)  # Specify width and height
    st.plotly_chart(fig_heatmap)

with chart2:
    fig = px.area(Cleaned_data, x='Region', y='Profit', title='Profit by Region',width=600, height=600)
    st.plotly_chart(fig)


# Scatter plot
fig_scatter = px.scatter(Cleaned_data, x='Sales', y='Profit', color='Category', title='Sales vs Profit', width=1000, height=600)
st.plotly_chart(fig_scatter)



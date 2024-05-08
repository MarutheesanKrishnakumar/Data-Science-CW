import streamlit as st
import plotly.express as px
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')
import openpyxl

df = pd.read_excel('Global Superstore lite.xlsx')
st.set_page_config(page_title="Global Superstore", page_icon=":chart_with_upwards_trend:",layout="wide")

st.title(" :chart_with_upwards_trend: Global SuperStore Sales Analysis")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)



col1, col2 = st.columns((2))
df["Order Date"] = pd.to_datetime(df["Order Date"])

# Getting the min and max date 
startDate = pd.to_datetime(df["Order Date"]).min()
endDate = pd.to_datetime(df["Order Date"]).max()

with col1:
    date1 = pd.to_datetime(st.date_input("Start Date", startDate))

with col2:
    date2 = pd.to_datetime(st.date_input("End Date", endDate))

df = df[(df["Order Date"] >= date1) & (df["Order Date"] <= date2)].copy()


st.sidebar.header("Choose your filter: ")
# Sidebar Filters



# Create for Region
region = st.sidebar.multiselect("Pick your Region", df["Region"].unique())
if not region:
    df2 = df.copy()
else:
    df2 = df[df["Region"].isin(region)]

# Create for Country
country = st.sidebar.multiselect("Pick the Country", df2["Country"].unique())
if not country:
    df3 = df2.copy()
else:
    df3 = df2[df2["Country"].isin(country)]

# Create for State
state = st.sidebar.multiselect("Pick the State", df3["State"].unique())
if not state:
    df4 = df3.copy()
else:
    df4 = df3[df3["State"].isin(state)]

# Create for City
city = st.sidebar.multiselect("Pick the City", df4["City"].unique())
if not city:
    df5 = df4.copy()
else:
    df5 = df4[df4["City"].isin(city)]

# Filter the data based on Region, Country, State, and City

if not region and not country and not state and not city:
    filtered_df = df
elif not country and not state and not city:
    filtered_df = df[df["Region"].isin(region)]
elif not region and not state and not city:
    filtered_df = df[df["Country"].isin(country)]
elif state and city and not region and not country:
    filtered_df = df4[df4["State"].isin(state) & df4["City"].isin(city)]
elif region and city and not country and not state:
    filtered_df = df4[df4["Region"].isin(region) & df4["City"].isin(city)]
elif region and state and not city and not country:
    filtered_df = df4[df4["Region"].isin(region) & df4["State"].isin(state)]
elif city and not country and not state:
    filtered_df = df4[df4["City"].isin(city)]
elif country and not state and not city:
    filtered_df = df4[df4["Country"].isin(country)]
else:
    filtered_df = df5[df5["Region"].isin(region) & df5["Country"].isin(country) & df5["State"].isin(state) & df5["City"].isin(city)]

SubCategory_df = filtered_df.groupby(by = ["Sub-Category"], as_index= False)["Sales"].sum()

with col1:
    st.subheader("Sales by Sub-Category")
    fig = px.bar(SubCategory_df, x = "Sub-Category", y = "Sales", text = ['${:,.2f}'.format(x) for x in SubCategory_df["Sales"]],
                 template = "seaborn")
    st.plotly_chart(fig,use_container_width=True, height = 200)

with col2:
    st.subheader("Sales by Region")
    fig = px.pie(filtered_df, values = "Sales", names = "Region", hole = 0.5)
    fig.update_traces(text = filtered_df["Region"], textposition = "outside")
    st.plotly_chart(fig,use_container_width=True)



chart1, chart2 = st.columns((2))
with chart1:
    st.subheader('Sales by Ship mode')
    fig = px.pie(filtered_df, values = "Profit", names = "Ship Mode", template = "plotly_dark")
    fig.update_traces(text = filtered_df["Ship Mode"], textposition = "inside")
    st.plotly_chart(fig,use_container_width=True)

with chart2:
    st.subheader('Category wise Sales')
    fig = px.pie(filtered_df, values = "Sales", names = "Category", template = "gridon")
    fig.update_traces(text = filtered_df["Category"], textposition = "inside")
    st.plotly_chart(fig,use_container_width=True)


filtered_df["Quarterly"] = filtered_df["Order Date"].dt.to_period("Q")

# Filter to include only January, April, July, and October quarters
filtered_df = filtered_df[filtered_df["Order Date"].dt.month.isin([1, 4, 7, 10])]

st.subheader('Quarterly Sales')

# Grouping Data by Quarters
linechart = pd.DataFrame(filtered_df.groupby(filtered_df["Quarterly"].dt.to_timestamp(freq='Q'))["Sales"].sum()).reset_index()

# Creating Line Chart
fig2 = px.line(linechart, x="Quarterly", y="Sales", labels={"Sales": "Amount"}, height=500, width=1000, template="gridon")

# Displaying the Chart
st.plotly_chart(fig2, use_container_width=True)

with st.expander("View Data of TimeSeries:"):
    st.write(linechart.T.style.background_gradient(cmap="Blues"))
    csv = linechart.to_csv(index=False).encode("utf-8")
    st.download_button('Download Data', data = csv, file_name = "TimeSeries.csv", mime ='text/csv')

 

chart1, chart2 = st.columns((2))
with chart1:
    # Group by Category and sum up Quantity
    df_sum = filtered_df.groupby('Sub-Category')['Quantity'].sum().reset_index()
    fig = px.bar(df_sum, x='Sub-Category', y='Quantity', title='Most Frequently Purchased Products',width=500, height=400)
    st.plotly_chart(fig)

with chart2:
    fig = px.area(filtered_df, x='Region', y='Profit', title='Profit by Region',width=500, height=400)
    st.plotly_chart(fig)

scatter_data = px.scatter(filtered_df, x="Sales", y="Profit", size="Quantity")

# Update layout
scatter_data.update_layout(
    title="Relationship between Sales and Profits using Scatter Plot.",
    titlefont=dict(size=20),
    xaxis=dict(title="Sales", titlefont=dict(size=19)),
    yaxis=dict(title="Profit", titlefont=dict(size=19))
)

# Display scatter plot
st.plotly_chart(scatter_data, use_container_width=True)

# Download orginal DataSet
csv = df.to_csv(index = False).encode('utf-8')
st.download_button('Download Data', data = csv, file_name = "Data.csv",mime = "text/csv")


import streamlit as st
import plotly.express as px
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import seaborn as sns
import matplotlib.pyplot as plt

Cleaned_data = pd.read_excel('Global Superstore lite.xlsx')
st.set_page_config(page_title="Global Superstore", page_icon=":chart_with_upwards_trend:",layout="wide")

st.title(" :chart_with_upwards_trend: Minger Sales Analysis")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)


# Filtering date
chart1, chart2 = st.columns((2))
Cleaned_data["Order Date"] = pd.to_datetime(Cleaned_data["Order Date"])
# Getting the min and max date
start_date = Cleaned_data["Order Date"].min()
end_date = Cleaned_data["Order Date"].max()

with chart1:
    date1 = st.date_input("Start Date", start_date)

with chart2:
    date2 = st.date_input("End Date", end_date)
Cleaned_data = Cleaned_data[(Cleaned_data["Order Date"] >= pd.to_datetime(date1)) & (Cleaned_data["Order Date"] <= pd.to_datetime(date2))].copy()

# Creating two chart layout
chart1, chart2 = st.columns(2)
# Summary Card for Total Profit
total_profit = Cleaned_data['Profit'].sum()
chart1.metric("Total Profit", f"${total_profit:,.2f}")

# Summary Card for Total Profit
total_sales = Cleaned_data['Sales'].sum()
chart2.metric("Total Sales", f"${total_sales:,.2f}")


SubCategory_df = Cleaned_data.groupby(by = ["Sub-Category"], as_index= False)["Sales"].sum()
with chart1:
    st.subheader("Sales by Sub-Category")
    fig = px.bar(SubCategory_df, x = "Sub-Category", y = "Sales", text = ['${:,.2f}'.format(x) for x in SubCategory_df["Sales"]],
                 template = "seaborn")
    st.plotly_chart(fig,use_container_width=True, height = 200)

with chart2:
    st.subheader("Sales by Segment")
    fig = px.pie(Cleaned_data, values = "Sales", names = "Segment", hole = 0.5)
    fig.update_traces(text = Cleaned_data["Segment"], textposition = "outside")
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
    # Calculate Quantity by ship mode
    ShipMode_df = Cleaned_data.groupby(by="Ship Mode", as_index=False)["Quantity"].sum()
    # Create column chart using Plotly Express
    fig = px.bar(ShipMode_df, x="Ship Mode", y="Quantity",
                 text=['${:,.2f}'.format(x) for x in ShipMode_df["Quantity"]])  # Format Quantity as Count
    fig.update_traces(textposition='outside')  # Show text outside the bars
    fig.update_layout(title='Products Shipping Mode', xaxis_title='Ship Mode', yaxis_title='Quantity')
    st.plotly_chart(fig)

with chart2:
    Market_Sales = Cleaned_data['Market'].value_counts().reset_index()
    Market_Sales.columns = ['Market', 'Sales']
    fig_pie = px.pie(Market_Sales, values='Sales', names='Market', title='Sales by Market')
    st.plotly_chart(fig_pie, use_container_width=True)


chart1, chart2 = st.columns((2))
with chart1: 
    heatmap_data = Cleaned_data.groupby(['Region', 'Sub-Category']).agg({'Sales': 'sum'}).reset_index()
    fig_heatmap = px.imshow(heatmap_data.pivot(index="Region", columns="Sub-Category", values="Sales"), 
                        labels=dict(x="Sub-Category", y="Region", color="Sales"),
                        title='Sales Heatmap by Sub-Category and Region',
                        width=500, height=600)  # Specify width and height
    st.plotly_chart(fig_heatmap)

with chart2:
    # Scatter plot
    fig_scatter = px.scatter(Cleaned_data, x='Sales', y='Profit', color='Category', title='Sales vs Profit', width=700, height=600)
    st.plotly_chart(fig_scatter)


chart1, chart2 = st.columns((2))
with chart1:    
    # Box plot
    fig_box = px.box(Cleaned_data, x='Category', y='Profit', title='Profit Obtained by Category', width=700, height=1000)
    st.plotly_chart(fig_box)

with chart2:
    fig = px.area(Cleaned_data, x='Region', y='Profit', title='Profit by Region',width=600, height=1000)
    st.plotly_chart(fig)


#MBA 
from mlxtend.frequent_patterns import apriori
from mlxtend.frequent_patterns import association_rules

# Read the cleaned dataset
Minger_Cleaned_data = Cleaned_data

# Group by 'Order ID' and 'Sub-Category' and count the occurrences of each combination
basket = (Minger_Cleaned_data.groupby(['Order ID', 'Sub-Category'])['Row ID']
          .count().unstack().reset_index().fillna(0)
          .set_index('Order ID'))

# Convert the occurrence counts to binary values (0 or 1)
basket_sets = basket.applymap(lambda x: 1 if x > 0 else 0)

#Generating frequent itemsets using the Apriori algorithm (Market Basket Analysis)
frequent_itemsets = apriori(basket_sets, min_support=0.001, use_colnames=True)

# Generate association rules using the frequent itemsets and specify the metric and minimum threshold
rules = association_rules(frequent_itemsets, metric="lift", min_threshold=1)
rules = rules[['antecedents', 'consequents', 'antecedent support', 'consequent support', 'support', 'confidence', 'lift']]


# Creating an empty DataFrame with bool type
binary_subcategories = pd.DataFrame(index=basket.index, dtype=bool)

# Iterate over unique sub-categories
for sub_category in Minger_Cleaned_data['Sub-Category'].unique():
    # Create a binary column indicating presence of sub-category
    binary_subcategories[sub_category] = (basket[sub_category] > 0)

# Create a pivot table for the heatmap
heatmap_data = rules.pivot_table(index='antecedents', columns='consequents', values='lift')

# Plot heatmap
fig, ax = plt.subplots(figsize=(10, 8))
sns.heatmap(heatmap_data, annot=True, cmap='coolwarm', fmt=".2f", linewidths=0.5, ax=ax)
plt.title('Association Rules Heatmap (Lift)')
plt.xlabel('Consequents')
plt.ylabel('Antecedents')
st.pyplot(fig)

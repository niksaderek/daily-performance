import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -----------------------------
# Data loading and preprocessing
# -----------------------------
@st.cache_data
def load_and_standardize(filepath='daily.xlsx'):
    df = pd.read_excel(filepath, engine='openpyxl')

    col_mapping = {
        'Media\nBuyer': 'Media_Buyer',
        'Traffic \nSource': 'Traffic_Source',
        'Platform \nFee': 'Platform_Fee',
        'Net \nProfit': 'Net_Profit',
        'ROI I%': 'ROI_Pct',
        'Conversion\nRate': 'Conversion_Rate',
        '% of Total\nCalls': 'Pct_Total_Calls',
        'No-connect \n%': 'No_Connect_Pct',
        'Affiliate \nMargin': 'Affiliate_Margin',
        'No Connect': 'No_Connect',
        'Aff Pub': 'Aff_Pub'
    }

    df.rename(columns=col_mapping, inplace=True)
    return df

# -----------------------------
# Summary and trend functions
# -----------------------------
def get_date_summary(df):
    dates = sorted(df['Date'].unique())
    summary = []

    for date in dates:
        day_data = df[df['Date'] == date]
        summary.append({
            'Date': date,
            'Day': day_data['Day'].iloc[0],
            'Campaigns': len(day_data),
            'Spend': day_data['Spend'].sum(),
            'Revenue': day_data['Revenue'].sum(),
            'Net_Profit': day_data['Net_Profit'].sum(),
            'ROI': round((day_data['Revenue'].sum() / day_data['Spend'].sum() - 1) * 100, 2) if day_data['Spend'].sum() > 0 else 0,
            'Incoming': day_data['Incoming'].sum(),
            'Connected': day_data['Connected'].sum(),
            'Converted': day_data['Converted'].sum(),
            'Conv_Rate': round((day_data['Converted'].sum() / day_data['Connected'].sum() * 100), 2) if day_data['Connected'].sum() > 0 else 0,
            'No_Connect_Pct': round(day_data['No_Connect_Pct'].mean() * 100, 2)
        })

    return pd.DataFrame(summary)

# -----------------------------
# Dashboard visualizations
# -----------------------------
def generate_dashboard(df):
    latest_date = df['Date'].max()
    latest_data = df[df['Date'] == latest_date]

    st.header(f"Daily Performance Dashboard - {latest_date} ({latest_data['Day'].iloc[0]})")

    # -----------------------------
    # Top 10 Publishers by Profit
    # -----------------------------
    pub_perf = latest_data.groupby('Media_Buyer').agg({
        'Net_Profit': 'sum',
        'Revenue': 'sum',
        'Spend': 'sum',
        'Conversion_Rate': 'mean'
    }).round(2)
    pub_perf['ROI'] = ((pub_perf['Revenue'] / pub_perf['Spend'] - 1) * 100).round(2)
    pub_perf['Margin'] = ((pub_perf['Net_Profit'] / pub_perf['Revenue']) * 100).round(2)
    pub_perf = pub_perf.sort_values('Net_Profit', ascending=False).head(10)

    colors_list = ['#2ecc71' if x > 0 else '#e74c3c' for x in pub_perf['Net_Profit']]

    fig1 = go.Figure()
    fig1.add_trace(
        go.Bar(
            x=pub_perf.index,
            y=pub_perf['Net_Profit'],
            marker_color=colors_list,
            text=[f"${x:,.0f}" for x in pub_perf['Net_Profit']],
            textposition='auto',
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Profit: $%{y:,.0f}<br>"
                "ROI: %{customdata[0]:.2f}%<br>"
                "Conv Rate: %{customdata[1]:.2f}%<br>"
                "Margin: %{customdata[2]:.2f}%<extra></extra>"
            ),
            customdata=list(zip(pub_perf['ROI'], pub_perf['Conversion_Rate'], pub_perf['Margin']))
        )
    )
    fig1.update_layout(
        yaxis_title='Net Profit ($)',
        xaxis_tickangle=-45,
        margin=dict(t=80, b=120, l=50, r=50)
    )
    fig1.update_xaxes(automargin=True)
    fig1.update_yaxes(automargin=True)
    st.plotly_chart(fig1, use_container_width=True)

    # -----------------------------
    # Conversion Funnel
    # -----------------------------
    total_incoming = latest_data['Incoming'].sum()
    total_connected = latest_data['Connected'].sum()
    total_converted = latest_data['Converted'].sum()

    fig2 = go.Figure()
    fig2.add_trace(
        go.Funnel(
            y=['Incoming Calls', 'Connected', 'Converted'],
            x=[total_incoming, total_connected, total_converted],
            textposition='inside',
            textinfo='value+percent initial',
            marker=dict(color=['#3498db', '#2ecc71', '#f39c12']),
            hovertemplate='<b>%{y}</b><br>Count: %{x:,.0f}<br>%{percentInitial}<extra></extra>'
        )
    )
    st.plotly_chart(fig2, use_container_width=True)

    # -----------------------------
    # Affiliate vs Internal - Profit
    # -----------------------------
    non_test_data = latest_data[latest_data['Spend'] > 50]
    affiliate = non_test_data[non_test_data['Aff_Pub'] == True]
    internal = non_test_data[non_test_data['Aff_Pub'] == False]

    aff_profit = affiliate['Net_Profit'].sum() if len(affiliate) > 0 else 0
    int_profit = internal['Net_Profit'].sum() if len(internal) > 0 else 0

    fig3 = go.Figure()
    fig3.add_trace(
        go.Bar(
            x=['Profit'],
            y=[aff_profit],
            name='Affiliate',
            marker_color='#4883aa',
            text=[f"${aff_profit:,.0f}"],
            textposition='auto'
        )
    )
    fig3.add_trace(
        go.Bar(
            x=['Profit'],
            y=[int_profit],
            name='Internal',
            marker_color='#de5dd7',
            text=[f"${int_profit:,.0f}"],
            textposition='auto'
        )
    )
    fig3.update_layout(
        yaxis_title='Profit ($)',
        barmode='group',
        margin=dict(t=80, b=120, l=50, r=50)
    )
    fig3.update_xaxes(automargin=True)
    fig3.update_yaxes(automargin=True)
    st.plotly_chart(fig3, use_container_width=True)

    # -----------------------------
    # Affiliate vs Internal - Percent Metrics
    # -----------------------------
    pct_categories = ['ROI', 'Margin', 'Conv (connected)']
    aff_roi = ((affiliate['Revenue'].sum() / affiliate['Spend'].sum() - 1) * 100) if len(affiliate) > 0 else 0
    aff_margin = (affiliate['Net_Profit'].sum() / affiliate['Revenue'].sum() * 100) if len(affiliate) > 0 else 0
    aff_conv = (affiliate['Converted'].sum() / affiliate['Connected'].sum() * 100) if len(affiliate) > 0 else 0

    int_roi = ((internal['Revenue'].sum() / internal['Spend'].sum() - 1) * 100) if len(internal) > 0 else 0
    int_margin = (internal['Net_Profit'].sum() / internal['Revenue'].sum() * 100) if len(internal) > 0 else 0
    int_conv = (internal['Converted'].sum() / internal['Connected'].sum() * 100) if len(internal) > 0 else 0

    fig4 = go.Figure()
    fig4.add_trace(
        go.Bar(
            x=pct_categories,
            y=[aff_roi, aff_margin, aff_conv],
            name='Affiliate',
            marker_color='#4883aa',
            text=[f'{aff_roi:.1f}%', f'{aff_margin:.1f}%', f'{aff_conv:.1f}%'],
            textposition='auto'
        )
    )
    fig4.add_trace(
        go.Bar(
            x=pct_categories,
            y=[int_roi, int_margin, int_conv],
            name='Internal',
            marker_color='#de5dd7',
            text=[f'{int_roi:.1f}%', f'{int_margin:.1f}%', f'{int_conv:.1f}%'],
            textposition='auto'
        )
    )
    fig4.update_layout(
        yaxis_title='Percentage (%)',
        barmode='group',
        margin=dict(t=80, b=120, l=50, r=50)
    )
    fig4.update_xaxes(automargin=True)
    fig4.update_yaxes(automargin=True)
    st.plotly_chart(fig4, use_container_width=True)

# -----------------------------
# Streamlit app
# -----------------------------
df = load_and_standardize()
generate_dashboard(df)

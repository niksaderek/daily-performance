"""
Multi-Day Performance Analysis Tool
Analyzes all dates in the cumulative daily.xlsx file and tracks trends over time.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -------------------- Data Loading --------------------
def load_and_standardize(filepath='daily.xlsx'):
    """Load Excel file and standardize column names."""
    df = pd.read_excel(filepath)

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

# -------------------- Summary Functions --------------------
def get_date_summary(df):
    """Get summary of all dates in dataset."""
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

# -------------------- Trend & Analysis --------------------
def show_trend_summary(df):
    date_summary = get_date_summary(df)

    if len(date_summary) <= 1:
        st.info("Not enough data for historical trend summary.")
        return

    st.subheader("Historical Trend Summary")
    st.write(f"Data Range: {date_summary['Date'].min()} to {date_summary['Date'].max()} ({len(date_summary)} days)")
    st.dataframe(date_summary)

    if len(date_summary) >= 3:
        st.subheader("3-Day Trends")
        recent = date_summary.tail(3)
        metrics_to_track = ['Spend', 'Revenue', 'Net_Profit', 'ROI', 'Conv_Rate', 'No_Connect_Pct']

        trends = []
        for metric in metrics_to_track:
            values = recent[metric].values
            if len(values) == 3:
                trend = 'UP' if values[2] > values[1] > values[0] else 'DOWN' if values[2] < values[1] < values[0] else 'MIXED'
                avg = values.mean()
                trends.append(f"{metric}: {trend} (3-day avg: {avg:.2f})")
        st.write("\n".join(trends))

def analyze_latest_day(df):
    latest_date = df['Date'].max()
    latest_data = df[df['Date'] == latest_date]

    st.subheader(f"Daily Performance Report - {latest_date} ({latest_data['Day'].iloc[0]})")

    total_spend = latest_data["Spend"].sum()
    total_revenue = latest_data["Revenue"].sum()
    total_profit = latest_data["Net_Profit"].sum()
    overall_roi = (total_revenue / total_spend - 1) * 100 if total_spend > 0 else 0
    overall_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
    conv_rate = (latest_data["Converted"].sum() / latest_data["Connected"].sum() * 100) if latest_data["Connected"].sum() > 0 else 0

    st.markdown(f"""
**Overall Metrics**  
- {len(latest_data)} campaigns running  
- Spend: ${total_spend:,.0f} → Revenue: ${total_revenue:,.0f}  
- Net Profit: ${total_profit:,.2f} ({overall_roi:.2f}% ROI, {overall_margin:.2f}% Margin)  
- Conversion Rate: {conv_rate:.2f}% (of connected calls)
""")

def compare_with_previous(df):
    dates = sorted(df['Date'].unique())
    if len(dates) < 2:
        st.info("Not enough data for day-over-day comparison.")
        return

    latest_date = dates[-1]
    previous_date = dates[-2]

    latest = df[df['Date'] == latest_date]
    previous = df[df['Date'] == previous_date]

    metrics = {
        'Campaigns': (len(latest), len(previous)),
        'Spend': (latest['Spend'].sum(), previous['Spend'].sum()),
        'Revenue': (latest['Revenue'].sum(), previous['Revenue'].sum()),
        'Net Profit': (latest['Net_Profit'].sum(), previous['Net_Profit'].sum()),
        'ROI %': ((latest['Revenue'].sum() / latest['Spend'].sum() - 1) * 100,
                  (previous['Revenue'].sum() / previous['Spend'].sum() - 1) * 100),
        'Conv Rate %': ((latest['Converted'].sum() / latest['Connected'].sum() * 100),
                        (previous['Converted'].sum() / previous['Connected'].sum() * 100)),
        'No-Connect %': (latest['No_Connect_Pct'].mean() * 100, previous['No_Connect_Pct'].mean() * 100)
    }

    st.subheader(f"Day-over-Day Comparison: {previous_date} vs {latest_date}")
    for metric, (curr, prev) in metrics.items():
        change = curr - prev
        pct_change = (change / prev * 100) if prev != 0 else 0
        st.write(f"{metric}: {curr} ({change:+.2f}, {pct_change:+.1f}%)")

# -------------------- Dashboard --------------------
def generate_interactive_dashboard(df):
    latest_date = df['Date'].max()
    latest_data = df[df['Date'] == latest_date]

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            'Top 10 Publishers by Profit',
            'Conversion Funnel',
            'Affiliate vs Internal - Profit ($)',
            'Affiliate vs Internal - Performance Metrics (%)'
        ),
        specs=[[{"type": "bar"}, {"type": "funnel"}],
               [{"type": "bar"}, {"type": "bar"}]],
        vertical_spacing=0.15,
        horizontal_spacing=0.12,
        row_heights=[0.5, 0.5]
    )

    # Top 10 Publishers
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

    fig.add_trace(
        go.Bar(
            x=pub_perf.index,
            y=pub_perf['Net_Profit'],
            marker_color=colors_list,
            text=[f'${x:,.0f}' for x in pub_perf['Net_Profit']],
            textposition='outside',
            name='Top 10 Profit',
            hovertemplate='<b>%{x}</b><br>Profit: $%{y:,.0f}<br>ROI: %{customdata[0]:.1f}%<br>Margin: %{customdata[2]:.1f}%<br>Conv Rate: %{customdata[1]:.1%}<extra></extra>',
            customdata=list(zip(pub_perf['ROI'], pub_perf['Conversion_Rate'], pub_perf['Margin']))
        ),
        row=1, col=1
    )

    # Conversion Funnel
    total_incoming = latest_data['Incoming'].sum()
    total_connected = latest_data['Connected'].sum()
    total_converted = latest_data['Converted'].sum()
    funnel = go.Funnel(
        y=['Incoming', 'Connected', 'Converted'],
        x=[total_incoming, total_connected, total_converted],
        textinfo='value+percent initial'
    )
    fig.add_trace(funnel, row=1, col=2)

    # Affiliate vs Internal Profit
    affiliate = latest_data.groupby('Aff_Pub')['Net_Profit'].sum()
    fig.add_trace(go.Bar(
        x=affiliate.index,
        y=affiliate.values,
        text=[f'${x:,.0f}' for x in affiliate.values],
        textposition='outside',
        name='Affiliate/Internal Profit'
    ), row=2, col=1)

    # Affiliate vs Internal Performance %
    affiliate_metrics = latest_data.groupby('Aff_Pub').agg({
        'ROI_Pct': 'mean',
        'Conversion_Rate': 'mean',
        'Net_Profit': 'sum'
    }).round(2)
    fig.add_trace(
        go.Bar(
            x=affiliate_metrics.index,
            y=affiliate_metrics['ROI_Pct'],
            name='ROI %',
            text=[f'{x:.1f}%' for x in affiliate_metrics['ROI_Pct']],
            textposition='outside'
        ),
        row=2, col=2
    )
    fig.add_trace(
        go.Bar(
            x=affiliate_metrics.index,
            y=affiliate_metrics['Conversion_Rate'],
            name='Conv Rate %',
            text=[f'{x:.1f}%' for x in affiliate_metrics['Conversion_Rate']],
            textposition='outside'
        ),
        row=2, col=2
    )

    fig.update_layout(height=900, width=1200, title_text=f"Interactive Dashboard - {latest_date}", barmode='group')
    
    st.plotly_chart(fig, use_container_width=True)

# -------------------- Streamlit App --------------------
st.title("Daily Performance Analysis Dashboard")

df = load_and_standardize('daily.xlsx')

show_trend_summary(df)
analyze_latest_day(df)
compare_with_previous(df)
generate_interactive_dashboard(df)

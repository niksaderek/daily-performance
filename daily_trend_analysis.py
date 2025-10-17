"""
Multi-Day Performance Analysis Tool
Analyzes all dates in the cumulative daily.xlsx file and tracks trends over time.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import webbrowser
import os

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

def analyze_latest_day(df):
    """Detailed analysis of the most recent day."""
    latest_date = df['Date'].max()
    latest_data = df[df['Date'] == latest_date]

    print('=' * 100)
    print(f'DAILY PERFORMANCE REPORT - {latest_date} ({latest_data["Day"].iloc[0]})')
    print('=' * 100)
    print()

    # Overall metrics - simplified
    print('>> OVERALL PERFORMANCE')
    print()
    total_spend = latest_data["Spend"].sum()
    total_revenue = latest_data["Revenue"].sum()
    total_profit = latest_data["Net_Profit"].sum()
    overall_roi = (total_revenue / total_spend - 1) * 100 if total_spend > 0 else 0
    overall_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0

    print(f'  • {len(latest_data)} campaigns running')
    print(f'  • ${total_spend:,.0f} spend -> ${total_revenue:,.0f} revenue')
    print(f'  • ${total_profit:,.2f} net profit ({overall_roi:.2f}% ROI, {overall_margin:.2f}% margin)')
    print(f'  • {(latest_data["Converted"].sum() / latest_data["Connected"].sum() * 100):.2f}% conversion rate (of connected calls)', end='')

    conv_rate = (latest_data["Converted"].sum() / latest_data["Connected"].sum() * 100)
    if conv_rate >= 40:
        print(' (excellent!)')
    elif conv_rate >= 25:
        print(' (strong)')
    elif conv_rate >= 15:
        print(' (good)')
    else:
        print(' (needs improvement)')
    print()

    # Publisher performance with concentration analysis
    pub_perf = latest_data.groupby('Media_Buyer').agg({
        'Spend': 'sum',
        'Revenue': 'sum',
        'Net_Profit': 'sum',
        'Conversion_Rate': 'mean',
        'Incoming': 'sum',
        'Connected': 'sum',
        'Converted': 'sum'
    }).round(2)
    pub_perf['Campaigns'] = latest_data.groupby('Media_Buyer').size()
    pub_perf['ROI'] = ((pub_perf['Revenue'] / pub_perf['Spend'] - 1) * 100).round(2)
    pub_perf['Margin'] = ((pub_perf['Net_Profit'] / pub_perf['Revenue']) * 100).round(2)
    pub_perf = pub_perf.sort_values('Net_Profit', ascending=False)

    total_profit = latest_data['Net_Profit'].sum()
    total_calls = latest_data['Incoming'].sum()

    print('>> TOP PERFORMERS')
    print()
    top_5 = pub_perf.head(5)
    top_5_profit = top_5['Net_Profit'].sum()
    top_5_calls = top_5['Incoming'].sum()

    for i, (publisher, row) in enumerate(top_5.iterrows(), 1):
        profit_share = (row['Net_Profit'] / total_profit * 100)
        print(f'  {i}. {publisher[:25]}: ${row["Net_Profit"]:,.0f} ({profit_share:.1f}% of total)')

    print()
    concentration_pct = top_5_profit/total_profit*100
    print(f'  Top 5 combined: {concentration_pct:.1f}% of total profit')
    if concentration_pct > 60:
        print(f'  [!] High concentration - diversify publisher base')
    print()

    # Affiliate vs Internal comparison (exclude test accounts with <$50 spend)
    non_test_data = latest_data[latest_data['Spend'] > 50]

    if len(non_test_data) > 0:
        print('>> AFFILIATE vs INTERNAL')
        print()

        # Split by Aff_Pub flag
        affiliate = non_test_data[non_test_data['Aff_Pub'] == True]
        internal = non_test_data[non_test_data['Aff_Pub'] == False]

        if len(affiliate) > 0:
            aff_spend = affiliate['Spend'].sum()
            aff_revenue = affiliate['Revenue'].sum()
            aff_profit = affiliate['Net_Profit'].sum()
            aff_roi = ((aff_revenue / aff_spend) - 1) * 100 if aff_spend > 0 else 0
            aff_margin = (aff_profit / aff_revenue * 100) if aff_revenue > 0 else 0
            aff_conv = (affiliate['Converted'].sum() / affiliate['Connected'].sum() * 100) if affiliate['Connected'].sum() > 0 else 0

            print(f'  AFFILIATE PUBLISHERS ({len(affiliate)} campaigns):')
            print(f'    Spend: ${aff_spend:,.0f} | Revenue: ${aff_revenue:,.0f} | Profit: ${aff_profit:,.0f}')
            print(f'    ROI: {aff_roi:.2f}% | Margin: {aff_margin:.2f}% | Conv Rate (of connected): {aff_conv:.2f}%')
            print()

        if len(internal) > 0:
            int_spend = internal['Spend'].sum()
            int_revenue = internal['Revenue'].sum()
            int_profit = internal['Net_Profit'].sum()
            int_roi = ((int_revenue / int_spend) - 1) * 100 if int_spend > 0 else 0
            int_margin = (int_profit / int_revenue * 100) if int_revenue > 0 else 0
            int_conv = (internal['Converted'].sum() / internal['Connected'].sum() * 100) if internal['Connected'].sum() > 0 else 0

            print(f'  INTERNAL PUBLISHERS ({len(internal)} campaigns):')
            print(f'    Spend: ${int_spend:,.0f} | Revenue: ${int_revenue:,.0f} | Profit: ${int_profit:,.0f}')
            print(f'    ROI: {int_roi:.2f}% | Margin: {int_margin:.2f}% | Conv Rate (of connected): {int_conv:.2f}%')
            print()

        # Comparison
        if len(affiliate) > 0 and len(internal) > 0:
            profit_diff = aff_profit - int_profit
            roi_diff = aff_roi - int_roi
            conv_diff = aff_conv - int_conv

            print(f'  COMPARISON:')
            if profit_diff > 0:
                print(f'    [+] Affiliate leading by ${profit_diff:,.0f} profit')
            else:
                print(f'    [+] Internal leading by ${abs(profit_diff):,.0f} profit')

            if roi_diff > 0:
                print(f'    [+] Affiliate ROI +{roi_diff:.2f} percentage points higher')
            else:
                print(f'    [+] Internal ROI +{abs(roi_diff):.2f} percentage points higher')

            if conv_diff > 0:
                print(f'    [+] Affiliate conversion +{conv_diff:.2f} percentage points higher')
            else:
                print(f'    [+] Internal conversion +{abs(conv_diff):.2f} percentage points higher')

        print()

    # Quality alerts - simplified (only campaigns with >$500 spend)
    print('>> QUALITY ALERTS')
    print()

    significant_campaigns = latest_data[latest_data['Spend'] > 500]
    neg_roi_all = significant_campaigns[significant_campaigns['ROI_Pct'] < 0]
    low_conv = significant_campaigns[significant_campaigns['Conversion_Rate'] < 0.10]
    high_no_connect = significant_campaigns[significant_campaigns['No_Connect_Pct'] > 0.50]

    if len(neg_roi_all) > 0:
        print(f'  [!] {len(neg_roi_all)} campaigns with negative ROI (>$500 spend):')
        for _, row in neg_roi_all.iterrows():
            print(f'      {row["Media_Buyer"]} - {row["Vertical"]}: ${row["Spend"]:,.0f} spend, {row["ROI_Pct"]:.1f}% ROI')
        print()

    if len(low_conv) > 0:
        print(f'  [!] {len(low_conv)} campaigns with low conversion <10% (>$500 spend):')
        for _, row in low_conv.iterrows():
            print(f'      {row["Media_Buyer"]} - {row["Vertical"]}: {row["Conversion_Rate"]:.1%} conv rate')
        print()

    if len(high_no_connect) > 0:
        print(f'  [!] {len(high_no_connect)} campaigns with high no-connect >50% (>$500 spend):')
        for _, row in high_no_connect.iterrows():
            print(f'      {row["Media_Buyer"]} - {row["Vertical"]}: {row["No_Connect_Pct"]:.1%} no-connect rate')
        print()

    if len(neg_roi_all) == 0 and len(low_conv) == 0 and len(high_no_connect) == 0:
        print('  [OK] No major quality issues detected (campaigns >$500 spend)')
        print()

    print()
    print('=' * 100)

def compare_with_previous(df):
    """Compare latest day with previous day."""
    dates = sorted(df['Date'].unique())

    if len(dates) < 2:
        print('### DAY-OVER-DAY COMPARISON ###')
        print('Not enough data yet - need at least 2 days for comparison')
        print()
        return

    latest_date = dates[-1]
    previous_date = dates[-2]

    latest = df[df['Date'] == latest_date]
    previous = df[df['Date'] == previous_date]

    print()
    print('=' * 100)
    print(f'DAY-OVER-DAY COMPARISON: {previous_date} vs {latest_date}')
    print('=' * 100)
    print()

    # Overall changes
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

    print('### OVERALL METRICS ###')
    for metric, (curr, prev) in metrics.items():
        change = curr - prev
        pct_change = (change / prev * 100) if prev != 0 else 0

        if metric in ['Spend', 'Revenue', 'Net Profit']:
            print(f'{metric}: ${curr:,.0f} (${change:+,.0f}, {pct_change:+.1f}%)')
        elif 'Rate' in metric or '%' in metric:
            print(f'{metric}: {curr:.2f}% ({change:+.2f} points)')
        else:
            print(f'{metric}: {curr:.0f} ({change:+.0f})')

    print()

    # Publisher changes
    print('### PUBLISHER PERFORMANCE CHANGES ###')

    latest_pubs = latest.groupby('Media_Buyer')['Net_Profit'].sum()
    previous_pubs = previous.groupby('Media_Buyer')['Net_Profit'].sum()

    all_pubs = set(latest_pubs.index) | set(previous_pubs.index)

    changes = []
    for pub in all_pubs:
        curr_profit = latest_pubs.get(pub, 0)
        prev_profit = previous_pubs.get(pub, 0)
        change = curr_profit - prev_profit

        if pub in latest_pubs.index and pub not in previous_pubs.index:
            status = 'NEW'
        elif pub not in latest_pubs.index and pub in previous_pubs.index:
            status = 'INACTIVE'
        else:
            status = 'ACTIVE'

        changes.append({
            'Publisher': pub,
            'Current': curr_profit,
            'Previous': prev_profit,
            'Change': change,
            'Status': status
        })

    changes_df = pd.DataFrame(changes).sort_values('Change', ascending=False)

    print('\nTop 5 Profit Gainers:')
    for _, row in changes_df.head(5).iterrows():
        if row['Status'] == 'NEW':
            print(f'  {row["Publisher"]}: ${row["Current"]:,.2f} (NEW PUBLISHER)')
        else:
            print(f'  {row["Publisher"]}: ${row["Current"]:,.2f} ({row["Change"]:+,.2f})')

    print('\nTop 5 Profit Decliners:')
    for _, row in changes_df.tail(5).iterrows():
        if row['Status'] == 'INACTIVE':
            print(f'  {row["Publisher"]}: ${row["Previous"]:,.2f} -> INACTIVE TODAY')
        else:
            print(f'  {row["Publisher"]}: ${row["Current"]:,.2f} ({row["Change"]:+,.2f})')

    print()

def show_trend_summary(df):
    """Show performance trends across all dates."""
    date_summary = get_date_summary(df)

    if len(date_summary) == 1:
        return

    print('=' * 100)
    print('HISTORICAL TREND SUMMARY')
    print('=' * 100)
    print()

    print(f'Data Range: {date_summary["Date"].min()} to {date_summary["Date"].max()} ({len(date_summary)} days)')
    print()

    print('### DAILY METRICS ###')
    print(date_summary.to_string(index=False))
    print()

    # Calculate trends
    if len(date_summary) >= 3:
        print('### 3-DAY TRENDS ###')
        recent = date_summary.tail(3)

        metrics_to_track = ['Spend', 'Revenue', 'Net_Profit', 'ROI', 'Conv_Rate', 'No_Connect_Pct']

        for metric in metrics_to_track:
            values = recent[metric].values
            if len(values) == 3:
                trend = 'UP' if values[2] > values[1] > values[0] else 'DOWN' if values[2] < values[1] < values[0] else 'MIXED'
                avg = values.mean()
                print(f'{metric}: {trend} (3-day avg: {avg:.2f})')

        print()

def identify_new_campaigns(df):
    """Find campaigns that appeared in latest day but not in previous days."""
    dates = sorted(df['Date'].unique())

    if len(dates) < 2:
        return

    latest_date = dates[-1]
    previous_dates = dates[:-1]

    latest = df[df['Date'] == latest_date]
    historical = df[df['Date'].isin(previous_dates)]

    # Identify new buyer-vertical-source combinations
    latest_combos = set(zip(latest['Media_Buyer'], latest['Vertical'], latest['Traffic_Source']))
    historical_combos = set(zip(historical['Media_Buyer'], historical['Vertical'], historical['Traffic_Source']))

    new_combos = latest_combos - historical_combos

    if len(new_combos) == 0:
        return

    print('=' * 100)
    print(f'NEW CAMPAIGN LAUNCHES ({latest_date})')
    print('=' * 100)
    print()

    new_campaign_data = latest[latest.apply(lambda x: (x['Media_Buyer'], x['Vertical'], x['Traffic_Source']) in new_combos, axis=1)]

    print(f'{len(new_campaign_data)} new campaign combinations detected:')
    print()

    for idx, row in new_campaign_data.iterrows():
        # Calculate ROI correctly
        campaign_roi = ((row['Revenue'] / row['Spend']) - 1) * 100 if row['Spend'] > 0 else 0

        print(f'{row["Media_Buyer"]} - {row["Vertical"]} ({row["Traffic_Source"]})')
        print(f'  Spend: ${row["Spend"]:,.0f} | Revenue: ${row["Revenue"]:,.0f} | Profit: ${row["Net_Profit"]:,.2f}')
        print(f'  ROI: {campaign_roi:.2f}% | Conv Rate: {row["Conversion_Rate"]:.2%}')

        # Early signal
        if row['Conversion_Rate'] > 0.15 and campaign_roi > 20:
            print('  [+] Strong Day 0 performance')
        elif campaign_roi < 0:
            print('  [-] Negative ROI on Day 0')

        print()

def generate_pdf_report(df, filename=None):
    """Generate a professional PDF report of the analysis."""
    if filename is None:
        latest_date = df['Date'].max()
        filename = f'Daily_Performance_Report_{latest_date}.pdf'

    # Create PDF document
    doc = SimpleDocTemplate(filename, pagesize=letter,
                           rightMargin=0.5*inch, leftMargin=0.5*inch,
                           topMargin=0.75*inch, bottomMargin=0.5*inch)

    # Container for PDF elements
    elements = []

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1a1a1a'),
        spaceAfter=30,
        alignment=TA_CENTER
    )
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#2c3e50'),
        spaceAfter=12,
        spaceBefore=12,
        borderWidth=0,
        borderColor=colors.HexColor('#3498db'),
        borderPadding=5,
        backColor=colors.HexColor('#ecf0f1')
    )

    # Get data
    latest_date = df['Date'].max()
    latest_data = df[df['Date'] == latest_date]

    # Title
    title = Paragraph(f"<b>Daily Performance Report</b><br/>{latest_date} ({latest_data['Day'].iloc[0]})", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.2*inch))

    # === OVERALL PERFORMANCE ===
    elements.append(Paragraph("<b>Overall Performance</b>", heading_style))
    elements.append(Spacer(1, 0.1*inch))

    overall_data = [
        ['Metric', 'Value'],
        ['Campaigns', f"{len(latest_data)}"],
        ['Spend', f"${latest_data['Spend'].sum():,.0f}"],
        ['Revenue', f"${latest_data['Revenue'].sum():,.0f}"],
        ['Net Profit', f"${latest_data['Net_Profit'].sum():,.2f}"],
        ['ROI', f"{(latest_data['Revenue'].sum() / latest_data['Spend'].sum() - 1) * 100:.2f}%"],
        ['Conversion Rate', f"{(latest_data['Converted'].sum() / latest_data['Connected'].sum() * 100):.2f}%"],
        ['Avg No-Connect', f"{latest_data['No_Connect_Pct'].mean():.1%}"]
    ]

    overall_table = Table(overall_data, colWidths=[3*inch, 2*inch])
    overall_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(overall_table)
    elements.append(Spacer(1, 0.3*inch))

    # === TOP 5 PUBLISHERS ===
    elements.append(Paragraph("<b>Top 5 Publishers</b>", heading_style))
    elements.append(Spacer(1, 0.1*inch))

    pub_perf = latest_data.groupby('Media_Buyer').agg({
        'Spend': 'sum',
        'Revenue': 'sum',
        'Net_Profit': 'sum',
        'Conversion_Rate': 'mean',
        'Incoming': 'sum'
    }).round(2)
    pub_perf['Campaigns'] = latest_data.groupby('Media_Buyer').size()
    pub_perf['ROI'] = ((pub_perf['Revenue'] / pub_perf['Spend'] - 1) * 100).round(2)
    pub_perf = pub_perf.sort_values('Net_Profit', ascending=False)

    top_5 = pub_perf.head(5)
    top_5_data = [['#', 'Publisher', 'Net Profit', 'ROI', 'Conv Rate', 'Calls', 'Camps']]

    for i, (publisher, row) in enumerate(top_5.iterrows(), 1):
        top_5_data.append([
            str(i),
            publisher[:25],
            f"${row['Net_Profit']:,.0f}",
            f"{row['ROI']:.1f}%",
            f"{row['Conversion_Rate']:.1%}",
            f"{row['Incoming']:.0f}",
            f"{row['Campaigns']:.0f}"
        ])

    top_5_table = Table(top_5_data, colWidths=[0.3*inch, 2.2*inch, 1*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.6*inch])
    top_5_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#27ae60')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(top_5_table)

    total_profit = latest_data['Net_Profit'].sum()
    total_calls = latest_data['Incoming'].sum()
    top_5_profit = top_5['Net_Profit'].sum()
    top_5_calls = top_5['Incoming'].sum()

    if top_5_profit >= 0:
        concentration_text = f"Top 5 Impact: ${top_5_profit:,.0f} profit ({top_5_profit/total_profit*100:.1f}% of total profit) from {top_5_calls/total_calls*100:.1f}% of calls"
    else:
        concentration_text = f"Top 5 Impact: ${abs(top_5_profit):,.0f} loss ({abs(top_5_profit)/total_profit*100:.1f}% drag on total profit) from {top_5_calls/total_calls*100:.1f}% of calls"

    elements.append(Spacer(1, 0.1*inch))
    elements.append(Paragraph(concentration_text, styles['Normal']))
    elements.append(Spacer(1, 0.3*inch))

    # === BOTTOM 10 PUBLISHERS ===
    elements.append(Paragraph("<b>Bottom 10 Publishers (Excluding Tests)</b>", heading_style))
    elements.append(Spacer(1, 0.1*inch))

    bottom_performers = pub_perf[pub_perf['Spend'] > 50].tail(10).sort_values('Net_Profit')

    if len(bottom_performers) > 0:
        bottom_data = [['#', 'Publisher', 'Net Profit', 'ROI', 'Conv Rate', 'Action']]

        for i, (publisher, row) in enumerate(bottom_performers.iterrows(), 1):
            if row['ROI'] < -20:
                action = 'PAUSE NOW'
            elif row['Conversion_Rate'] < 0.08:
                action = 'FIX QUALITY'
            elif row['ROI'] < 10 and row['ROI'] > 0:
                action = 'OPTIMIZE'
            else:
                action = 'REVIEW'

            bottom_data.append([
                str(i),
                publisher[:25],
                f"${row['Net_Profit']:,.0f}",
                f"{row['ROI']:.1f}%",
                f"{row['Conversion_Rate']:.1%}",
                action
            ])

        bottom_table = Table(bottom_data, colWidths=[0.3*inch, 2.5*inch, 1*inch, 0.8*inch, 0.8*inch, 1.2*inch])
        bottom_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e74c3c')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightpink),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(bottom_table)

        bottom_10_profit = bottom_performers['Net_Profit'].sum()
        if bottom_10_profit < 0:
            impact_text = f"Bottom 10 Impact: ${abs(bottom_10_profit):,.0f} loss ({abs(bottom_10_profit)/total_profit*100:.1f}% drag on total profit)"
        else:
            impact_text = f"Bottom 10 Impact: ${bottom_10_profit:,.0f} profit ({bottom_10_profit/total_profit*100:.1f}% of total profit)"
        elements.append(Spacer(1, 0.1*inch))
        elements.append(Paragraph(impact_text, styles['Normal']))

    elements.append(PageBreak())

    # === QUALITY ISSUES ===
    elements.append(Paragraph("<b>Quality Issues Summary</b>", heading_style))
    elements.append(Spacer(1, 0.1*inch))

    neg_roi_all = latest_data[latest_data['ROI_Pct'] < 0]
    low_conv = latest_data[latest_data['Conversion_Rate'] < 0.10]
    high_no_connect = latest_data[latest_data['No_Connect_Pct'] > 0.50]

    quality_data = [
        ['Issue', 'Count', '% of Total'],
        ['Negative ROI campaigns', len(neg_roi_all), f"{len(neg_roi_all)/len(latest_data)*100:.0f}%"],
        ['Low conversion (<10%)', len(low_conv), f"{len(low_conv)/len(latest_data)*100:.0f}%"],
        ['High no-connect (>50%)', len(high_no_connect), f"{len(high_no_connect)/len(latest_data)*100:.0f}%"]
    ]

    quality_table = Table(quality_data, colWidths=[3*inch, 1*inch, 1.2*inch])
    quality_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f39c12')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lemonchiffon),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(quality_table)
    elements.append(Spacer(1, 0.3*inch))

    # === KEY INSIGHTS ===
    elements.append(Paragraph("<b>Key Insights</b>", heading_style))
    elements.append(Spacer(1, 0.1*inch))

    insights_text = []
    insights_text.append("<b>What's Working:</b>")

    # Top vertical
    vertical_perf = latest_data.groupby('Vertical').agg({
        'Spend': 'sum',
        'Revenue': 'sum',
        'Net_Profit': 'sum'
    })
    vertical_perf['Campaigns'] = latest_data.groupby('Vertical').size()
    vertical_perf['ROI'] = ((vertical_perf['Revenue'] / vertical_perf['Spend'] - 1) * 100).round(2)
    top_vertical = vertical_perf.sort_values('Net_Profit', ascending=False).iloc[0]
    top_vertical_name = vertical_perf.sort_values('Net_Profit', ascending=False).index[0]
    insights_text.append(f"• Vertical: {top_vertical_name} ({top_vertical['Campaigns']:.0f} campaigns, ${top_vertical['Net_Profit']:,.0f} profit, {top_vertical['ROI']:.1f}% ROI)")

    # Top traffic source
    source_perf = latest_data.groupby('Traffic_Source').agg({
        'Spend': 'sum',
        'Revenue': 'sum',
        'Net_Profit': 'sum'
    })
    source_perf['Campaigns'] = latest_data.groupby('Traffic_Source').size()
    source_perf['ROI'] = ((source_perf['Revenue'] / source_perf['Spend'] - 1) * 100).round(2)
    top_source = source_perf.sort_values('Net_Profit', ascending=False).iloc[0]
    top_source_name = source_perf.sort_values('Net_Profit', ascending=False).index[0]
    insights_text.append(f"• Traffic Source: {top_source_name} ({top_source['Campaigns']:.0f} campaigns, ${top_source['Net_Profit']:,.0f} profit, {top_source['ROI']:.2f}% ROI)")

    insights_text.append("")
    insights_text.append("<b>What's Failing:</b>")

    # Facebook issues
    fb_data = latest_data[latest_data['Traffic_Source'] == 'FACEBOOK']
    if len(fb_data) > 0:
        fb_profit = fb_data['Net_Profit'].sum()
        fb_roi = (fb_data['Revenue'].sum() / fb_data['Spend'].sum() - 1) * 100 if fb_data['Spend'].sum() > 0 else 0
        if fb_profit < 0:
            insights_text.append(f"• Facebook traffic: ${fb_profit:,.0f} loss, {fb_roi:.1f}% ROI")

    # No-connect issues
    avg_no_connect = latest_data['No_Connect_Pct'].mean()
    if avg_no_connect > 0.3:
        insights_text.append(f"• High no-connect rates: {avg_no_connect:.0%} average")

    for line in insights_text:
        elements.append(Paragraph(line, styles['Normal']))

    elements.append(Spacer(1, 0.3*inch))

    # === RECOMMENDATIONS ===
    elements.append(Paragraph("<b>Immediate Actions</b>", heading_style))
    elements.append(Spacer(1, 0.1*inch))

    recommendations = []

    # Facebook
    fb_data_nonzero = fb_data[fb_data['Spend'] > 0]
    if len(fb_data_nonzero) > 0 and fb_data_nonzero['Net_Profit'].sum() < -100:
        recommendations.append("• PAUSE/MONITOR Facebook traffic - Currently unprofitable during testing phase")

    # Scale winners
    top_3_pubs = pub_perf[pub_perf['ROI'] > 30].head(3)
    if len(top_3_pubs) > 0:
        pub_names = ', '.join(top_3_pubs.index.tolist())
        recommendations.append(f"• Scale proven winners: {pub_names}")

    # No-connect optimization
    worst_no_connect = latest_data[latest_data['No_Connect_Pct'] > 0.70].groupby('Media_Buyer')['No_Connect_Pct'].mean().sort_values(ascending=False)
    if len(worst_no_connect) > 0:
        recommendations.append("• Optimize no-connect rates - Work with publishers showing >50% no-connect")

    for rec in recommendations:
        elements.append(Paragraph(rec, styles['Normal']))

    # Build PDF
    doc.build(elements)
    return filename

def generate_interactive_dashboard(df, filename='daily_performance_dashboard.html'):
    """Generate an interactive HTML dashboard with visualizations."""
    latest_date = df['Date'].max()
    latest_data = df[df['Date'] == latest_date]

    # Create subplots - 2 rows, 2 columns
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            'Top 10 Publishers by Profit',
            'Conversion Funnel',
            'Affiliate vs Internal - Profit ($)',
            'Affiliate vs Internal - Performance Metrics (%)'
        ),
        specs=[
            [{"type": "bar"}, {"type": "funnel"}],
            [{"type": "bar"}, {"type": "bar"}]
        ],
        vertical_spacing=0.15,
        horizontal_spacing=0.12,
        row_heights=[0.5, 0.5]
    )

    # === Chart 1: Top 10 Publishers ===
    pub_perf = latest_data.groupby('Media_Buyer').agg({
        'Net_Profit': 'sum',
        'Revenue': 'sum',
        'Spend': 'sum',
        'Conversion_Rate': 'mean'
    }).round(2)
    # Calculate ROI from aggregated Revenue/Spend (not from averaging ROI_Pct)
    pub_perf['ROI'] = ((pub_perf['Revenue'] / pub_perf['Spend'] - 1) * 100).round(2)
    pub_perf['Margin'] = ((pub_perf['Net_Profit'] / pub_perf['Revenue']) * 100).round(2)
    pub_perf = pub_perf.sort_values('Net_Profit', ascending=False)

    top_10 = pub_perf.head(10)

    # Color bars: green for positive profit, red for negative
    colors_list = ['#2ecc71' if x > 0 else '#e74c3c' for x in top_10['Net_Profit']]

    fig.add_trace(
        go.Bar(
            x=top_10.index,
            y=top_10['Net_Profit'],
            marker_color=colors_list,
            text=[f'${x:,.0f}' for x in top_10['Net_Profit']],
            textposition='outside',
            name='Top 10 Profit',
            hovertemplate='<b>%{x}</b><br>Profit: $%{y:,.0f}<br>ROI: %{customdata[0]:.1f}%<br>Margin: %{customdata[2]:.1f}%<br>Conv Rate: %{customdata[1]:.1%}<extra></extra>',
            customdata=list(zip(top_10['ROI'], top_10['Conversion_Rate'], top_10['Margin']))
        ),
        row=1, col=1
    )

    # === Chart 2: Conversion Funnel ===
    total_incoming = latest_data['Incoming'].sum()
    total_connected = latest_data['Connected'].sum()
    total_converted = latest_data['Converted'].sum()

    # Calculate rates
    connect_rate = (total_connected / total_incoming * 100) if total_incoming > 0 else 0
    conv_rate = (total_converted / total_connected * 100) if total_connected > 0 else 0

    fig.add_trace(
        go.Funnel(
            name='Funnel',
            y=['Incoming Calls', 'Connected', 'Converted'],
            x=[total_incoming, total_connected, total_converted],
            textposition='inside',
            textinfo='value+percent initial',
            marker=dict(color=['#3498db', '#2ecc71', '#f39c12']),
            hovertemplate='<b>%{y}</b><br>Count: %{x:,.0f}<br>%{percentInitial}<extra></extra>'
        ),
        row=1, col=2
    )

    # === Chart 3 & 4: Affiliate vs Internal ===
    non_test_data = latest_data[latest_data['Spend'] > 50]

    affiliate = non_test_data[non_test_data['Aff_Pub'] == True]
    internal = non_test_data[non_test_data['Aff_Pub'] == False]

    # Calculate metrics
    if len(affiliate) > 0:
        aff_spend = affiliate['Spend'].sum()
        aff_revenue = affiliate['Revenue'].sum()
        aff_profit = affiliate['Net_Profit'].sum()
        aff_roi = ((aff_revenue / aff_spend) - 1) * 100 if aff_spend > 0 else 0
        aff_margin = (aff_profit / aff_revenue * 100) if aff_revenue > 0 else 0
        aff_conv = (affiliate['Converted'].sum() / affiliate['Connected'].sum() * 100) if affiliate['Connected'].sum() > 0 else 0
    else:
        aff_profit = aff_roi = aff_margin = aff_conv = 0

    if len(internal) > 0:
        int_spend = internal['Spend'].sum()
        int_revenue = internal['Revenue'].sum()
        int_profit = internal['Net_Profit'].sum()
        int_roi = ((int_revenue / int_spend) - 1) * 100 if int_spend > 0 else 0
        int_margin = (int_profit / int_revenue * 100) if int_revenue > 0 else 0
        int_conv = (internal['Converted'].sum() / internal['Connected'].sum() * 100) if internal['Connected'].sum() > 0 else 0
    else:
        int_profit = int_roi = int_margin = int_conv = 0

    # Chart 3: Profit comparison (left, row 2)
    if len(affiliate) > 0:
        fig.add_trace(
            go.Bar(
                name='Affiliate',
                x=['Profit'],
                y=[aff_profit],
                marker_color='#4883aa',
                text=[f'${aff_profit:,.0f}'],
                textposition='outside',
                hovertemplate='<b>Affiliate</b><br>Profit: $%{y:,.0f}<extra></extra>',
                showlegend=True
            ),
            row=2, col=1
        )

    if len(internal) > 0:
        fig.add_trace(
            go.Bar(
                name='Internal',
                x=['Profit'],
                y=[int_profit],
                marker_color='#de5dd7',
                text=[f'${int_profit:,.0f}'],
                textposition='outside',
                hovertemplate='<b>Internal</b><br>Profit: $%{y:,.0f}<extra></extra>',
                showlegend=True
            ),
            row=2, col=1
        )

    # Chart 4: Percentage metrics comparison (right, row 2)
    pct_categories = ['ROI', 'Margin', 'Conv (connected)']

    if len(affiliate) > 0:
        fig.add_trace(
            go.Bar(
                name='Affiliate',
                x=pct_categories,
                y=[aff_roi, aff_margin, aff_conv],
                marker_color="#4883aa",
                text=[f'{aff_roi:.1f}%', f'{aff_margin:.1f}%', f'{aff_conv:.1f}%'],
                textposition='outside',
                hovertemplate='<b>Affiliate</b><br>%{x}: %{y:.2f}%<extra></extra>',
                showlegend=False
            ),
            row=2, col=2
        )

    if len(internal) > 0:
        fig.add_trace(
            go.Bar(
                name='Internal',
                x=pct_categories,
                y=[int_roi, int_margin, int_conv],
                marker_color="#de5dd7",
                text=[f'{int_roi:.1f}%', f'{int_margin:.1f}%', f'{int_conv:.1f}%'],
                textposition='outside',
                hovertemplate='<b>Internal</b><br>%{x}: %{y:.2f}%<extra></extra>',
                showlegend=False
            ),
            row=2, col=2
        )

    # Update layout
    fig.update_xaxes(tickangle=-45, row=1, col=1)
    fig.update_xaxes(title_text='', row=2, col=1)
    fig.update_xaxes(title_text='Metric', row=2, col=2)

    fig.update_yaxes(title_text='Net Profit ($)', row=1, col=1)
    fig.update_yaxes(title_text='Profit ($)', row=2, col=1)
    fig.update_yaxes(title_text='Percentage (%)', row=2, col=2)

    fig.update_layout(
        title_text=f"Daily Performance Dashboard - {latest_date} ({latest_data['Day'].iloc[0]})",
        title_font_size=24,
        showlegend=True,
        height=1000,
        template='plotly_white',
        barmode='group'
    )

    # Save and open
    fig.write_html(filename)
    print(f'\nInteractive dashboard saved: {filename}')

    # Auto-open in browser
    webbrowser.open('file://' + os.path.abspath(filename))

    return filename

if __name__ == '__main__':
    print('Loading data from daily.xlsx...')
    print()

    df = load_and_standardize()

    # Show trend summary first (if multiple days)
    show_trend_summary(df)

    # Latest day detailed analysis
    analyze_latest_day(df)

    # Day-over-day comparison
    compare_with_previous(df)

    # Generate PDF report silently
    generate_pdf_report(df)

    # Generate interactive dashboard
    print('\nGenerating interactive visualizations...')
    generate_interactive_dashboard(df)

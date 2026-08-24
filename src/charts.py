# ABOUTME: Builds the Plotly figures for the dashboard, keeping colors, hover text,
# ABOUTME: and layout conventions consistent across every panel.

import plotly.graph_objects as go
from plotly.subplots import make_subplots

PROFIT_GREEN = "#2ecc71"
LOSS_RED = "#e74c3c"
FUNNEL_COLORS = ["#3498db", "#2ecc71", "#f39c12"]
AFFILIATE_BLUE = "#4883aa"
INTERNAL_PURPLE = "#de5dd7"

BASE_LAYOUT = dict(
    margin=dict(t=60, b=80, l=60, r=40),
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    hoverlabel=dict(font_size=13),
)


def _finalize(fig, **layout):
    fig.update_layout(**{**BASE_LAYOUT, **layout})
    fig.update_xaxes(automargin=True)
    fig.update_yaxes(automargin=True)
    return fig


def publisher_profit_chart(board):
    """Bar chart of the most profitable media buyers, colored by profit or loss."""
    colors = [PROFIT_GREEN if value > 0 else LOSS_RED for value in board["Net_Profit"]]

    fig = go.Figure(
        go.Bar(
            x=board.index,
            y=board["Net_Profit"],
            marker_color=colors,
            text=[f"${value:,.0f}" for value in board["Net_Profit"]],
            textposition="auto",
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Profit: $%{y:,.0f}<br>"
                "ROI: %{customdata[0]:.1f}%<br>"
                "Conv Rate: %{customdata[1]:.1%}<br>"
                "Margin: %{customdata[2]:.1f}%<extra></extra>"
            ),
            customdata=list(
                zip(board["ROI"], board["Conversion_Rate"], board["Margin"])
            ),
        )
    )
    return _finalize(fig, yaxis_title="Net Profit ($)", xaxis_tickangle=-45)


def profit_trend_chart(summary):
    """Profit bars against an ROI line, so volume and efficiency read together."""
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=summary["Label"],
            y=summary["Net_Profit"],
            name="Net Profit",
            marker_color=[PROFIT_GREEN if v > 0 else LOSS_RED for v in summary["Net_Profit"]],
            hovertemplate="<b>%{x}</b><br>Profit: $%{y:,.0f}<extra></extra>",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=summary["Label"],
            y=summary["ROI"],
            name="ROI %",
            mode="lines+markers",
            line=dict(color="#34495e", width=3),
            hovertemplate="<b>%{x}</b><br>ROI: %{y:.1f}%<extra></extra>",
        ),
        secondary_y=True,
    )

    fig.update_yaxes(title_text="Net Profit ($)", secondary_y=False)
    fig.update_yaxes(title_text="ROI (%)", secondary_y=True, showgrid=False)
    # Without an explicit day scale, a short window renders as an hourly axis.
    fig.update_xaxes(type="category")
    return _finalize(fig, legend=dict(orientation="h", y=1.12, x=0), bargap=0.55)


def conversion_funnel_chart(stages, values):
    """Funnel from incoming calls through to conversions."""
    fig = go.Figure(
        go.Funnel(
            y=stages,
            x=values,
            textposition="inside",
            textinfo="value+percent initial",
            marker=dict(color=FUNNEL_COLORS),
            hovertemplate="<b>%{y}</b><br>Count: %{x:,.0f}<br>%{percentInitial}<extra></extra>",
        )
    )
    return _finalize(fig)


def segment_profit_chart(segments):
    """Absolute profit contributed by affiliate versus internal traffic."""
    fig = go.Figure()
    for name, color in (("Affiliate", AFFILIATE_BLUE), ("Internal", INTERNAL_PURPLE)):
        profit = segments[name]["net_profit"]
        fig.add_trace(
            go.Bar(
                x=["Net Profit"],
                y=[profit],
                name=name,
                marker_color=color,
                text=[f"${profit:,.0f}"],
                textposition="auto",
            )
        )
    return _finalize(fig, yaxis_title="Profit ($)", barmode="group")


def segment_rate_chart(segments):
    """Rate-based comparison, where the two segments are directly comparable."""
    categories = ["ROI", "Margin", "Conversion"]
    fig = go.Figure()

    for name, color in (("Affiliate", AFFILIATE_BLUE), ("Internal", INTERNAL_PURPLE)):
        values = [
            segments[name]["roi_pct"],
            segments[name]["margin_pct"],
            segments[name]["conversion_pct"],
        ]
        fig.add_trace(
            go.Bar(
                x=categories,
                y=values,
                name=name,
                marker_color=color,
                text=[f"{value:.1f}%" for value in values],
                textposition="auto",
            )
        )
    return _finalize(fig, yaxis_title="Percentage (%)", barmode="group")

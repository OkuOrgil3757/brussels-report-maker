import os
import io
import json
import warnings
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, render_template
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

warnings.filterwarnings("ignore")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB for multi-file

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

PALETTE = ["#6366f1", "#8b5cf6", "#06b6d4", "#10b981", "#f59e0b",
           "#ef4444", "#ec4899", "#3b82f6", "#84cc16", "#f97316",
           "#14b8a6", "#a855f7", "#0ea5e9", "#22c55e", "#fb923c"]

CHART_BG   = "#ffffff"
GRID_COLOR = "#f1f5f9"
FONT_COLOR = "#1e293b"
FONT_FAMILY = "Inter, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif"

BASE_LAYOUT = dict(
    paper_bgcolor=CHART_BG,
    plot_bgcolor=CHART_BG,
    font=dict(family=FONT_FAMILY, color=FONT_COLOR, size=12),
    title_font=dict(size=15, color=FONT_COLOR, family=FONT_FAMILY),
    hoverlabel=dict(
        bgcolor="white",
        bordercolor="#e2e8f0",
        font=dict(size=12, family=FONT_FAMILY, color=FONT_COLOR),
    ),
    xaxis=dict(gridcolor=GRID_COLOR, linecolor="#e2e8f0", zerolinecolor="#e2e8f0"),
    yaxis=dict(gridcolor=GRID_COLOR, linecolor="#e2e8f0", zerolinecolor="#e2e8f0"),
    margin=dict(t=64, b=48, l=48, r=32),
)


# ─── helpers ──────────────────────────────────────────────────────────────────

def coerce_numerics(df):
    for col in df.columns:
        if df[col].dtype == object:
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().mean() > 0.6:
                df[col] = converted
    return df


def detect_columns(df):
    roles = {}
    for col in df.columns:
        s = df[col].dropna()
        if len(s) == 0:
            roles[col] = "empty"
            continue
        if s.dtype == object:
            try:
                parsed = pd.to_datetime(s.astype(str), errors="coerce")
                if parsed.notna().mean() > 0.7:
                    roles[col] = "date"
                    continue
            except Exception:
                pass
        if pd.api.types.is_datetime64_any_dtype(s):
            roles[col] = "date"
            continue
        if pd.api.types.is_numeric_dtype(s):
            roles[col] = "numeric"
            continue
        nuniq = s.nunique()
        if nuniq <= 50:
            roles[col] = "category"
        else:
            roles[col] = "text"
    return roles


def to_fig_json(fig):
    return json.loads(fig.to_json())


def fmt_val(v):
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"{v/1_000:.1f}K"
    return f"{v:.0f}"


def layout(**kwargs):
    d = dict(**BASE_LAYOUT)
    for k, v in kwargs.items():
        if isinstance(v, dict) and k in d and isinstance(d[k], dict):
            d[k] = {**d[k], **v}
        else:
            d[k] = v
    return d


def load_all_files(sheet_name):
    """Load and concatenate all saved upload files for the given sheet."""
    frames = []
    i = 0
    while True:
        path = os.path.join(UPLOAD_FOLDER, f"upload_{i}.xlsx")
        if not os.path.exists(path):
            break
        try:
            xl = pd.ExcelFile(path)
            # try exact sheet name first, then index 0
            if sheet_name in xl.sheet_names:
                df = pd.read_excel(path, sheet_name=sheet_name)
            elif isinstance(sheet_name, int) and sheet_name < len(xl.sheet_names):
                df = pd.read_excel(path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(path, sheet_name=0)
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
            df = coerce_numerics(df)
            frames.append(df)
        except Exception:
            pass
        i += 1
    if not frames:
        raise ValueError("No valid files found")
    if len(frames) == 1:
        return frames[0]
    # align columns: union of all columns
    return pd.concat(frames, ignore_index=True, sort=False)


def clear_uploads():
    i = 0
    while True:
        path = os.path.join(UPLOAD_FOLDER, f"upload_{i}.xlsx")
        if not os.path.exists(path):
            break
        os.remove(path)
        i += 1


# ─── chart generators ─────────────────────────────────────────────────────────

def chart_top_items(df, item_col, num_col, top=20):
    g = df.groupby(item_col)[num_col].sum().nlargest(top).sort_values()
    n = len(g)
    colors = [PALETTE[i % len(PALETTE)] for i in range(n)]
    fig = go.Figure(go.Bar(
        x=g.values, y=g.index.astype(str), orientation="h",
        marker=dict(color=colors, line=dict(width=0)),
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
        textfont=dict(size=11, color=FONT_COLOR),
        hovertemplate="<b>%{y}</b><br>%{x:,.0f}<extra></extra>",
    ))
    fig.update_layout(**layout(
        title=f"Top {n} Items — {num_col}",
        xaxis_title=num_col, yaxis_title="",
        height=max(420, n * 32 + 80),
        xaxis=dict(gridcolor=GRID_COLOR),
        margin=dict(l=220, r=90, t=64, b=48),
    ))
    return {"id": "top_items", "title": f"Top Items by {num_col}", "fig": to_fig_json(fig)}


def chart_category_bar(df, cat_col, num_col):
    g = df.groupby(cat_col)[num_col].sum().sort_values(ascending=False)
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(g))]
    fig = go.Figure(go.Bar(
        x=g.index.astype(str), y=g.values,
        marker=dict(color=colors, line=dict(width=0)),
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>%{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(**layout(
        title=f"{num_col} by {cat_col}",
        yaxis_title=num_col, xaxis_title="",
        height=440,
        yaxis=dict(gridcolor=GRID_COLOR),
        margin=dict(t=64, b=90, l=60, r=32),
    ))
    return {"id": f"cat_bar_{cat_col}", "title": f"{num_col} by {cat_col}", "fig": to_fig_json(fig)}


def chart_pie(df, cat_col, num_col):
    g = df.groupby(cat_col)[num_col].sum().sort_values(ascending=False).head(12)
    fig = go.Figure(go.Pie(
        labels=g.index.astype(str), values=g.values,
        hole=0.5,
        marker=dict(colors=PALETTE[:len(g)], line=dict(color="white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=11),
        hovertemplate="<b>%{label}</b><br>%{value:,.0f} (%{percent})<extra></extra>",
        pull=[0.03] + [0] * (len(g) - 1),
    ))
    fig.update_layout(**layout(
        title=f"Share of {num_col} by {cat_col}",
        height=460,
        showlegend=True,
        legend=dict(orientation="v", x=1.02, y=0.5),
        xaxis=dict(visible=False), yaxis=dict(visible=False),
    ))
    return {"id": f"pie_{cat_col}", "title": f"{cat_col} Share", "fig": to_fig_json(fig)}


def chart_dow(df, date_col, num_col):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    d["dow"] = d[date_col].dt.day_name()
    order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    g = d.groupby("dow")[num_col].sum().reindex(order).fillna(0)
    best_idx = int(g.values.argmax())
    colors = [PALETTE[0] if i != best_idx else "#ef4444" for i in range(7)]
    fig = go.Figure(go.Bar(
        x=g.index, y=g.values,
        marker=dict(color=colors, line=dict(width=0)),
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>%{y:,.0f}<extra></extra>",
    ))
    fig.add_annotation(
        text=f"★ Best: {order[best_idx]}",
        x=best_idx, y=g.values[best_idx],
        yshift=28, showarrow=False,
        font=dict(size=11, color="#ef4444", family=FONT_FAMILY),
    )
    fig.update_layout(**layout(
        title=f"{num_col} by Day of Week",
        yaxis_title=num_col, xaxis_title="",
        height=440,
        yaxis=dict(gridcolor=GRID_COLOR),
    ))
    return {"id": "dow", "title": "Revenue by Day of Week", "fig": to_fig_json(fig)}


def chart_trend(df, date_col, num_col):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    g = d.groupby(d[date_col].dt.date)[num_col].sum()
    roll = pd.Series(g.values, index=g.index).rolling(7, min_periods=1).mean()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=list(g.index), y=list(g.values),
        mode="lines+markers",
        name=num_col,
        line=dict(color=PALETTE[0], width=2, shape="spline", smoothing=0.8),
        marker=dict(size=5, color=PALETTE[0]),
        fill="tozeroy",
        fillcolor="rgba(99,102,241,0.08)",
        hovertemplate="<b>%{x}</b><br>%{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=list(roll.index), y=list(roll.values),
        mode="lines",
        name="7-day avg",
        line=dict(color="#ef4444", width=2, dash="dot", shape="spline", smoothing=0.8),
        hovertemplate="7-day avg: %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(**layout(
        title=f"{num_col} — Daily Trend",
        yaxis_title=num_col, xaxis_title="",
        height=420,
        hovermode="x unified",
        legend=dict(orientation="h", y=1.08, x=1, xanchor="right"),
        yaxis=dict(gridcolor=GRID_COLOR),
    ))
    return {"id": "trend", "title": "Daily Trend", "fig": to_fig_json(fig)}


def chart_diverging(df, item_col, num_col):
    g = df.groupby(item_col)[num_col].sum()
    g = g[g != 0].sort_values()
    if len(g) == 0:
        return None
    colors = ["#ef4444" if v < 0 else "#10b981" for v in g.values]
    fig = go.Figure(go.Bar(
        x=g.values, y=g.index.astype(str), orientation="h",
        marker=dict(color=colors, line=dict(width=0)),
        text=[f"{v:+,.0f}" for v in g.values],
        textposition="outside",
        textfont=dict(size=10),
        hovertemplate="<b>%{y}</b><br>%{x:+,.0f}<extra></extra>",
    ))
    fig.add_vline(x=0, line_width=1.5, line_color="#94a3b8")
    fig.update_layout(**layout(
        title=f"Surplus / Deficit — {item_col}",
        xaxis_title=num_col, yaxis_title="",
        height=max(420, len(g) * 28 + 80),
        margin=dict(l=220, r=90, t=64, b=48),
    ))
    return {"id": "diverging", "title": "Surplus / Deficit", "fig": to_fig_json(fig)}


def chart_employee_compare(df, emp_col, num_col):
    g = df.groupby(emp_col)[num_col].sum().sort_values(ascending=False)
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(g))]
    fig = go.Figure(go.Bar(
        x=g.index.astype(str), y=g.values,
        marker=dict(color=colors, line=dict(width=0)),
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>%{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(**layout(
        title=f"{num_col} by {emp_col}",
        yaxis_title=num_col, xaxis_title="",
        height=440,
        yaxis=dict(gridcolor=GRID_COLOR),
        margin=dict(t=64, b=80, l=60, r=32),
    ))
    return {"id": "employee", "title": f"By {emp_col}", "fig": to_fig_json(fig)}


def chart_heatmap(df, date_col, cat_col, num_col):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    d["week"] = d[date_col].dt.isocalendar().week.astype(str)
    pivot = d.groupby(["week", cat_col])[num_col].sum().unstack(fill_value=0)
    fig = go.Figure(go.Heatmap(
        z=pivot.values,
        x=[str(c) for c in pivot.columns],
        y=pivot.index.tolist(),
        colorscale=[[0, "#f8fafc"], [0.5, "#6366f1"], [1, "#312e81"]],
        hoverongaps=False,
        hovertemplate="Week %{y} / %{x}<br>%{z:,.0f}<extra></extra>",
        showscale=True,
    ))
    fig.update_layout(**layout(
        title=f"Weekly Heatmap — {num_col} by {cat_col}",
        xaxis_title=cat_col, yaxis_title="Week",
        height=max(380, len(pivot) * 30 + 100),
        xaxis=dict(side="bottom"),
        margin=dict(t=64, b=90, l=60, r=60),
    ))
    return {"id": "heatmap", "title": "Weekly Heatmap", "fig": to_fig_json(fig)}


def chart_treemap(df, cat_col, num_col, sub_col=None):
    if sub_col:
        g = df.groupby([cat_col, sub_col])[num_col].sum().reset_index()
        g = g[g[num_col] > 0]
        g[cat_col] = g[cat_col].astype(str)
        g[sub_col] = g[sub_col].astype(str)
        fig = px.treemap(
            g, path=[cat_col, sub_col], values=num_col,
            color=num_col,
            color_continuous_scale=[[0, "#e0e7ff"], [0.5, "#6366f1"], [1, "#312e81"]],
            title=f"Treemap — {num_col} by {cat_col} › {sub_col}",
        )
    else:
        g = df.groupby(cat_col)[num_col].sum().reset_index()
        g = g[g[num_col] > 0]
        g[cat_col] = g[cat_col].astype(str)
        fig = px.treemap(
            g, path=[cat_col], values=num_col,
            color=num_col,
            color_continuous_scale=[[0, "#e0e7ff"], [0.5, "#6366f1"], [1, "#312e81"]],
            title=f"Treemap — {num_col} by {cat_col}",
        )
    fig.update_traces(
        hovertemplate="<b>%{label}</b><br>%{value:,.0f}<extra></extra>",
        textfont=dict(family=FONT_FAMILY),
    )
    fig.update_layout(
        paper_bgcolor=CHART_BG, height=500,
        font=dict(family=FONT_FAMILY, color=FONT_COLOR),
        title_font=dict(size=15),
        margin=dict(t=64, b=20, l=20, r=20),
    )
    return {"id": "treemap", "title": "Treemap", "fig": to_fig_json(fig)}


def chart_cumulative(df, date_col, num_col):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    g = d.groupby(d[date_col].dt.date)[num_col].sum().cumsum()
    fig = go.Figure(go.Scatter(
        x=list(g.index), y=list(g.values),
        mode="lines",
        line=dict(color=PALETTE[1], width=3, shape="spline", smoothing=0.6),
        fill="tozeroy",
        fillcolor="rgba(139,92,246,0.1)",
        name=f"Cumulative {num_col}",
        hovertemplate="<b>%{x}</b><br>Total: %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(**layout(
        title=f"Cumulative {num_col}",
        yaxis_title=num_col, xaxis_title="",
        height=400,
        hovermode="x unified",
        yaxis=dict(gridcolor=GRID_COLOR),
    ))
    return {"id": "cumulative", "title": "Cumulative Total", "fig": to_fig_json(fig)}


def chart_scatter(df, num_col1, num_col2, cat_col=None):
    clean = df[[num_col1, num_col2] + ([cat_col] if cat_col else [])].dropna()
    if cat_col and clean[cat_col].nunique() <= 15:
        fig = px.scatter(
            clean, x=num_col1, y=num_col2, color=cat_col,
            color_discrete_sequence=PALETTE,
            title=f"{num_col1} vs {num_col2}",
            opacity=0.75,
        )
    else:
        fig = go.Figure(go.Scatter(
            x=clean[num_col1], y=clean[num_col2],
            mode="markers",
            marker=dict(color=PALETTE[2], size=7, opacity=0.65,
                        line=dict(color="white", width=0.5)),
            hovertemplate=f"{num_col1}: %{{x:,.0f}}<br>{num_col2}: %{{y:,.0f}}<extra></extra>",
        ))
        fig.update_layout(title=f"{num_col1} vs {num_col2}",
                          xaxis_title=num_col1, yaxis_title=num_col2)
    fig.update_layout(**layout(height=440, yaxis=dict(gridcolor=GRID_COLOR)))
    return {"id": "scatter", "title": f"Scatter: {num_col1} vs {num_col2}", "fig": to_fig_json(fig)}


def chart_distribution(df, num_col):
    vals = df[num_col].dropna()
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=["Histogram", "Box Plot"],
                        horizontal_spacing=0.12)
    fig.add_trace(go.Histogram(
        x=vals, nbinsx=30,
        marker=dict(color=PALETTE[0], line=dict(color="white", width=0.5)),
        hovertemplate="Range: %{x}<br>Count: %{y}<extra></extra>",
        name="count",
    ), row=1, col=1)
    fig.add_trace(go.Box(
        y=vals,
        marker=dict(color=PALETTE[1]),
        line=dict(color=PALETTE[1]),
        fillcolor="rgba(139,92,246,0.15)",
        hovertemplate="%{y:,.0f}<extra></extra>",
        name=num_col,
        boxpoints="outliers",
    ), row=1, col=2)
    fig.update_layout(
        title=f"Distribution — {num_col}",
        showlegend=False, height=400,
        paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG,
        font=dict(family=FONT_FAMILY, color=FONT_COLOR),
        title_font=dict(size=15),
        margin=dict(t=64, b=48, l=60, r=32),
    )
    fig.update_xaxes(gridcolor=GRID_COLOR, linecolor="#e2e8f0")
    fig.update_yaxes(gridcolor=GRID_COLOR, linecolor="#e2e8f0")
    return {"id": f"dist_{num_col}", "title": f"Distribution: {num_col}", "fig": to_fig_json(fig)}


def chart_monthly(df, date_col, num_col, cat_col=None):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    d["month"] = d[date_col].dt.to_period("M").astype(str)
    if cat_col and d[cat_col].nunique() <= 10:
        pivot = d.groupby(["month", cat_col])[num_col].sum().unstack(fill_value=0)
        fig = go.Figure()
        for i, col in enumerate(pivot.columns):
            fig.add_trace(go.Bar(
                name=str(col), x=pivot.index.tolist(), y=pivot[col].tolist(),
                marker=dict(color=PALETTE[i % len(PALETTE)], line=dict(width=0)),
                hovertemplate=f"<b>%{{x}}</b><br>{col}: %{{y:,.0f}}<extra></extra>",
            ))
        fig.update_layout(barmode="group", title=f"Monthly {num_col} by {cat_col}")
    else:
        g = d.groupby("month")[num_col].sum()
        fig = go.Figure(go.Bar(
            x=g.index.tolist(), y=g.values.tolist(),
            marker=dict(color=PALETTE[0], line=dict(width=0)),
            text=[fmt_val(v) for v in g.values],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>%{y:,.0f}<extra></extra>",
        ))
        fig.update_layout(title=f"Monthly {num_col}")
    fig.update_layout(**layout(
        yaxis_title=num_col, xaxis_title="",
        height=440,
        yaxis=dict(gridcolor=GRID_COLOR),
        margin=dict(t=64, b=80, l=60, r=32),
    ))
    return {"id": "monthly", "title": "Monthly Comparison", "fig": to_fig_json(fig)}


def chart_summary_table(df, roles):
    num_cols = [c for c, r in roles.items() if r == "numeric"]
    if not num_cols:
        return None
    stats = df[num_cols].describe().T.round(2)
    even_rows = ["#f8fafc" if i % 2 == 0 else "white" for i in range(len(stats))]
    fig = go.Figure(go.Table(
        header=dict(
            values=["<b>Column</b>"] + [f"<b>{c}</b>" for c in stats.columns],
            fill_color="#6366f1",
            font=dict(color="white", size=12, family=FONT_FAMILY),
            align="left", height=36,
            line=dict(color="#6366f1"),
        ),
        cells=dict(
            values=[stats.index.tolist()] + [stats[c].tolist() for c in stats.columns],
            fill_color=[even_rows],
            align="left",
            font=dict(size=12, family=FONT_FAMILY, color=FONT_COLOR),
            height=32,
            line=dict(color="#e2e8f0"),
        ),
    ))
    fig.update_layout(
        title="Summary Statistics",
        height=max(280, len(num_cols) * 34 + 110),
        paper_bgcolor=CHART_BG,
        font=dict(family=FONT_FAMILY),
        title_font=dict(size=15),
        margin=dict(t=64, b=20, l=20, r=20),
    )
    return {"id": "summary_table", "title": "Summary Statistics", "fig": to_fig_json(fig)}


def chart_category_mix(df, emp_col, cat_col, num_col):
    pivot = df.groupby([emp_col, cat_col])[num_col].sum().unstack(fill_value=0)
    pct = pivot.div(pivot.sum(axis=1), axis=0) * 100
    fig = go.Figure()
    for i, col in enumerate(pct.columns):
        fig.add_trace(go.Bar(
            name=str(col),
            x=pct.index.astype(str).tolist(),
            y=pct[col].tolist(),
            marker=dict(color=PALETTE[i % len(PALETTE)], line=dict(width=0)),
            hovertemplate="<b>%{x}</b> — " + str(col) + ": %{y:.1f}%<extra></extra>",
        ))
    fig.update_layout(**layout(
        barmode="stack",
        title=f"Category Mix by {emp_col} (%)",
        yaxis_title="%", xaxis_title="",
        height=460,
        yaxis=dict(gridcolor=GRID_COLOR, range=[0, 100]),
        legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center"),
        margin=dict(t=64, b=100, l=60, r=32),
    ))
    return {"id": "cat_mix", "title": f"Category Mix by {emp_col}", "fig": to_fig_json(fig)}


# ─── main chart pipeline ──────────────────────────────────────────────────────

def generate_charts(df):
    roles = detect_columns(df)
    charts = []

    date_cols = [c for c, r in roles.items() if r == "date"]
    num_cols  = [c for c, r in roles.items() if r == "numeric"]
    cat_cols  = [c for c, r in roles.items() if r == "category"]
    text_cols = [c for c, r in roles.items() if r == "text"]

    for dc in date_cols:
        df[dc] = pd.to_datetime(df[dc], errors="coerce")

    def pick(hints, pool):
        for h in hints:
            for c in pool:
                if h in c.lower():
                    return c
        return pool[0] if pool else None

    main_num  = pick(["дүн","amount","total","revenue","value","price","орлого"], num_cols)
    if main_num is None and num_cols:
        main_num = max(num_cols, key=lambda c: df[c].sum(skipna=True))

    qty_col = None
    for c in num_cols:
        if c != main_num and any(h in c.lower() for h in ["qty","тоо","count","quantity"]):
            qty_col = c
            break
    if qty_col is None and len(num_cols) > 1:
        qty_col = next((c for c in num_cols if c != main_num), None)

    emp_col  = pick(["ажилтан","staff","employee","waiter","person","user"], cat_cols + text_cols)
    item_col = None
    for c in text_cols + cat_cols:
        if c != emp_col and any(h in c.lower() for h in ["бараа","item","product","goods","нэр","name"]):
            item_col = c
            break
    if item_col is None:
        cands = [c for c in text_cols if c != emp_col]
        item_col = max(cands, key=lambda c: df[c].nunique()) if cands else None

    main_date = date_cols[0] if date_cols else None
    main_cat  = cat_cols[0] if cat_cols else None

    # 1. Summary table
    t = chart_summary_table(df, roles)
    if t: charts.append(t)

    # 2. Daily trend
    if main_date and main_num:
        c = chart_trend(df, main_date, main_num)
        if c: charts.append(c)

    # 3. Cumulative
    if main_date and main_num:
        c = chart_cumulative(df, main_date, main_num)
        if c: charts.append(c)

    # 4. Day of week
    if main_date and main_num:
        c = chart_dow(df, main_date, main_num)
        if c: charts.append(c)

    # 5. Monthly
    if main_date and main_num:
        c = chart_monthly(df, main_date, main_num, main_cat)
        if c: charts.append(c)

    # 6. Category bar + pie
    for cat in cat_cols[:3]:
        if main_num:
            charts.append(chart_category_bar(df, cat, main_num))
            charts.append(chart_pie(df, cat, main_num))

    # 7. Top items
    if item_col and main_num:
        charts.append(chart_top_items(df, item_col, main_num))
    if item_col and qty_col and qty_col != main_num:
        charts.append(chart_top_items(df, item_col, qty_col))

    # 8. Employee
    if emp_col and main_num:
        if main_cat and main_cat != emp_col:
            charts.append(chart_category_mix(df, emp_col, main_cat, main_num))
        charts.append(chart_employee_compare(df, emp_col, main_num))

    # 9. Diverging
    if item_col and main_num:
        grp = df.groupby(item_col)[main_num].sum()
        if (grp > 0).any() and (grp < 0).any():
            charts.append(chart_diverging(df, item_col, main_num))

    # 10. Heatmap
    if main_date and main_cat and main_num:
        c = chart_heatmap(df, main_date, main_cat, main_num)
        if c: charts.append(c)

    # 11. Treemap
    if main_cat and main_num:
        sub = item_col if item_col and item_col != main_cat and df[item_col].nunique() <= 200 else None
        charts.append(chart_treemap(df, main_cat, main_num, sub))

    # 12. Scatter
    if len(num_cols) >= 2:
        charts.append(chart_scatter(df, num_cols[0], num_cols[1], main_cat))

    # 13. Distributions
    for nc in num_cols[:2]:
        charts.append(chart_distribution(df, nc))

    return charts


# ─── routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    """Accept up to 10 files, save them, return sheet names from first file."""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    clear_uploads()
    saved = []
    for i, f in enumerate(files[:10]):
        if not f.filename.lower().endswith((".xlsx", ".xls", ".xlsm")):
            continue
        buf = f.read()
        path = os.path.join(UPLOAD_FOLDER, f"upload_{i}.xlsx")
        with open(path, "wb") as fp:
            fp.write(buf)
        saved.append(f.filename)

    if not saved:
        return jsonify({"error": "No valid Excel files (.xlsx/.xls) found"}), 400

    # Get sheet names from first saved file
    first = os.path.join(UPLOAD_FOLDER, "upload_0.xlsx")
    xl = pd.ExcelFile(first)
    return jsonify({"sheets": xl.sheet_names, "files": saved})


@app.route("/api/preview", methods=["POST"])
def preview():
    data = request.get_json()
    sheet = data.get("sheet", 0)
    try:
        df = load_all_files(sheet)
        roles = detect_columns(df)
        sample = df.head(5).fillna("").astype(str).to_dict("records")
        return jsonify({
            "columns": list(df.columns),
            "roles": roles,
            "shape": [int(df.shape[0]), int(df.shape[1])],
            "sample": sample,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/charts", methods=["POST"])
def charts():
    data = request.get_json()
    sheet = data.get("sheet", 0)
    try:
        df = load_all_files(sheet)
        result = generate_charts(df)
        return jsonify({"charts": result})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    print("=" * 60)
    print("  Brussels Report Maker")
    print("  Open http://localhost:8080 in your browser")
    print("=" * 60)
    app.run(debug=False, port=8080, host="0.0.0.0")

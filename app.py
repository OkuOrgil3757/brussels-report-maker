import os
import io
import json
import warnings
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_from_directory
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

warnings.filterwarnings("ignore")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

PALETTE = ["#1a3c5e", "#2e6da4", "#4a9fd4", "#7dc4e8", "#b3e0f5",
           "#f5cba7", "#e74c3c", "#2ecc71", "#9b59b6", "#e67e22",
           "#1abc9c", "#e91e63", "#ff9800", "#607d8b", "#795548"]


# ─── helpers ──────────────────────────────────────────────────────────────────

def coerce_numerics(df):
    """Try to convert object columns that look numeric into float."""
    for col in df.columns:
        if df[col].dtype == object:
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().mean() > 0.6:
                df[col] = converted
    return df


def detect_columns(df):
    """Return dict with column role hints."""
    roles = {}
    for col in df.columns:
        s = df[col].dropna()
        if len(s) == 0:
            roles[col] = "empty"
            continue
        # try date
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
        if nuniq <= 2 and set(s.unique()) <= {True, False, 0, 1, "yes", "no", "true", "false"}:
            roles[col] = "boolean"
        elif nuniq <= 50:
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
        return f"{v/1_000:.0f}K"
    return f"{v:.0f}"


# ─── chart generators ─────────────────────────────────────────────────────────

def chart_top_items(df, item_col, num_col, top=20):
    g = df.groupby(item_col)[num_col].sum().nlargest(top).sort_values()
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(g))]
    fig = go.Figure(go.Bar(
        x=g.values, y=g.index, orientation="h",
        marker_color=colors,
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
    ))
    fig.update_layout(
        title=f"Top {top} — {item_col} by {num_col}",
        xaxis_title=num_col, yaxis_title="",
        height=max(400, len(g) * 30 + 100),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=200, r=80, t=60, b=40),
    )
    return {"id": "top_items", "title": f"Top Items by {num_col}", "fig": to_fig_json(fig)}


def chart_category_bar(df, cat_col, num_col):
    g = df.groupby(cat_col)[num_col].sum().sort_values(ascending=False)
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(g))]
    fig = go.Figure(go.Bar(
        x=g.index, y=g.values,
        marker_color=colors,
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
    ))
    fig.update_layout(
        title=f"{num_col} by {cat_col}",
        yaxis_title=num_col, xaxis_title="",
        plot_bgcolor="white", paper_bgcolor="white",
        height=420, margin=dict(t=60, b=80),
    )
    return {"id": f"cat_bar_{cat_col}", "title": f"{num_col} by {cat_col}", "fig": to_fig_json(fig)}


def chart_pie(df, cat_col, num_col):
    g = df.groupby(cat_col)[num_col].sum().sort_values(ascending=False)
    fig = go.Figure(go.Pie(
        labels=g.index, values=g.values,
        hole=0.45,
        marker_colors=PALETTE[:len(g)],
        textinfo="label+percent",
        hovertemplate="%{label}: %{value:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        title=f"Share of {num_col} by {cat_col}",
        height=460,
        paper_bgcolor="white",
    )
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
    best = g.idxmax()
    colors = ["#e74c3c" if x == best else PALETTE[1] for x in g.index]
    fig = go.Figure(go.Bar(
        x=g.index, y=g.values,
        marker_color=colors,
        text=[fmt_val(v) for v in g.values],
        textposition="outside",
    ))
    fig.update_layout(
        title=f"{num_col} by Day of Week  (best: {best})",
        yaxis_title=num_col, xaxis_title="",
        plot_bgcolor="white", paper_bgcolor="white",
        height=420, margin=dict(t=60, b=40),
    )
    return {"id": "dow", "title": "Revenue by Day of Week", "fig": to_fig_json(fig)}


def chart_trend(df, date_col, num_col):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    g = d.groupby(d[date_col].dt.date)[num_col].sum()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=g.index, y=g.values,
        mode="lines+markers",
        line=dict(color=PALETTE[0], width=2.5),
        marker=dict(size=6),
        fill="tozeroy",
        fillcolor="rgba(74,159,212,0.15)",
        name=num_col,
    ))
    # 7-day rolling average
    roll = pd.Series(g.values, index=g.index).rolling(7, min_periods=1).mean()
    fig.add_trace(go.Scatter(
        x=roll.index, y=roll.values,
        mode="lines",
        line=dict(color="#e74c3c", width=1.5, dash="dot"),
        name="7-day avg",
    ))
    fig.update_layout(
        title=f"{num_col} — Daily Trend",
        yaxis_title=num_col, xaxis_title="Date",
        plot_bgcolor="white", paper_bgcolor="white",
        height=400, margin=dict(t=60, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return {"id": "trend", "title": "Daily Trend", "fig": to_fig_json(fig)}


def chart_diverging(df, item_col, num_col):
    g = df.groupby(item_col)[num_col].sum()
    g = g[g != 0].sort_values()
    if len(g) == 0:
        return None
    colors = ["#cc3333" if v < 0 else "#2e8b57" for v in g.values]
    fig = go.Figure(go.Bar(
        x=g.values, y=g.index, orientation="h",
        marker_color=colors,
        text=[f"{v:+.0f}" for v in g.values],
        textposition="outside",
    ))
    fig.add_vline(x=0, line_width=1, line_color="#555")
    fig.update_layout(
        title=f"Surplus / Deficit — {item_col}",
        xaxis_title=num_col, yaxis_title="",
        height=max(400, len(g) * 28 + 100),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=200, r=80, t=60, b=40),
    )
    return {"id": "diverging", "title": "Surplus / Deficit", "fig": to_fig_json(fig)}


def chart_employee_compare(df, emp_col, num_col, cat_col=None):
    if cat_col:
        pivot = df.groupby([emp_col, cat_col])[num_col].sum().unstack(fill_value=0)
        fig = go.Figure()
        for i, cat in enumerate(pivot.columns):
            fig.add_trace(go.Bar(
                name=str(cat), x=pivot.index, y=pivot[cat],
                marker_color=PALETTE[i % len(PALETTE)],
            ))
        fig.update_layout(barmode="stack",
                          title=f"{num_col} by {emp_col} (stacked by {cat_col})")
    else:
        g = df.groupby(emp_col)[num_col].sum().sort_values(ascending=False)
        colors = [PALETTE[i % len(PALETTE)] for i in range(len(g))]
        fig = go.Figure(go.Bar(
            x=g.index, y=g.values,
            marker_color=colors,
            text=[fmt_val(v) for v in g.values],
            textposition="outside",
        ))
        fig.update_layout(title=f"{num_col} by {emp_col}")

    fig.update_layout(
        yaxis_title=num_col, xaxis_title="",
        plot_bgcolor="white", paper_bgcolor="white",
        height=440, margin=dict(t=60, b=80),
    )
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
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale="Blues",
        hoverongaps=False,
        hovertemplate="Week %{y} / %{x}: %{z:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        title=f"Heatmap: {num_col} by Week × {cat_col}",
        xaxis_title=cat_col, yaxis_title="Week",
        height=max(380, len(pivot) * 28 + 100),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=60, b=80, l=60, r=40),
    )
    return {"id": "heatmap", "title": f"Weekly Heatmap", "fig": to_fig_json(fig)}


def chart_treemap(df, cat_col, num_col, sub_col=None):
    if sub_col:
        g = df.groupby([cat_col, sub_col])[num_col].sum().reset_index()
        g = g[g[num_col] > 0]
        fig = px.treemap(
            g, path=[cat_col, sub_col], values=num_col,
            color=num_col, color_continuous_scale="Blues",
            title=f"Treemap — {num_col} by {cat_col} / {sub_col}",
        )
    else:
        g = df.groupby(cat_col)[num_col].sum().reset_index()
        g = g[g[num_col] > 0]
        fig = px.treemap(
            g, path=[cat_col], values=num_col,
            color=num_col, color_continuous_scale="Blues",
            title=f"Treemap — {num_col} by {cat_col}",
        )
    fig.update_layout(height=480, paper_bgcolor="white", margin=dict(t=60, b=20, l=20, r=20))
    return {"id": "treemap", "title": "Treemap", "fig": to_fig_json(fig)}


def chart_cumulative(df, date_col, num_col):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    g = d.groupby(d[date_col].dt.date)[num_col].sum().cumsum()
    fig = go.Figure(go.Scatter(
        x=g.index, y=g.values,
        mode="lines",
        line=dict(color=PALETTE[0], width=3),
        fill="tozeroy",
        fillcolor="rgba(26,60,94,0.12)",
        name=f"Cumulative {num_col}",
    ))
    fig.update_layout(
        title=f"Cumulative {num_col}",
        yaxis_title=num_col, xaxis_title="Date",
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, margin=dict(t=60, b=40),
    )
    return {"id": "cumulative", "title": f"Cumulative Total", "fig": to_fig_json(fig)}


def chart_scatter(df, num_col1, num_col2, cat_col=None):
    if cat_col and df[cat_col].nunique() <= 20:
        fig = px.scatter(
            df, x=num_col1, y=num_col2, color=cat_col,
            color_discrete_sequence=PALETTE,
            title=f"{num_col1} vs {num_col2} (colored by {cat_col})",
        )
    else:
        fig = go.Figure(go.Scatter(
            x=df[num_col1], y=df[num_col2],
            mode="markers",
            marker=dict(color=PALETTE[1], size=6, opacity=0.6),
        ))
        fig.update_layout(title=f"{num_col1} vs {num_col2}", xaxis_title=num_col1, yaxis_title=num_col2)
    fig.update_layout(height=420, plot_bgcolor="white", paper_bgcolor="white", margin=dict(t=60))
    return {"id": "scatter", "title": f"Scatter: {num_col1} vs {num_col2}", "fig": to_fig_json(fig)}


def chart_numeric_distribution(df, num_col):
    vals = df[num_col].dropna()
    fig = make_subplots(rows=1, cols=2, subplot_titles=["Distribution (Histogram)", "Box Plot"])
    fig.add_trace(go.Histogram(x=vals, nbinsx=30, marker_color=PALETTE[1], name="count"), row=1, col=1)
    fig.add_trace(go.Box(y=vals, marker_color=PALETTE[0], name=num_col), row=1, col=2)
    fig.update_layout(
        title=f"Distribution of {num_col}",
        showlegend=False, height=380,
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=70, b=40),
    )
    return {"id": f"dist_{num_col}", "title": f"Distribution: {num_col}", "fig": to_fig_json(fig)}


def chart_monthly_comparison(df, date_col, num_col, cat_col=None):
    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d[date_col].notna()]
    if len(d) == 0:
        return None
    d["month"] = d[date_col].dt.to_period("M").astype(str)
    if cat_col and d[cat_col].nunique() <= 12:
        pivot = d.groupby(["month", cat_col])[num_col].sum().unstack(fill_value=0)
        fig = go.Figure()
        for i, col in enumerate(pivot.columns):
            fig.add_trace(go.Bar(name=str(col), x=pivot.index, y=pivot[col],
                                 marker_color=PALETTE[i % len(PALETTE)]))
        fig.update_layout(barmode="group",
                          title=f"Monthly {num_col} by {cat_col}")
    else:
        g = d.groupby("month")[num_col].sum()
        fig = go.Figure(go.Bar(x=g.index, y=g.values,
                               marker_color=PALETTE[0],
                               text=[fmt_val(v) for v in g.values],
                               textposition="outside"))
        fig.update_layout(title=f"Monthly {num_col}")
    fig.update_layout(
        yaxis_title=num_col, xaxis_title="",
        plot_bgcolor="white", paper_bgcolor="white",
        height=420, margin=dict(t=60, b=60),
    )
    return {"id": "monthly", "title": "Monthly Comparison", "fig": to_fig_json(fig)}


def chart_numeric_summary_table(df, roles):
    num_cols = [c for c, r in roles.items() if r == "numeric"]
    if not num_cols:
        return None
    stats = df[num_cols].describe().T.round(2)
    fig = go.Figure(go.Table(
        header=dict(
            values=["Column"] + list(stats.columns),
            fill_color=PALETTE[0], font=dict(color="white", size=12),
            align="left",
        ),
        cells=dict(
            values=[stats.index] + [stats[c] for c in stats.columns],
            fill_color=[["#f0f4f8", "#ffffff"] * (len(stats) // 2 + 1)],
            align="left", font=dict(size=11),
        ),
    ))
    fig.update_layout(title="Numeric Columns — Summary Statistics",
                      height=max(300, len(num_cols) * 35 + 120),
                      paper_bgcolor="white", margin=dict(t=60, b=20))
    return {"id": "summary_table", "title": "Summary Statistics", "fig": to_fig_json(fig)}


def chart_category_mix_per_employee(df, emp_col, cat_col, num_col):
    pivot = df.groupby([emp_col, cat_col])[num_col].sum().unstack(fill_value=0)
    pct = pivot.div(pivot.sum(axis=1), axis=0) * 100
    fig = go.Figure()
    for i, col in enumerate(pct.columns):
        fig.add_trace(go.Bar(
            name=str(col), x=pct.index, y=pct[col],
            marker_color=PALETTE[i % len(PALETTE)],
            hovertemplate="%{x} — " + str(col) + ": %{y:.1f}%<extra></extra>",
        ))
    fig.update_layout(
        barmode="stack",
        title=f"Category Mix by {emp_col} (%)",
        yaxis_title="%", xaxis_title="",
        plot_bgcolor="white", paper_bgcolor="white",
        height=450, margin=dict(t=60, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
    )
    return {"id": "cat_mix_emp", "title": f"Category Mix by {emp_col}", "fig": to_fig_json(fig)}


# ─── main chart pipeline ──────────────────────────────────────────────────────

def generate_charts(df):
    roles = detect_columns(df)
    charts = []

    date_cols = [c for c, r in roles.items() if r == "date"]
    num_cols  = [c for c, r in roles.items() if r == "numeric"]
    cat_cols  = [c for c, r in roles.items() if r == "category"]
    text_cols = [c for c, r in roles.items() if r == "text"]

    # Coerce date columns
    for dc in date_cols:
        df[dc] = pd.to_datetime(df[dc], errors="coerce")

    # Pick "main" columns heuristically
    main_num = None
    for hint in ["amount", "total", "revenue", "дүн", "value", "price", "qty", "count", "тоо", "орлого"]:
        for c in num_cols:
            if hint in c.lower():
                main_num = c
                break
        if main_num:
            break
    if main_num is None and num_cols:
        # pick the one with the largest sum
        main_num = max(num_cols, key=lambda c: df[c].sum(skipna=True))

    qty_col = None
    for hint in ["qty", "тоо", "count", "quantity", "amount"]:
        for c in num_cols:
            if hint in c.lower() and c != main_num:
                qty_col = c
                break
        if qty_col:
            break
    if qty_col is None and len(num_cols) > 1:
        qty_col = [c for c in num_cols if c != main_num][0]

    emp_col = None
    for hint in ["staff", "employee", "waiter", "ажилтан", "person", "user", "name"]:
        for c in cat_cols + text_cols:
            if hint in c.lower():
                emp_col = c
                break
        if emp_col:
            break

    item_col = None
    for hint in ["item", "product", "бараа", "goods", "name", "нэр"]:
        for c in cat_cols + text_cols:
            if hint in c.lower() and c != emp_col:
                item_col = c
                break
        if item_col:
            break
    if item_col is None:
        # pick text col with most unique values
        candidates = [c for c in text_cols if c != emp_col]
        if candidates:
            item_col = max(candidates, key=lambda c: df[c].nunique())

    main_date = date_cols[0] if date_cols else None
    main_cat  = cat_cols[0] if cat_cols else None

    # 1. Summary statistics table
    tbl = chart_numeric_summary_table(df, roles)
    if tbl:
        charts.append(tbl)

    # 2. Daily trend
    if main_date and main_num:
        c = chart_trend(df, main_date, main_num)
        if c:
            charts.append(c)

    # 3. Cumulative
    if main_date and main_num:
        c = chart_cumulative(df, main_date, main_num)
        if c:
            charts.append(c)

    # 4. Revenue by day of week
    if main_date and main_num:
        c = chart_dow(df, main_date, main_num)
        if c:
            charts.append(c)

    # 5. Monthly comparison
    if main_date and main_num:
        c = chart_monthly_comparison(df, main_date, main_num, main_cat)
        if c:
            charts.append(c)

    # 6. Category bar + pie (for each categorical column)
    for cat in cat_cols[:4]:  # limit to 4
        if main_num:
            charts.append(chart_category_bar(df, cat, main_num))
            charts.append(chart_pie(df, cat, main_num))

    # 7. Top items
    if item_col and main_num:
        charts.append(chart_top_items(df, item_col, main_num))
    if item_col and qty_col and qty_col != main_num:
        charts.append(chart_top_items(df, item_col, qty_col))

    # 8. Employee charts
    if emp_col and main_num:
        if main_cat and main_cat != emp_col:
            charts.append(chart_category_mix_per_employee(df, emp_col, main_cat, main_num))
        charts.append(chart_employee_compare(df, emp_col, main_num, None))

    # 9. Diverging bar (if main_num has both positive & negative)
    if item_col and main_num:
        grp = df.groupby(item_col)[main_num].sum()
        if (grp > 0).any() and (grp < 0).any():
            charts.append(chart_diverging(df, item_col, main_num))
    elif main_cat and main_num:
        grp = df.groupby(main_cat)[main_num].sum()
        if (grp > 0).any() and (grp < 0).any():
            charts.append(chart_diverging(df, main_cat, main_num))

    # 10. Heatmap (week × category)
    if main_date and main_cat and main_num and df[main_date].notna().sum() > 7:
        c = chart_heatmap(df, main_date, main_cat, main_num)
        if c:
            charts.append(c)

    # 11. Treemap
    if main_cat and main_num:
        sub = item_col if item_col and item_col != main_cat else None
        if sub and df[sub].nunique() <= 200:
            charts.append(chart_treemap(df, main_cat, main_num, sub))
        else:
            charts.append(chart_treemap(df, main_cat, main_num, None))

    # 12. Scatter (if 2+ numeric columns)
    if len(num_cols) >= 2:
        c_col = main_cat if main_cat else None
        charts.append(chart_scatter(df, num_cols[0], num_cols[1], c_col))

    # 13. Distributions for main numeric columns
    for nc in num_cols[:2]:
        charts.append(chart_numeric_distribution(df, nc))

    return charts


# ─── routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/sheets", methods=["POST"])
def get_sheets():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls", ".xlsm")):
        return jsonify({"error": "Please upload an Excel file (.xlsx / .xls)"}), 400
    try:
        buf = f.read()
        xl = pd.ExcelFile(io.BytesIO(buf))
        sheets = xl.sheet_names
        # cache file in uploads
        save_path = os.path.join(UPLOAD_FOLDER, "last_upload.xlsx")
        with open(save_path, "wb") as fp:
            fp.write(buf)
        return jsonify({"sheets": sheets})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/preview", methods=["POST"])
def preview():
    data = request.get_json()
    sheet = data.get("sheet", 0)
    try:
        path = os.path.join(UPLOAD_FOLDER, "last_upload.xlsx")
        df = pd.read_excel(path, sheet_name=sheet)
        df.columns = df.columns.astype(str).str.strip()
        df = coerce_numerics(df)
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
    skip_rows = int(data.get("skip_rows", 0))
    try:
        path = os.path.join(UPLOAD_FOLDER, "last_upload.xlsx")
        df = pd.read_excel(path, sheet_name=sheet, skiprows=skip_rows)
        df.columns = df.columns.astype(str).str.strip()
        # drop fully empty columns/rows
        df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
        df = coerce_numerics(df)
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

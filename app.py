import os, io, json, warnings
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, render_template
import plotly.graph_objects as go

warnings.filterwarnings("ignore")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ── colours matching reference images ──────────────────────────────────────────
C_PREV   = "#4a9fd4"   # light blue  — previous year
C_CURR   = "#1a3c5e"   # dark navy   — current year
C_UP     = "#27ae60"   # green ▲
C_DOWN   = "#e74c3c"   # red ▼
C_GRID   = "#e8ecf0"
C_BG     = "#ffffff"
FONT     = "Inter, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif"

# ── category keyword map ────────────────────────────────────────────────────────
CATEGORY_MAP = {
    "Beer": {
        "kw": ["beer","pils","lager","ale","ipa","stout","porter","brew","бира",
                "chimay","westmalle","leffe","kwaremont","bavik","lindemans","corona",
                "heineken","stella","budweiser","petrus","bernardus","wittekerke","draft"],
        "sub": {
            "Craft Beer": ["chimay","westmalle","leffe","kwaremont","bernardus","wittekerke",
                           "lindemans","petrus","craft","abbey","trappist"],
            "Draft Beer":  ["draft"],
            "Guest Beer":  ["guest","import","corona","heineken","stella","budweiser"],
        },
    },
    "Wine": {
        "kw": ["wine","дарс","merlot","chardonnay","sauvignon","champagne","prosecco",
                "moet","ros","shiraz","cabernet","pinot","riesling","шампань","bardolino",
                "varaschin","cormons","poesie","terre forti","toso","moscato","trebbino"],
        "sub": {
            "Red Wine":    ["merlot","shiraz","cabernet","pinot noir","bardolino","улаан"],
            "White Wine":  ["chardonnay","sauvignon","riesling","pinot grigio","цагаан",
                            "cormons","moscato","trebbino","varaschin"],
            "Champagne":   ["champagne","prosecco","moet","sparkling","шампань","cava"],
            "Rosé":        ["ros","ягаан"],
        },
    },
    "Spirits": {
        "kw": ["vodka","gin","whiskey","whisky","cognac","brandy","tequila","rum",
                "scotch","bourbon","liqueur","absolut","grey goose","greygoose","beluga",
                "finlandia","soyombo","bombay","hendricks","gordons","gordon","tanqueray",
                "hennessy","jameson","patron","jose cuervo","sierra","glenmorange",
                "single malt","set","архи","спирт"],
        "sub": {
            "Vodka":    ["vodka","beluga","finlandia","grey goose","greygoose","absolut","soyombo"],
            "Gin":      ["gin","bombay","hendricks","gordons","gordon","tanqueray"],
            "Whiskey":  ["whiskey","whisky","jameson","scotch","bourbon","glenmorange","single malt"],
            "Cognac":   ["cognac","hennessy","brandy"],
            "Tequila":  ["tequila","patron","jose cuervo","sierra"],
            "Rum":      ["rum","bacardi"],
        },
    },
    "Cocktails": {
        "kw": ["cocktail","mojito","margarita","martini","negroni","sour","highball",
                "punch","belgian long island","pink dream","red diamond","tom collins",
                "коктейль","lemonade","лимонад","blackcurrant","mango lemon","strawberry lemon",
                "shake","cinamon"],
        "sub": {
            "Cocktails":  ["negroni","martini","manhattan","old fashioned","daiquiri",
                           "belgian long island","pink dream","red diamond","tom collins",
                           "whiskey sour","hennessy highball","special cocktail"],
            "Lemonade":   ["lemonade","лимонад","blackcurrant","mango lemon","strawberry lemon",
                           "punch","cinamon"],
            "Milkshake":  ["shake","milkshake"],
        },
    },
    "Soft Drinks": {
        "kw": ["soda","cola","pepsi","sprite","fanta","juice","water","ус","tonic",
                "ginger beer","redbull","red bull","schweppes","khujirt","millenia",
                "club soda","soft drink","langers","ice tea","iced tea","lemon tea",
                "шүүс","premium"],
        "sub": {
            "Water":       ["water","ус","khujirt","millenia"],
            "Soft Drinks": ["soda","cola","pepsi","sprite","fanta","schweppes","tonic",
                            "ginger beer","club soda","premium","soft drink"],
            "Juice":       ["juice","шүүс","langers","apple","mango","orange","ice tea","lemon tea"],
            "Energy":      ["redbull","red bull","energy"],
        },
    },
    "Hot Drinks": {
        "kw": ["tea","coffee","цай","кофе","espresso","latte","cappuccino","americano",
                "hot chocolate","althaus","халуун"],
        "sub": {
            "Tea":       ["tea","цай","althaus","green","black","berry"],
            "Coffee":    ["coffee","кофе","espresso","latte","cappuccino","americano"],
            "Other Hot": ["hot chocolate","халуун ундаа"],
        },
    },
    "Food": {
        "kw": ["pizza","steak","chicken","pork","beef","salmon","fish","pasta","fries",
                "salad","soup","wings","burger","platter","appetizer","starter","dessert",
                "хоол","пицца","goulash","carbonara","spareribs","margherita","meat lovers",
                "spicy salami","penne","ox bone","pork belly","chicken leg","fillet",
                "cobb","avocado","cheese","sausage","french fries","bacon","snack","grill"],
        "sub": {
            "Pizza":            ["pizza","margherita","meat lovers","spicy salami","пицца"],
            "Main Course":      ["steak","salmon","pork chop","pork belly","bbq","spareribs",
                                 "chicken leg","goulash","carbonara","penne","ox bone","fillet",
                                 "үндсэн"],
            "Starters & Sides": ["fries","wings","snack","platter","salad","cheese","sausage",
                                  "appetizer","starter","cobb","avocado","bacon","grill","share"],
            "Pasta":            ["pasta","carbonara","penne"],
            "Pub Food":         ["burger","sandwich","pub food"],
        },
    },
    "Karaoke & Shisha": {
        "kw": ["karaoke","shisha","hookah","карааке","шиша"],
        "sub": {
            "Karaoke": ["karaoke","карааке"],
            "Shisha":  ["shisha","hookah","шиша"],
        },
    },
}

# ── helpers ─────────────────────────────────────────────────────────────────────

def coerce_numerics(df):
    for col in df.columns:
        if df[col].dtype == object:
            cv = pd.to_numeric(df[col], errors="coerce")
            if cv.notna().mean() > 0.6:
                df[col] = cv
    return df


def detect_cols(df):
    """Return best-guess column names for date, item, qty, revenue, staff."""
    cols = {c.lower(): c for c in df.columns}
    def pick(*hints):
        for h in hints:
            for k, v in cols.items():
                if h in k:
                    return v
        return None

    date_col = pick("огноо","date","дата","datetime","time","өдөр")
    item_col = pick("бараа нэр","нэр","бараа","item","product","goods","name","барааны")
    qty_col  = pick("тоо","qty","quantity","count","ширхэг","amount")
    rev_col  = pick("нийт дүн","дүн","total","revenue","amount","орлого","price","sum")
    emp_col  = pick("ажилтан","staff","employee","waiter","cashier","хүн")
    cat_col  = pick("ангилал","category","type","бүлэг","group")

    # fallback: pick largest-sum numeric for revenue
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if rev_col is None and num_cols:
        rev_col = max(num_cols, key=lambda c: df[c].sum(skipna=True))
    if qty_col is None and num_cols:
        others = [c for c in num_cols if c != rev_col]
        qty_col = others[0] if others else rev_col

    return date_col, item_col, qty_col, rev_col, emp_col, cat_col


def assign_category(name: str):
    """Return (category, subcategory) for a product name string."""
    n = str(name).lower()
    for cat, meta in CATEGORY_MAP.items():
        if any(kw in n for kw in meta["kw"]):
            for sub, skws in meta["sub"].items():
                if any(kw in n for kw in skws):
                    return cat, sub
            return cat, cat   # no sub match → sub = cat itself
    return "Other", "Other"


def fmt(v):
    if pd.isna(v) or v == 0:
        return "0"
    av = abs(v)
    if av >= 1_000_000:
        return f"{v/1_000_000:.1f}M"
    if av >= 1_000:
        return f"{v/1_000:.0f}K"
    return f"{v:.0f}"


def pct_change(prev, curr):
    if prev and prev != 0:
        return (curr - prev) / abs(prev) * 100
    return 100.0 if curr else 0.0


def clear_uploads():
    i = 0
    while True:
        p = os.path.join(UPLOAD_FOLDER, f"upload_{i}.xlsx")
        if not os.path.exists(p):
            break
        os.remove(p)
        np = p.replace(".xlsx", ".name")
        if os.path.exists(np):
            os.remove(np)
        i += 1


def to_json(fig):
    return json.loads(fig.to_json())


def base_layout(**kw):
    d = dict(
        paper_bgcolor=C_BG, plot_bgcolor=C_BG,
        font=dict(family=FONT, color="#1e293b", size=12),
        title_font=dict(family=FONT, size=14, color="#0f172a"),
        legend=dict(
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor=C_GRID, borderwidth=1,
            font=dict(size=11),
        ),
        hoverlabel=dict(bgcolor="white", bordercolor=C_GRID,
                        font=dict(size=11, family=FONT)),
        margin=dict(t=90, b=60, l=60, r=40),
    )
    d.update(kw)
    return d


# ── chart builders ──────────────────────────────────────────────────────────────

def yoy_vertical(prev_vals, curr_vals, labels, title, subtitle,
                 metric_label, prev_lbl, curr_lbl, source_label="", top=20):
    """Grouped vertical bar with % change annotations — mirrors reference overview charts."""
    # sort by current value desc, cap at top
    paired = sorted(zip(curr_vals, prev_vals, labels), reverse=True)[:top]
    curr_vals, prev_vals, labels = zip(*paired) if paired else ([], [], [])
    curr_vals, prev_vals, labels = list(curr_vals), list(prev_vals), list(labels)

    annotations = []
    for i, (p, c) in enumerate(zip(prev_vals, curr_vals)):
        pct = pct_change(p, c)
        col = C_UP if pct >= 0 else C_DOWN
        arrow = "▲" if pct >= 0 else "▼"
        y_pos = max(p, c)
        annotations.append(dict(
            x=labels[i], y=y_pos,
            xref="x", yref="y",
            text=f"<b><span style='color:{col}'>{arrow}{abs(pct):.0f}%</span></b>",
            showarrow=False, yshift=10,
            font=dict(size=11, family=FONT),
            align="center",
        ))

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=prev_lbl, x=labels, y=prev_vals,
        marker=dict(color=C_PREV, line=dict(width=0)),
        text=[fmt(v) for v in prev_vals], textposition="outside",
        textfont=dict(size=10, color="#475569"),
        hovertemplate="<b>%{x}</b><br>" + prev_lbl + ": %{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name=curr_lbl, x=labels, y=curr_vals,
        marker=dict(color=C_CURR, line=dict(width=0)),
        text=[fmt(v) for v in curr_vals], textposition="outside",
        textfont=dict(size=10, color="#0f172a"),
        hovertemplate="<b>%{x}</b><br>" + curr_lbl + ": %{y:,.0f}<extra></extra>",
    ))

    full_title = f"<b>{title}</b><br><sup style='color:#64748b'>{subtitle}</sup>"
    height = max(480, len(labels) * 18 + 180)

    layout = base_layout(
        barmode="group",
        title=dict(text=full_title, x=0.5, xanchor="center", y=0.97),
        xaxis=dict(gridcolor=C_GRID, linecolor=C_GRID, tickfont=dict(size=11)),
        yaxis=dict(gridcolor=C_GRID, linecolor=C_GRID, title=metric_label,
                   tickfont=dict(size=10)),
        annotations=annotations,
        height=height,
        margin=dict(t=100, b=70, l=70, r=40),
        legend=dict(orientation="v", x=1.01, y=0.99, xanchor="left"),
    )
    if source_label:
        layout["annotations"] = layout.get("annotations", []) + [dict(
            text=f"<b>● {source_label}</b>",
            x=0, y=1.08, xref="paper", yref="paper",
            showarrow=False, font=dict(size=11, color=C_CURR),
            align="left",
        )]
    fig.update_layout(**layout)
    return fig


def yoy_horizontal(prev_vals, curr_vals, labels, title, subtitle,
                    metric_label, prev_lbl, curr_lbl, source_label=""):
    """Grouped horizontal bar with % on right — mirrors reference item-detail charts."""
    # sort by current value asc (so biggest is at top)
    paired = sorted(zip(curr_vals, prev_vals, labels))
    curr_vals, prev_vals, labels = zip(*paired) if paired else ([], [], [])
    curr_vals, prev_vals, labels = list(curr_vals), list(prev_vals), list(labels)

    annotations = []
    max_x = max(list(curr_vals) + list(prev_vals), default=1)
    for i, (p, c) in enumerate(zip(prev_vals, curr_vals)):
        pct = pct_change(p, c)
        col = C_UP if pct >= 0 else C_DOWN
        arrow = "▲" if pct >= 0 else "▼"
        annotations.append(dict(
            x=max_x * 1.02, y=labels[i],
            xref="x", yref="y",
            text=f"<b><span style='color:{col}'>{arrow}{abs(pct):.0f}%</span></b>",
            showarrow=False, xanchor="left",
            font=dict(size=10, family=FONT),
        ))

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=prev_lbl, y=labels, x=prev_vals, orientation="h",
        marker=dict(color=C_PREV, line=dict(width=0)),
        text=[fmt(v) for v in prev_vals], textposition="outside",
        textfont=dict(size=10, color="#475569"),
        hovertemplate="<b>%{y}</b><br>" + prev_lbl + ": %{x:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name=curr_lbl, y=labels, x=curr_vals, orientation="h",
        marker=dict(color=C_CURR, line=dict(width=0)),
        text=[fmt(v) for v in curr_vals], textposition="outside",
        textfont=dict(size=10, color="#0f172a"),
        hovertemplate="<b>%{y}</b><br>" + curr_lbl + ": %{x:,.0f}<extra></extra>",
    ))

    full_title = f"<b>{title}</b><br><sup style='color:#64748b'>{subtitle}</sup>"
    height = max(420, len(labels) * 42 + 120)

    layout = base_layout(
        barmode="group",
        title=dict(text=full_title, x=0.5, xanchor="center", y=0.97),
        xaxis=dict(gridcolor=C_GRID, linecolor=C_GRID, title=metric_label,
                   tickfont=dict(size=10), range=[0, max_x * 1.25]),
        yaxis=dict(gridcolor=C_GRID, linecolor=C_GRID, tickfont=dict(size=11)),
        annotations=annotations,
        height=height,
        margin=dict(t=100, b=60, l=160, r=80),
        legend=dict(orientation="v", x=1.0, y=0.01, xanchor="left"),
    )
    if source_label:
        layout["annotations"] = layout.get("annotations", []) + [dict(
            text=f"<b>● {source_label}</b>",
            x=0, y=1.10, xref="paper", yref="paper",
            showarrow=False, font=dict(size=11, color=C_CURR), align="left",
        )]
    fig.update_layout(**layout)
    return fig


def staff_revenue_chart(curr_df, emp_col, rev_col, curr_lbl, subtitle, source_label=""):
    """Staff revenue bar with total footer — mirrors sales_karaoke.png style."""
    g = curr_df.groupby(emp_col)[rev_col].sum().sort_values(ascending=False)
    total = g.sum()

    fig = go.Figure(go.Bar(
        x=g.index.astype(str).tolist(),
        y=g.values.tolist(),
        marker=dict(color=C_CURR, line=dict(width=0)),
        text=[fmt(v) for v in g.values],
        textposition="outside",
        textfont=dict(size=12, color="#0f172a"),
        hovertemplate="<b>%{x}</b><br>%{y:,.0f}<extra></extra>",
    ))

    title_text = (f"<b>Staff Revenue — {source_label + ' | ' if source_label else ''}"
                  f"{curr_lbl}</b><br><sup style='color:#64748b'>{subtitle}</sup>")

    fig.update_layout(**base_layout(
        title=dict(text=title_text, x=0.5, xanchor="center", y=0.95),
        xaxis=dict(gridcolor=C_GRID, linecolor=C_GRID, tickfont=dict(size=12)),
        yaxis=dict(gridcolor=C_GRID, linecolor=C_GRID,
                   tickfont=dict(size=10), title="Revenue"),
        height=520,
        margin=dict(t=100, b=120, l=60, r=40),
        annotations=[
            dict(
                text=(f"<b style='color:#94a3b8;font-size:10px;letter-spacing:1px'>"
                      f"TOTAL REVENUE — {curr_lbl.upper()}</b><br>"
                      f"<b style='font-size:22px;color:white'>{fmt(total)} ₮</b>"),
                x=0.5, y=-0.25, xref="paper", yref="paper",
                showarrow=False, align="center",
                bgcolor=C_CURR, bordercolor=C_CURR,
                borderpad=14, borderwidth=0,
                font=dict(family=FONT),
            )
        ],
    ))
    if source_label:
        fig.add_annotation(
            text=f"<b>● {source_label}</b>",
            x=0, y=1.08, xref="paper", yref="paper",
            showarrow=False, font=dict(size=11, color=C_CURR), align="left",
        )
    return fig


def daily_trend_chart(prev_df, curr_df, date_col, rev_col, prev_lbl, curr_lbl, subtitle):
    """Line chart overlaying two years' daily revenue."""
    def daily(df):
        d = df.copy()
        d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
        d = d.dropna(subset=[date_col])
        d["day"] = d[date_col].dt.day
        return d.groupby("day")[rev_col].sum()

    p = daily(prev_df) if prev_df is not None else pd.Series(dtype=float)
    c = daily(curr_df)

    fig = go.Figure()
    if not p.empty:
        fig.add_trace(go.Scatter(
            x=p.index.tolist(), y=p.values.tolist(),
            name=prev_lbl, mode="lines+markers",
            line=dict(color=C_PREV, width=2.5, shape="spline", smoothing=0.7),
            marker=dict(size=5),
            hovertemplate="Day %{x}<br>" + prev_lbl + ": %{y:,.0f}<extra></extra>",
        ))
    fig.add_trace(go.Scatter(
        x=c.index.tolist(), y=c.values.tolist(),
        name=curr_lbl, mode="lines+markers",
        line=dict(color=C_CURR, width=2.5, shape="spline", smoothing=0.7),
        marker=dict(size=5),
        hovertemplate="Day %{x}<br>" + curr_lbl + ": %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(**base_layout(
        title=dict(
            text=f"<b>Daily Revenue Trend</b><br><sup style='color:#64748b'>{subtitle}</sup>",
            x=0.5, xanchor="center",
        ),
        xaxis=dict(gridcolor=C_GRID, title="Day of Month"),
        yaxis=dict(gridcolor=C_GRID, title="Revenue"),
        hovermode="x unified",
        height=420,
        legend=dict(orientation="h", y=1.08, x=1, xanchor="right"),
    ))
    return fig


# ── main pipeline ───────────────────────────────────────────────────────────────

def generate_all_charts(df, source_label=""):
    charts = []
    date_col, item_col, qty_col, rev_col, emp_col, cat_col = detect_cols(df)

    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    # ── assign categories ──
    if item_col and cat_col is None:
        df["_cat"], df["_sub"] = zip(*df[item_col].fillna("").map(assign_category))
        cat_col, sub_col = "_cat", "_sub"
    elif cat_col:
        df["_cat"] = df[cat_col]
        df["_sub"] = df[cat_col]  # no sub info
        sub_col = "_sub"
    else:
        df["_cat"] = "All"
        df["_sub"] = "All"
        cat_col, sub_col = "_cat", "_sub"

    # ── detect periods ──
    prev_df, curr_df, prev_lbl, curr_lbl = None, df, "", ""
    date_range = ""
    has_yoy = False

    if date_col and df[date_col].notna().any():
        years = sorted(df[date_col].dt.year.dropna().unique().astype(int))
        if len(years) >= 2:
            py, cy = years[-2], years[-1]
            prev_df = df[df[date_col].dt.year == py].copy()
            curr_df = df[df[date_col].dt.year == cy].copy()
            # detect month
            months = df[date_col].dt.month.dropna().unique()
            if len(months) == 1:
                mname = pd.Timestamp(2000, int(months[0]), 1).strftime("%B")
                prev_lbl = f"{py} {mname}"
                curr_lbl = f"{cy} {mname}"
            else:
                prev_lbl = str(py)
                curr_lbl = str(cy)
            dr_start = df[date_col].min().strftime("%m.%d")
            dr_end   = df[date_col].max().strftime("%m.%d")
            date_range = f"{dr_start} – {dr_end}"
            has_yoy = True
        else:
            cy = years[0]
            curr_lbl = str(cy)
            months = df[date_col].dt.month.dropna().unique()
            if len(months) == 1:
                mname = pd.Timestamp(2000, int(months[0]), 1).strftime("%B")
                curr_lbl = f"{cy} {mname}"
            dr_start = df[date_col].min().strftime("%m.%d")
            dr_end   = df[date_col].max().strftime("%m.%d")
            date_range = f"{dr_start} – {dr_end}"
    else:
        curr_lbl = "Current"

    subtitle = (f"{prev_lbl} vs. {curr_lbl} | {date_range}"
                if has_yoy else f"{curr_lbl} | {date_range}")

    def grp_rev(d, col):
        return d.groupby(col)[rev_col].sum() if d is not None and rev_col else pd.Series(dtype=float)

    def grp_qty(d, col):
        return d.groupby(col)[qty_col].sum() if d is not None and qty_col else pd.Series(dtype=float)

    def make_yoy(prev_s, curr_s, title, metric_label, orientation="v"):
        cats = list(curr_s.index.union(prev_s.index if prev_s is not None else []))
        cv = [float(curr_s.get(c, 0)) for c in cats]
        pv = [float(prev_s.get(c, 0) if prev_s is not None else 0) for c in cats]
        if orientation == "v":
            return yoy_vertical(pv, cv, cats, title, subtitle,
                                metric_label, prev_lbl, curr_lbl, source_label)
        else:
            return yoy_horizontal(pv, cv, cats, title, subtitle,
                                  metric_label, prev_lbl, curr_lbl, source_label)

    # 1. ── Revenue overview by category ──────────────────────────────────────
    if rev_col:
        curr_cat_rev = grp_rev(curr_df, "_cat")
        prev_cat_rev = grp_rev(prev_df, "_cat") if has_yoy else None
        curr_total = curr_cat_rev.sum()
        prev_total = prev_cat_rev.sum() if prev_cat_rev is not None else 0
        overall_pct = pct_change(prev_total, curr_total)

        title_rev = (f"SALES REPORT — {source_label.upper() if source_label else 'TOTAL'}"
                     f"{'  ▲' if overall_pct >= 0 else '  ▼'}{abs(overall_pct):.1f}%  |  "
                     f"{fmt(prev_total)} ₮  →  {fmt(curr_total)} ₮")

        fig = make_yoy(prev_cat_rev, curr_cat_rev, title_rev, "Revenue (₮)")
        charts.append({"id": "rev_overview", "title": "Revenue by Category", "fig": to_json(fig)})

    # 2. ── Staff revenue ──────────────────────────────────────────────────────
    if emp_col and rev_col:
        fig = staff_revenue_chart(curr_df, emp_col, rev_col, curr_lbl, subtitle, source_label)
        charts.append({"id": "staff_rev", "title": "Staff Revenue", "fig": to_json(fig)})

    # 3. ── Daily trend ────────────────────────────────────────────────────────
    if date_col and rev_col:
        fig = daily_trend_chart(prev_df if has_yoy else None,
                                 curr_df, date_col, rev_col,
                                 prev_lbl, curr_lbl, subtitle)
        charts.append({"id": "daily_trend", "title": "Daily Revenue Trend", "fig": to_json(fig)})

    # 4. ── Per-category: quantity overview (sub-categories) ──────────────────
    if qty_col:
        categories = [c for c in curr_df["_cat"].unique() if c != "Other"]
        for cat in sorted(categories):
            c_data = curr_df[curr_df["_cat"] == cat]
            p_data = prev_df[prev_df["_cat"] == cat] if has_yoy and prev_df is not None else None

            curr_sub = grp_qty(c_data, "_sub")
            prev_sub = grp_qty(p_data, "_sub") if p_data is not None else pd.Series(dtype=float)

            if curr_sub.empty or curr_sub.sum() == 0:
                continue

            # show sub-category overview only if >1 sub
            if len(curr_sub) > 1:
                fig = make_yoy(prev_sub, curr_sub,
                               f"{cat} — {source_label + ' | ' if source_label else ''}"
                               f"{prev_lbl + ' vs ' + curr_lbl if has_yoy else curr_lbl} (Qty)",
                               "Quantity", "v")
                charts.append({"id": f"cat_{cat}", "title": f"{cat} Overview", "fig": to_json(fig)})

            # 5. ── Per sub-category: individual items (horizontal) ──────────
            subs = [s for s in c_data["_sub"].unique()]
            for sub in sorted(subs):
                c_items = c_data[c_data["_sub"] == sub]
                p_items = p_data[p_data["_sub"] == sub] if p_data is not None else None

                if item_col is None:
                    continue
                curr_items = c_items.groupby(item_col)[qty_col].sum()
                prev_items = p_items.groupby(item_col)[qty_col].sum() if p_items is not None else pd.Series(dtype=float)

                # skip if only 1 item or too small
                if len(curr_items) < 2 or curr_items.sum() == 0:
                    continue

                # top 15 items
                top_items = curr_items.nlargest(15).index
                curr_items = curr_items.reindex(top_items, fill_value=0)
                prev_items = prev_items.reindex(top_items, fill_value=0)

                fig = make_yoy(
                    prev_items, curr_items,
                    f"{sub} — {source_label + ' | ' if source_label else ''}"
                    f"{prev_lbl + ' vs ' + curr_lbl if has_yoy else curr_lbl} (Qty)",
                    "Quantity", "h",
                )
                charts.append({"id": f"sub_{cat}_{sub}", "title": f"{sub} Detail", "fig": to_json(fig)})

    return charts


# ── routes ───────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
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
        with open(path.replace(".xlsx", ".name"), "w") as fn:
            fn.write(f.filename)
        saved.append(f.filename)
    if not saved:
        return jsonify({"error": "No valid Excel files found"}), 400
    first = pd.ExcelFile(os.path.join(UPLOAD_FOLDER, "upload_0.xlsx"))
    return jsonify({"sheets": first.sheet_names, "files": saved})


@app.route("/api/preview", methods=["POST"])
def preview():
    data = request.get_json()
    sheet = data.get("sheet", 0)
    try:
        frames = _load_files(sheet)
        df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        df = coerce_numerics(df)
        date_col, item_col, qty_col, rev_col, emp_col, _ = detect_cols(df)
        roles = {}
        for col in df.columns:
            s = df[col].dropna()
            if pd.api.types.is_datetime64_any_dtype(s):
                roles[col] = "date"
            elif col == date_col:
                roles[col] = "date"
            elif pd.api.types.is_numeric_dtype(s):
                roles[col] = "numeric"
            elif s.nunique() <= 50:
                roles[col] = "category"
            else:
                roles[col] = "text"
        detected = {k: v for k, v in {
            "date": date_col, "item": item_col,
            "qty": qty_col, "revenue": rev_col, "staff": emp_col
        }.items() if v}
        sample = df.head(5).fillna("").astype(str).to_dict("records")
        return jsonify({
            "columns": list(df.columns), "roles": roles,
            "shape": [int(df.shape[0]), int(df.shape[1])],
            "sample": sample, "detected": detected,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/charts", methods=["POST"])
def charts():
    data = request.get_json()
    sheet = data.get("sheet", 0)
    try:
        frames = _load_files(sheet)
        if not frames:
            return jsonify({"error": "No data loaded"}), 400

        # If multiple files, try to keep source labels per file
        if len(frames) == 1:
            df = frames[0]
            source = _source_label(0)
            df = coerce_numerics(df)
            result = generate_all_charts(df, source)
        else:
            # Merge all files — auto-detect years across them
            for i, f in enumerate(frames):
                f["_file_source"] = _source_label(i)
                coerce_numerics(f)
            df = pd.concat(frames, ignore_index=True)
            result = generate_all_charts(df, "")

        return jsonify({"charts": result})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


def _load_files(sheet):
    frames = []
    i = 0
    while True:
        p = os.path.join(UPLOAD_FOLDER, f"upload_{i}.xlsx")
        if not os.path.exists(p):
            break
        try:
            xl = pd.ExcelFile(p)
            sname = xl.sheet_names[sheet] if isinstance(sheet, int) and sheet < len(xl.sheet_names) else xl.sheet_names[0]
            df = pd.read_excel(p, sheet_name=sname)
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
            frames.append(df)
        except Exception:
            pass
        i += 1
    return frames


def _source_label(i):
    p = os.path.join(UPLOAD_FOLDER, f"upload_{i}.xlsx")
    # read original filename from a sidecar if stored, else return empty
    name_p = p.replace(".xlsx", ".name")
    if os.path.exists(name_p):
        with open(name_p) as f:
            n = f.read().strip()
        # extract useful part e.g. "karaoke_2026.xlsx" → "Karaoke"
        base = os.path.splitext(n)[0]
        for kw in ["karaoke","restaurant","bar","cafe","kitchen"]:
            if kw in base.lower():
                return kw.capitalize()
    return ""


if __name__ == "__main__":
    print("=" * 60)
    print("  Brussels Report Maker")
    print("  Open http://localhost:8080 in your browser")
    print("=" * 60)
    app.run(debug=False, port=8080, host="0.0.0.0")

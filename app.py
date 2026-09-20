"""
ScaleForce Capital — Deal Lifecycle Dashboard
Upload the Deal Pipeline & Lifecycle Checklist workbook and get a live dashboard
plus a branded, self-contained HTML report.

Run locally:  streamlit run app.py
"""

import base64
import datetime as dt
import io
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

# ----------------------------------------------------------------- branding --
NAVY = "#0A1628"
GOLD = "#C9A84C"
INK = "#1F2933"
MUTED = "#6B7280"
RED = "#B3261E"
GREEN = "#2E7D32"
AMBER = "#B26A00"
PAPER = "#FFFFFF"
WASH = "#F5F6F8"

DONE = "Done"
NA = "N/A"
SLA_DAYS = 7

st.set_page_config(page_title="ScaleForce — Deal Lifecycle Dashboard",
                   page_icon="◆", layout="wide")

st.markdown(f"""
<style>
  .stApp {{ background: {WASH}; }}
  html, body, [class*="css"] {{ font-family: Arial, Helvetica, sans-serif; color: {INK}; }}
  .sf-head {{ background:{NAVY}; padding:22px 26px; border-radius:6px;
              border-left:6px solid {GOLD}; margin-bottom:18px; }}
  .sf-head h1 {{ color:#fff; font-size:25px; margin:0; font-weight:700; letter-spacing:.2px; }}
  .sf-head p  {{ color:{GOLD}; margin:5px 0 0; font-size:13px; letter-spacing:1.4px;
                 text-transform:uppercase; }}
  .kpi {{ background:{PAPER}; border:1px solid #E3E6EA; border-top:3px solid {GOLD};
          border-radius:5px; padding:14px 16px; height:100%; }}
  .kpi .lab {{ font-size:11px; text-transform:uppercase; letter-spacing:.9px; color:{MUTED}; }}
  .kpi .val {{ font-size:26px; font-weight:700; color:{NAVY}; margin-top:4px; line-height:1.1; }}
  .kpi .sub {{ font-size:11px; color:{MUTED}; margin-top:2px; }}
  .sf-sec {{ color:{NAVY}; font-size:17px; font-weight:700; margin:22px 0 8px;
             border-bottom:2px solid {GOLD}; padding-bottom:5px; }}
  section[data-testid="stSidebar"] {{ background:{NAVY}; }}
  section[data-testid="stSidebar"] * {{ color:#E8EAED !important; }}
  div[data-testid="stMetricValue"] {{ color:{NAVY}; }}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------ parsing --
@st.cache_data(show_spinner=False)
def load_workbook_bytes(raw: bytes):
    """Parse the Pipeline and Deal Checklist sheets out of the uploaded workbook."""
    bio = io.BytesIO(raw)
    sheets = pd.read_excel(bio, sheet_name=None, header=None)

    def find(*candidates):
        for name in sheets:
            if name.strip().lower() in candidates:
                return sheets[name]
        return None

    raw_pipe = find("pipeline")
    raw_chk = find("deal checklist", "checklist", "deal_checklist")
    if raw_pipe is None or raw_chk is None:
        raise ValueError(
            "Could not find both a 'Pipeline' and a 'Deal Checklist' sheet in this workbook."
        )

    pipe = _header_frame(raw_pipe, "Deal ID")
    chk = _header_frame(raw_chk, "Deal ID")
    return pipe, chk


def _header_frame(df: pd.DataFrame, anchor: str) -> pd.DataFrame:
    """Locate the header row by its anchor label, then rebuild the frame beneath it."""
    hdr = None
    for i in range(min(15, len(df))):
        row = [str(v).strip() for v in df.iloc[i].tolist()]
        if anchor in row:
            hdr = i
            break
    if hdr is None:
        raise ValueError(f"No header row containing '{anchor}' was found.")
    out = df.iloc[hdr + 1:].copy()
    out.columns = [str(v).strip() for v in df.iloc[hdr].tolist()]
    out = out.loc[:, [c for c in out.columns if c and c.lower() != "nan"]]
    out = out[out["Deal ID"].notna()]
    out["Deal ID"] = out["Deal ID"].astype(str).str.strip()
    return out.reset_index(drop=True)


def as_date(v):
    if pd.isna(v) or v == "":
        return None
    try:
        d = pd.to_datetime(v, errors="coerce")
        return None if pd.isna(d) else d.date()
    except Exception:
        return None


def as_num(v):
    if pd.isna(v) or v == "":
        return 0.0
    if isinstance(v, str):
        v = re.sub(r"[^\d.\-]", "", v)
        if v in ("", "-", "."):
            return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0


def stage_no(stage):
    m = re.match(r"\s*(\d+)", str(stage))
    return int(m.group(1)) if m else None


# --------------------------------------------------------------- derivation --
def build_model(pipe: pd.DataFrame, chk: pd.DataFrame, today: dt.date):
    """Recompute every progress metric from the checklist rather than trusting cached formulas."""
    chk = chk.copy()
    chk["Status"] = chk.get("Status", "").fillna("Not started").astype(str).str.strip()
    chk["Target date"] = chk.get("Target date").map(as_date) if "Target date" in chk else None
    chk["Completed date"] = chk.get("Completed date").map(as_date) if "Completed date" in chk else None
    chk["_in_scope"] = chk["Status"] != NA
    chk["_done"] = chk["Status"] == DONE
    chk["_overdue"] = (
        (~chk["_done"]) & chk["_in_scope"]
        & chk["Target date"].map(lambda d: d is not None and d < today)
    )
    chk["_days_late"] = chk.apply(
        lambda r: (today - r["Target date"]).days if r["_overdue"] else 0, axis=1)

    rows = []
    for _, p in pipe.iterrows():
        did = p["Deal ID"]
        sub = chk[chk["Deal ID"] == did]
        scope = int(sub["_in_scope"].sum())
        done = int(sub["_done"].sum())
        stage = str(p.get("Current stage", "") or "").strip()
        received = as_date(p.get("Date received"))
        sno = stage_no(stage)
        cur = sub[(sub["Stage"].astype(str).str.strip() == stage) & sub["_in_scope"] & ~sub["_done"]] \
            if "Stage" in sub else sub.iloc[0:0]

        if received is None or sno is None:
            sla = "—"
        elif sno > 2:
            sla = "Passed"
        elif today > received + dt.timedelta(days=SLA_DAYS):
            sla = "BREACH"
        else:
            sla = "On track"

        rows.append({
            "Deal ID": did,
            "Client": p.get("Client / entity", ""),
            "Product": p.get("Product", ""),
            "Amount": as_num(p.get("Funding amount (R)")),
            "Channel": p.get("Application channel", ""),
            "Received": received,
            "Stage": stage,
            "Stage #": sno,
            "IM": p.get("Investment Manager", ""),
            "AM": p.get("AM", ""),
            "Status": str(p.get("Deal status", "") or "").strip(),
            "Next action": p.get("Next action", ""),
            "Next action date": as_date(p.get("Next action date")),
            "SLA": sla,
            "SLA due": received + dt.timedelta(days=SLA_DAYS) if received else None,
            "Age (days)": (today - received).days if received else None,
            "In scope": scope,
            "Done": done,
            "Outstanding": scope - done,
            "% complete": (done / scope) if scope else 0.0,
            "Overdue": int(sub["_overdue"].sum()),
            "Outstanding in stage": int(len(cur)),
            "Max days late": int(sub["_days_late"].max()) if len(sub) else 0,
        })
    return pd.DataFrame(rows), chk


def rands(v):
    return f"R {v:,.0f}".replace(",", " ")


# -------------------------------------------------------------- reporting ----
def build_report(deals_v, chk_v, firm, author, asof, scope_note):
    def badge(text, color):
        return (f'<span style="background:{color};color:#fff;padding:2px 8px;border-radius:10px;'
                f'font-size:11px;font-weight:700;">{text}</span>')

    def sla_badge(s):
        return {"BREACH": badge("BREACH", RED), "On track": badge("On track", GREEN),
                "Passed": badge("Passed", MUTED)}.get(s, "—")

    reg = []
    for _, r in deals_v.iterrows():
        bar = (f'<div style="background:#E8EAED;border-radius:3px;height:9px;width:90px;">'
               f'<div style="background:{GOLD};height:9px;border-radius:3px;'
               f'width:{r["% complete"]*90:.0f}px;"></div></div>')
        reg.append(f"""<tr>
<td><b>{r['Deal ID']}</b></td><td>{r['Client']}</td><td>{r['Product']}</td>
<td class="num">{rands(r['Amount'])}</td><td>{r['Stage']}</td>
<td class="num">{'' if r['Age (days)'] is None else int(r['Age (days)'])}</td>
<td>{sla_badge(r['SLA'])}</td>
<td>{bar}<span style="font-size:11px;">{r['% complete']:.0%}</span></td>
<td class="num">{int(r['Outstanding'])}</td>
<td class="num" style="color:{RED if r['Overdue'] else INK};font-weight:{700 if r['Overdue'] else 400}">
{int(r['Overdue'])}</td></tr>""")

    ex = chk_v[chk_v["_overdue"]].sort_values("_days_late", ascending=False)
    if ex.empty:
        exc_html = '<p class="ok">No overdue checklist items at the reporting date.</p>'
    else:
        rows = "".join(
            f"<tr><td><b>{r['Deal ID']}</b></td><td>{r.get('Ref','')}</td><td>{r.get('Task','')}</td>"
            f"<td>{r.get('Responsible','')}</td><td>{r.get('Status','')}</td>"
            f"<td>{r['Target date']}</td><td class='num' style='color:{RED};font-weight:700'>"
            f"{int(r['_days_late'])}</td></tr>" for _, r in ex.iterrows())
        exc_html = ("<table><thead><tr><th>Deal</th><th>Ref</th><th>Task</th><th>Responsible</th>"
                    f"<th>Status</th><th>Target</th><th>Days late</th></tr></thead><tbody>{rows}</tbody></table>")

    stage_rows = ""
    for s in sorted(set(deals_v["Stage"]), key=lambda x: (stage_no(x) or 99)):
        sel = deals_v[deals_v["Stage"] == s]
        stage_rows += (f"<tr><td>{s}</td><td class='num'>{len(sel)}</td>"
                       f"<td class='num'>{rands(sel['Amount'].sum())}</td>"
                       f"<td class='num'>{int(sel['Outstanding'].sum())}</td></tr>")

    k = [("Deals", len(deals_v)), ("Value", rands(deals_v["Amount"].sum())),
         ("SLA breaches", int((deals_v["SLA"] == "BREACH").sum())),
         ("Overdue tasks", int(deals_v["Overdue"].sum())),
         ("Avg completion", f"{deals_v['% complete'].mean():.0%}" if len(deals_v) else "—")]
    kpi_html = "".join(
        f'<div class="k"><div class="kl">{a}</div><div class="kv">{b}</div></div>' for a, b in k)

    return f"""<!doctype html><html><head><meta charset="utf-8">
<title>Deal Lifecycle Report — {asof}</title>
<style>
@page {{ size: A4 landscape; margin: 14mm; }}
body {{ font-family: Arial, Helvetica, sans-serif; color:{INK}; font-size:11px; margin:0; background:#fff; }}
.hd {{ background:{NAVY}; color:#fff; padding:20px 24px; border-left:7px solid {GOLD}; }}
.hd h1 {{ margin:0; font-size:21px; }}
.hd .s {{ color:{GOLD}; font-size:11px; letter-spacing:1.5px; text-transform:uppercase; margin-top:4px; }}
.hd .m {{ color:#C9CDD3; font-size:11px; margin-top:8px; }}
.wrap {{ padding:18px 24px 30px; }}
.kpis {{ display:flex; gap:10px; margin-bottom:18px; }}
.k {{ flex:1; border:1px solid #E3E6EA; border-top:3px solid {GOLD}; padding:9px 11px; }}
.kl {{ font-size:9.5px; text-transform:uppercase; letter-spacing:.8px; color:{MUTED}; }}
.kv {{ font-size:19px; font-weight:700; color:{NAVY}; margin-top:2px; }}
h2 {{ color:{NAVY}; font-size:13px; border-bottom:2px solid {GOLD}; padding-bottom:4px;
      margin:20px 0 8px; page-break-after:avoid; }}
table {{ border-collapse:collapse; width:100%; }}
th {{ background:{NAVY}; color:#fff; text-align:left; padding:6px 8px; font-size:10px; }}
td {{ border-bottom:1px solid #E8EAED; padding:6px 8px; vertical-align:top; }}
tr:nth-child(even) td {{ background:{WASH}; }}
.num {{ text-align:right; }}
.ok {{ color:{GREEN}; font-weight:700; }}
.ft {{ margin-top:24px; border-top:1px solid #E3E6EA; padding-top:8px;
       color:{MUTED}; font-size:9.5px; }}
</style></head><body>
<div class="hd"><h1>Deal Lifecycle Report</h1>
<div class="s">{firm}</div>
<div class="m">As at {asof:%d %B %Y} &nbsp;·&nbsp; Prepared by {author} &nbsp;·&nbsp; {scope_note}</div></div>
<div class="wrap">
<div class="kpis">{kpi_html}</div>
<h2>Deal register</h2>
<table><thead><tr><th>Deal</th><th>Client</th><th>Product</th><th class="num">Amount</th>
<th>Stage</th><th class="num">Age (d)</th><th>Screening SLA</th><th>Progress</th>
<th class="num">Open</th><th class="num">Overdue</th></tr></thead><tbody>{''.join(reg)}</tbody></table>
<h2>Distribution by stage</h2>
<table><thead><tr><th>Stage</th><th class="num">Deals</th><th class="num">Value</th>
<th class="num">Outstanding tasks</th></tr></thead><tbody>{stage_rows}</tbody></table>
<h2>Exceptions — overdue checklist items</h2>
{exc_html}
<div class="ft">Generated from the Deal Pipeline &amp; Lifecycle Checklist workbook. Process basis:
Business Partners Limited business process flow. Screening SLA measured as {SLA_DAYS} calendar days from
date of application to client meeting. Items marked N/A are excluded from completion percentages.
This report is internal and unaudited.</div>
</div></body></html>"""



# ----------------------------------------------------------------- sidebar ---
with st.sidebar:
    st.markdown("### Deal Lifecycle Dashboard")
    st.caption("ScaleForce Capital (Pty) Ltd")
    up = st.file_uploader("Upload checklist workbook (.xlsx)", type=["xlsx"])
    st.divider()
    firm = st.text_input("Report entity", "ScaleForce Capital (Pty) Ltd")
    author = st.text_input("Prepared by", "Jaques Davidson")
    asof = st.date_input("Report as at", dt.date.today())
    st.divider()
    st.caption("Metrics are recomputed from the Deal Checklist tab, so they stay correct even "
               "if Excel has not recalculated the workbook.")

st.markdown(
    f'<div class="sf-head"><h1>Deal Lifecycle Dashboard</h1>'
    f'<p>Origination &nbsp;→&nbsp; Screening &nbsp;→&nbsp; Due diligence &nbsp;→&nbsp; '
    f'Approval &nbsp;→&nbsp; Payout</p></div>', unsafe_allow_html=True)

if up is None:
    st.info("Upload the **Deal Pipeline & Lifecycle Checklist** workbook in the sidebar to begin. "
            "The app reads the `Pipeline` and `Deal Checklist` sheets.")
    st.stop()

try:
    pipe_raw, chk_raw = load_workbook_bytes(up.getvalue())
    deals, chk = build_model(pipe_raw, chk_raw, asof)
except Exception as e:
    st.error(f"Could not read that workbook: {e}")
    st.stop()

if deals.empty:
    st.warning("No deals found on the Pipeline sheet.")
    st.stop()

# ------------------------------------------------------------------ filters --
f1, f2, f3 = st.columns([2, 2, 3])
statuses = sorted([s for s in deals["Status"].unique() if s])
pick_status = f1.multiselect("Deal status", statuses,
                             default=[s for s in statuses if s in ("Active", "On hold", "Approved")] or statuses)
stages = sorted([s for s in deals["Stage"].unique() if s], key=lambda x: (stage_no(x) or 99))
pick_stage = f2.multiselect("Stage", stages, default=stages)
only_flag = f3.checkbox("Show only deals with an SLA breach or overdue tasks", value=False)

view = deals[deals["Status"].isin(pick_status) & deals["Stage"].isin(pick_stage)]
if only_flag:
    view = view[(view["SLA"] == "BREACH") | (view["Overdue"] > 0)]

# --------------------------------------------------------------------- KPIs --
active_val = view["Amount"].sum()
avg_done = view["% complete"].mean() if len(view) else 0
kpis = [
    ("Deals in view", f"{len(view)}", f"{len(deals)} captured in total"),
    ("Value in view", rands(active_val), "Sum of funding amounts"),
    ("SLA breaches", f"{int((view['SLA'] == 'BREACH').sum())}", f"{SLA_DAYS}-day screening gate"),
    ("Overdue tasks", f"{int(view['Overdue'].sum())}", "Past target date, not done"),
    ("Average completion", f"{avg_done:.0%}", "Checklist items closed"),
]
for col, (lab, val, sub) in zip(st.columns(len(kpis)), kpis):
    col.markdown(f'<div class="kpi"><div class="lab">{lab}</div>'
                 f'<div class="val">{val}</div><div class="sub">{sub}</div></div>',
                 unsafe_allow_html=True)

# ------------------------------------------------------------------ charts ---
st.markdown('<div class="sf-sec">Pipeline by stage</div>', unsafe_allow_html=True)
c1, c2 = st.columns([3, 2])

order = sorted(set(list(view["Stage"].dropna())), key=lambda x: (stage_no(x) or 99))
counts = [int((view["Stage"] == s).sum()) for s in order]
values = [float(view.loc[view["Stage"] == s, "Amount"].sum()) for s in order]

fig = go.Figure(go.Bar(
    x=counts, y=order, orientation="h", marker_color=NAVY,
    text=[f"{c} · {rands(v)}" for c, v in zip(counts, values)],
    textposition="outside", cliponaxis=False,
    hovertemplate="%{y}<br>%{x} deals<extra></extra>"))
fig.update_layout(height=max(280, 46 * len(order)), margin=dict(l=0, r=90, t=10, b=10),
                  plot_bgcolor=PAPER, paper_bgcolor=PAPER,
                  xaxis=dict(title="Deals", showgrid=True, gridcolor="#E8EAED", dtick=1),
                  yaxis=dict(autorange="reversed"), font=dict(family="Arial", color=INK))
c1.plotly_chart(fig, use_container_width=True)

prog = view.sort_values("% complete")
fig2 = go.Figure(go.Bar(
    x=(prog["% complete"] * 100).round(0), y=prog["Deal ID"], orientation="h",
    marker_color=[RED if o > 0 else GOLD for o in prog["Overdue"]],
    text=[f"{p:.0%}" for p in prog["% complete"]], textposition="outside", cliponaxis=False,
    hovertemplate="%{y}: %{x}%<extra></extra>"))
fig2.update_layout(height=max(280, 46 * len(prog)), margin=dict(l=0, r=50, t=10, b=10),
                   plot_bgcolor=PAPER, paper_bgcolor=PAPER,
                   xaxis=dict(title="% of checklist complete", range=[0, 108],
                              showgrid=True, gridcolor="#E8EAED"),
                   font=dict(family="Arial", color=INK))
c2.markdown("**Completion by deal** — red where tasks are overdue")
c2.plotly_chart(fig2, use_container_width=True)

# ------------------------------------------------------------------- table ---
st.markdown('<div class="sf-sec">Deal register</div>', unsafe_allow_html=True)
tbl = view[["Deal ID", "Client", "Product", "Amount", "Stage", "Status", "Received",
            "Age (days)", "SLA", "Done", "In scope", "% complete", "Overdue",
            "Outstanding in stage", "Next action", "Next action date"]].copy()
st.dataframe(
    tbl, use_container_width=True, hide_index=True,
    column_config={
        "Amount": st.column_config.NumberColumn("Amount (R)", format="%,.0f"),
        "% complete": st.column_config.ProgressColumn("% complete", min_value=0, max_value=1,
                                                      format="%.0f%%"),
        "Next action": st.column_config.TextColumn(width="large"),
    })

# ------------------------------------------------------------- exceptions ----
st.markdown('<div class="sf-sec">Exceptions — overdue checklist items</div>', unsafe_allow_html=True)
ex = chk[chk["_overdue"] & chk["Deal ID"].isin(view["Deal ID"])].copy()
if ex.empty:
    st.success("No overdue items across the deals in view.")
else:
    ex = ex.sort_values("_days_late", ascending=False)
    cols = [c for c in ["Deal ID", "Stage", "Ref", "Task", "Responsible", "Status",
                        "Target date", "_days_late", "Notes"] if c in ex.columns]
    show = ex[cols].rename(columns={"_days_late": "Days late"})
    st.dataframe(show, use_container_width=True, hide_index=True,
                 column_config={"Task": st.column_config.TextColumn(width="large")})

# ------------------------------------------------------------- drill-down ----
st.markdown('<div class="sf-sec">Deal drill-down</div>', unsafe_allow_html=True)
pick = st.selectbox("Deal", view["Deal ID"].tolist())
d = view[view["Deal ID"] == pick].iloc[0]
m = st.columns(5)
m[0].metric("Client", str(d["Client"])[:22] or "—")
m[1].metric("Amount", rands(d["Amount"]))
m[2].metric("Stage", str(d["Stage"]) or "—")
m[3].metric("Complete", f"{d['% complete']:.0%}", f"{int(d['Outstanding'])} open")
m[4].metric("Overdue", int(d["Overdue"]), d["SLA"] if d["SLA"] != "—" else None)

sub = chk[chk["Deal ID"] == pick]
if "Stage" in sub.columns:
    per_stage = (sub[sub["_in_scope"]].groupby("Stage")
                 .agg(Items=("Status", "size"), Done=("_done", "sum"), Overdue=("_overdue", "sum"))
                 .reset_index())
    per_stage["% complete"] = per_stage["Done"] / per_stage["Items"]
    per_stage = per_stage.sort_values("Stage", key=lambda s: s.map(lambda x: stage_no(x) or 99))
    st.dataframe(per_stage, use_container_width=True, hide_index=True,
                 column_config={"% complete": st.column_config.ProgressColumn(
                     "% complete", min_value=0, max_value=1, format="%.0f%%")})

with st.expander(f"All checklist items for {pick}"):
    cols = [c for c in ["Stage", "Ref", "Task", "Responsible", "Status", "Target date",
                        "Completed date", "Evidence / document location", "Notes"] if c in sub.columns]
    st.dataframe(sub[cols], use_container_width=True, hide_index=True,
                 column_config={"Task": st.column_config.TextColumn(width="large")})


# -------------------------------------------------------------- reporting ----
st.markdown('<div class="sf-sec">Branded report</div>', unsafe_allow_html=True)
scope_note = f"{len(view)} deal(s) in scope"
html = build_report(view, chk[chk["Deal ID"].isin(view["Deal ID"])], firm, author, asof, scope_note)

r1, r2, r3 = st.columns(3)
r1.download_button("⬇ Download branded report (HTML)", html,
                   file_name=f"ScaleForce_Deal_Report_{asof:%Y%m%d}.html",
                   mime="text/html", use_container_width=True)
r2.download_button("⬇ Deal register (CSV)", tbl.to_csv(index=False),
                   file_name=f"deal_register_{asof:%Y%m%d}.csv",
                   mime="text/csv", use_container_width=True)
ex_all = chk[chk["_overdue"] & chk["Deal ID"].isin(view["Deal ID"])]
r3.download_button("⬇ Overdue items (CSV)", ex_all.drop(columns=["_in_scope", "_done", "_overdue"],
                                                        errors="ignore").to_csv(index=False),
                   file_name=f"overdue_items_{asof:%Y%m%d}.csv",
                   mime="text/csv", use_container_width=True,
                   disabled=ex_all.empty)

with st.expander("Preview the report"):
    components.html(html, height=680, scrolling=True)
st.caption("The report is a single self-contained HTML file. Open it and print to PDF "
           "(A4 landscape) for circulation to the AM or investment committee.")

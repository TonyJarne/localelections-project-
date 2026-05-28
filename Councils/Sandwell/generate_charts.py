"""
Generates 5 standalone D3.js HTML charts for the Sandwell 2026 election analysis.
Run: python generate_charts.py
Output: charts/ folder with 5 HTML files (D3 embedded, no internet needed).
"""
import json, os

OUT_DIR  = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project\Sandwell\charts"
D3_PATH  = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project\Sandwell\d3.v7.min.js"
os.makedirs(OUT_DIR, exist_ok=True)

with open(D3_PATH, "r", encoding="utf-8") as f:
    D3_JS = f.read()

# ── DATA ─────────────────────────────────────────────────────────────────────

party_totals = [
    {"party": "Reform UK",    "votes": 82541, "pct": 36.27, "seats": 41},
    {"party": "Labour",       "votes": 78083, "pct": 34.31, "seats": 28},
    {"party": "Green",        "votes": 36968, "pct": 16.25, "seats":  2},
    {"party": "Conservative", "votes": 24071, "pct": 10.58, "seats":  0},
    {"party": "Lib Dem",      "votes":  3202, "pct":  1.41, "seats":  0},
    {"party": "Independent",  "votes":  2273, "pct":  1.00, "seats":  1},
    {"party": "Other",        "votes":   423, "pct":  0.19, "seats":  0},
]

bloc_totals = [
    {"bloc": "Left (Labour + Green)", "votes": 115051, "pct": 50.56, "color": "#C0392B"},
    {"bloc": "Right (Reform + Con.)", "votes": 106612, "pct": 46.85, "color": "#1A8CB0"},
    {"bloc": "Liberal Democrats",     "votes":   3202, "pct":  1.41, "color": "#F39C12"},
    {"bloc": "Other / Independent",   "votes":   2696, "pct":  1.18, "color": "#95A5A6"},
]

PARTY_COLOURS = {
    "Reform UK":    "#12B6CF",
    "Labour":       "#E4003B",
    "Green":        "#02A95B",
    "Conservative": "#0087DC",
    "Lib Dem":      "#FAA61A",
    "Independent":  "#888888",
    "Other":        "#AAAAAA",
}

ward_data = [
    {"ward":"Bearwood",                          "reform":2222,"labour":4313,"green":4661,"con":805, "libdem":0,  "ind":0,   "other":58,  "reform_seats":0,"labour_seats":2,"green_seats":1,"split":False},
    {"ward":"Blackheath",                        "reform":5453,"labour":2851,"green":1461,"con":1202,"libdem":0,  "ind":100, "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Bristnall",                         "reform":3562,"labour":3453,"green":1824,"con":1221,"libdem":0,  "ind":0,   "other":0,   "reform_seats":2,"labour_seats":1,"green_seats":0,"split":True},
    {"ward":"Charlemont & Grove Vale",           "reform":4429,"labour":2812,"green":1372,"con":2313,"libdem":182,"ind":0,   "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Cradley Heath & Old Hill",          "reform":4304,"labour":2663,"green":1388,"con":1013,"libdem":230,"ind":0,   "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Friar Park & Stone Cross",          "reform":4843,"labour":2191,"green": 836,"con":1025,"libdem":0,  "ind":0,   "other":38,  "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Great Barr, Tamebridge & Yew Tree", "reform":3753,"labour":3950,"green":1363,"con":1487,"libdem":617,"ind":0,   "other":0,   "reform_seats":1,"labour_seats":2,"green_seats":0,"split":True},
    {"ward":"Great Bridge",                      "reform":4380,"labour":2247,"green":2067,"con": 739,"libdem":0,  "ind":0,   "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Greets Green & Lyng",               "reform":2338,"labour":4312,"green":1424,"con": 677,"libdem":215,"ind":0,   "other":0,   "reform_seats":0,"labour_seats":3,"green_seats":0,"split":False},
    {"ward":"Hateley Heath",                     "reform":3277,"labour":3505,"green": 873,"con": 745,"libdem":80, "ind":0,   "other":0,   "reform_seats":1,"labour_seats":2,"green_seats":0,"split":True},
    {"ward":"Hill Top",                          "reform":3757,"labour":2964,"green":1000,"con":1031,"libdem":191,"ind":0,   "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":True},
    {"ward":"Langley",                           "reform":4078,"labour":2875,"green":1697,"con":1130,"libdem":205,"ind":0,   "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":True},
    {"ward":"Newton & Valley",                   "reform":3155,"labour":3400,"green":1600,"con":1110,"libdem":556,"ind":0,   "other":0,   "reform_seats":1,"labour_seats":2,"green_seats":0,"split":True},
    {"ward":"Old Warley",                        "reform":3956,"labour":3530,"green":1623,"con":1145,"libdem":258,"ind":0,   "other":0,   "reform_seats":2,"labour_seats":1,"green_seats":0,"split":True},
    {"ward":"Oldbury",                           "reform":2242,"labour":3371,"green":1341,"con": 955,"libdem":0,  "ind":0,   "other":0,   "reform_seats":0,"labour_seats":3,"green_seats":0,"split":False},
    {"ward":"Princes End",                       "reform":4435,"labour":1699,"green": 687,"con":1111,"libdem":126,"ind":0,   "other":113, "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Rowley",                            "reform":4695,"labour":2668,"green":1090,"con":1068,"libdem":0,  "ind":0,   "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"Smethwick",                         "reform":1778,"labour":3819,"green":1908,"con": 616,"libdem":0,  "ind":0,   "other":0,   "reform_seats":0,"labour_seats":3,"green_seats":0,"split":False},
    {"ward":"Soho & Victoria",                   "reform": 734,"labour":3811,"green":1870,"con": 661,"libdem":0,  "ind":0,   "other":0,   "reform_seats":0,"labour_seats":3,"green_seats":0,"split":False},
    {"ward":"St Paul's",                         "reform": 769,"labour":5015,"green":1625,"con": 539,"libdem":0,  "ind":474, "other":214, "reform_seats":0,"labour_seats":3,"green_seats":0,"split":False},
    {"ward":"Tipton Green",                      "reform":3687,"labour":3001,"green":1444,"con": 801,"libdem":0,  "ind":1443,"other":0,   "reform_seats":2,"labour_seats":0,"green_seats":0,"split":True},
    {"ward":"Tividale",                          "reform":3913,"labour":3385,"green":1116,"con": 764,"libdem":148,"ind":0,   "other":0,   "reform_seats":2,"labour_seats":1,"green_seats":0,"split":True},
    {"ward":"Wednesbury",                        "reform":5150,"labour":2459,"green":1414,"con": 870,"libdem":394,"ind":106, "other":0,   "reform_seats":3,"labour_seats":0,"green_seats":0,"split":False},
    {"ward":"West Bromwich Central",             "reform":1631,"labour":3789,"green":1284,"con":1043,"libdem":0,  "ind":150, "other":0,   "reform_seats":0,"labour_seats":3,"green_seats":0,"split":False},
]

# ── HTML WRAPPER ─────────────────────────────────────────────────────────────

def html_wrap(title, subtitle, body_js, svg_width=820, svg_height=500, extra_css=""):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f5f6fa; padding: 20px; }}
h1 {{ text-align: center; padding: 10px 20px 4px; font-size: 1.3rem; color: #1F3864; }}
p.sub {{ text-align: center; color: #555; font-size: 0.88rem; padding-bottom: 16px; }}
#chart {{ display: flex; justify-content: center; overflow-x: auto; }}
.tooltip {{
  position: fixed; background: rgba(20,20,20,0.92); color: #fff;
  padding: 8px 12px; border-radius: 6px; font-size: 0.82rem;
  pointer-events: none; opacity: 0; transition: opacity 0.12s;
  max-width: 240px; line-height: 1.6; z-index: 9999;
}}
{extra_css}
</style>
</head>
<body>
<h1>{title}</h1>
<p class="sub">{subtitle}</p>
<div id="chart"></div>
<div class="tooltip" id="tt"></div>
<script>{D3_JS}</script>
<script>
const tt = d3.select("#tt");
function showTip(event, html) {{
  tt.style("opacity", 1).html(html)
    .style("left", (event.clientX + 14) + "px")
    .style("top",  (event.clientY - 28) + "px");
}}
function moveTip(event) {{
  tt.style("left", (event.clientX + 14) + "px")
    .style("top",  (event.clientY - 28) + "px");
}}
function hideTip() {{ tt.style("opacity", 0); }}

{body_js}
</script>
</body>
</html>"""

def save(filename, html):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Saved: {filename}")

# ════════════════════════════════════════════════════════════════════════════
# CHART 1 — Donut: Left vs Right Blocs
# ════════════════════════════════════════════════════════════════════════════

js1 = f"""
const data = {json.dumps(bloc_totals)};
const W = 680, H = 420, R = 150, iR = 70;

const svg = d3.select("#chart").append("svg")
  .attr("width", W).attr("height", H);
const g = svg.append("g").attr("transform", `translate(${{W/2 - 60}},${{H/2}})`);

const pie  = d3.pie().value(d => d.votes).sort(null);
const arc  = d3.arc().innerRadius(iR).outerRadius(R);
const arcH = d3.arc().innerRadius(iR).outerRadius(R + 14);

g.selectAll("path").data(pie(data)).enter().append("path")
  .attr("d", arc)
  .attr("fill", d => d.data.color)
  .attr("stroke", "#fff").attr("stroke-width", 3)
  .style("cursor", "pointer")
  .on("mouseover", function(event, d) {{
    d3.select(this).transition().duration(120).attr("d", arcH);
    showTip(event, `<strong>${{d.data.bloc}}</strong><br/>
      ${{d.data.pct.toFixed(1)}}% of all votes<br/>
      ${{d.data.votes.toLocaleString()}} votes`);
  }})
  .on("mousemove", moveTip)
  .on("mouseout", function() {{
    d3.select(this).transition().duration(120).attr("d", arc);
    hideTip();
  }});

g.append("text").attr("text-anchor","middle").attr("dy","-0.4em")
  .style("font-size","0.85rem").style("fill","#666").text("Total votes");
g.append("text").attr("text-anchor","middle").attr("dy","1em")
  .style("font-size","1.1rem").style("font-weight","700").style("fill","#1F3864")
  .text("227,561");

// Percentage labels on slices
g.selectAll("text.pct").data(pie(data)).enter().append("text")
  .attr("class","pct")
  .attr("transform", d => `translate(${{arc.centroid(d)}})`)
  .attr("text-anchor","middle").attr("dy","0.35em")
  .style("font-size","0.9rem").style("font-weight","700").style("fill","#fff")
  .text(d => d.data.pct >= 2 ? d.data.pct.toFixed(1) + "%" : "");

// Legend
const leg = svg.append("g").attr("transform", `translate(${{W/2 + 110}}, ${{H/2 - 80}})`);
data.forEach((d, i) => {{
  const row = leg.append("g").attr("transform", `translate(0,${{i * 46}})`);
  row.append("rect").attr("width",16).attr("height",16).attr("rx",3).attr("fill", d.color);
  row.append("text").attr("x",24).attr("y",12)
    .style("font-size","0.8rem").style("fill","#333").text(d.bloc);
  row.append("text").attr("x",24).attr("y",28)
    .style("font-size","0.85rem").style("fill",d.color).style("font-weight","700")
    .text(`${{d.pct.toFixed(1)}}%`);
}});
"""

save("chart1_blocs_donut.html", html_wrap(
    "Left vs Right: Vote Shares — Sandwell 2026",
    "The Left bloc (Labour + Green) won more votes overall, yet Reform took most council seats",
    js1))

# ════════════════════════════════════════════════════════════════════════════
# CHART 2 — Grouped bar: Votes % vs Seats %
# ════════════════════════════════════════════════════════════════════════════

js2 = f"""
const data = {json.dumps(party_totals)};
const colours = {json.dumps(PARTY_COLOURS)};
const totalSeats = d3.sum(data, d => d.seats);

const margin = {{top:30, right:20, bottom:70, left:55}};
const W = 740 - margin.left - margin.right;
const H = 420 - margin.top - margin.bottom;

const svg = d3.select("#chart").append("svg")
  .attr("width", W + margin.left + margin.right)
  .attr("height", H + margin.top + margin.bottom)
  .append("g").attr("transform", `translate(${{margin.left}},${{margin.top}})`);

const parties = data.map(d => d.party);
const x0 = d3.scaleBand().domain(parties).range([0, W]).paddingInner(0.3).paddingOuter(0.1);
const x1 = d3.scaleBand().domain(["votes","seats"]).range([0, x0.bandwidth()]).padding(0.08);
const y  = d3.scaleLinear().domain([0, 60]).range([H, 0]);

// Gridlines
svg.append("g").selectAll("line").data(y.ticks(6)).enter().append("line")
  .attr("x1",0).attr("x2",W)
  .attr("y1", d => y(d)).attr("y2", d => y(d))
  .attr("stroke","#dde").attr("stroke-dasharray","4,3");

// Y axis
svg.append("g").call(d3.axisLeft(y).ticks(6).tickFormat(d => d + "%"))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.78rem");

// X axis
svg.append("g").attr("transform",`translate(0,${{H}})`)
  .call(d3.axisBottom(x0).tickSize(0))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.82rem");

data.forEach(d => {{
  const seatPct = d.seats / totalSeats * 100;
  const gx = svg.append("g").attr("transform",`translate(${{x0(d.party)}},0)`);

  // Vote % bar (solid)
  gx.append("rect")
    .attr("x", x1("votes")).attr("width", x1.bandwidth())
    .attr("y", y(d.pct)).attr("height", H - y(d.pct))
    .attr("fill", colours[d.party]).attr("opacity", 0.9)
    .style("cursor","pointer")
    .on("mouseover", ev => showTip(ev,
      `<strong>${{d.party}}</strong><br/>
       Votes: <b>${{d.pct.toFixed(1)}}%</b> (${{d.votes.toLocaleString()}})<br/>
       Seats: <b>${{seatPct.toFixed(1)}}%</b> (${{d.seats}} seats)`))
    .on("mousemove", moveTip).on("mouseout", hideTip);

  // Seat % bar (hatched / semi-transparent)
  gx.append("rect")
    .attr("x", x1("seats")).attr("width", x1.bandwidth())
    .attr("y", y(seatPct)).attr("height", H - y(seatPct))
    .attr("fill", colours[d.party]).attr("opacity", 0.35)
    .attr("stroke", colours[d.party]).attr("stroke-width", 2)
    .style("cursor","pointer")
    .on("mouseover", ev => showTip(ev,
      `<strong>${{d.party}}</strong><br/>
       Votes: <b>${{d.pct.toFixed(1)}}%</b> (${{d.votes.toLocaleString()}})<br/>
       Seats: <b>${{seatPct.toFixed(1)}}%</b> (${{d.seats}} seats)`))
    .on("mousemove", moveTip).on("mouseout", hideTip);
}});

// Legend
const leg = svg.append("g").attr("transform",`translate(${{W - 200}},0)`);
[["% of votes (solid)","#555",0.9],["% of seats (faded)","#555",0.35]].forEach(([lbl,,op],i)=>{{
  leg.append("rect").attr("x",0).attr("y",i*22).attr("width",14).attr("height",14)
    .attr("fill","#555").attr("opacity",op);
  leg.append("text").attr("x",20).attr("y",i*22+11).style("font-size","0.78rem").text(lbl);
}});
"""

save("chart2_votes_vs_seats.html", html_wrap(
    "Votes vs Seats: The Disconnect — Sandwell 2026",
    "Solid bar = share of votes cast · Faded bar = share of council seats won · Hover for details",
    js2))

# ════════════════════════════════════════════════════════════════════════════
# CHART 3 — Horizontal stacked bar: Party votes per ward
# ════════════════════════════════════════════════════════════════════════════

js3 = f"""
const raw = {json.dumps(ward_data)};
const colours = {{
  "Reform UK":"#12B6CF","Conservative":"#0087DC",
  "Labour":"#E4003B","Green":"#02A95B",
  "Lib Dem":"#FAA61A","Other/Ind":"#AAAAAA"
}};
// Order: Right bloc first (Reform UK, Conservative), then Left bloc (Labour, Green), then others
const keys = ["Reform UK","Conservative","Labour","Green","Lib Dem","Other/Ind"];

// Build absolute vote data, sorted by Reform % descending
const absData = raw.map(d => {{
  const tot = d.reform + d.labour + d.green + d.con + d.libdem + d.ind + d.other;
  return {{
    ward: d.ward,
    "Reform UK": d.reform, "Conservative": d.con,
    "Labour": d.labour,    "Green": d.green,
    "Lib Dem": d.libdem,   "Other/Ind": d.ind + d.other,
    _total: tot
  }};
}}).sort((a,b) => (b["Reform UK"]/b._total) - (a["Reform UK"]/a._total));

// Convert to % of ward total for each party
const data = absData.map(d => {{
  const row = {{ ward: d.ward, _total: d._total }};
  keys.forEach(k => row[k] = d[k] / d._total * 100);
  return row;
}});

const margin = {{top:22, right:160, bottom:46, left:215}};
const rowH = 26;
const W = 860 - margin.left - margin.right;
const H = data.length * rowH;

const svg = d3.select("#chart").append("svg")
  .attr("width", W + margin.left + margin.right)
  .attr("height", H + margin.top + margin.bottom)
  .append("g").attr("transform",`translate(${{margin.left}},${{margin.top}})`);

const x = d3.scaleLinear().domain([0, 100]).range([0, W]);
const y = d3.scaleBand().domain(data.map(d => d.ward)).range([0, H]).padding(0.18);

// Light gridlines at 25% and 75% (behind bars)
[25, 75].forEach(v => {{
  svg.append("line")
    .attr("x1",x(v)).attr("x2",x(v)).attr("y1",0).attr("y2",H)
    .attr("stroke","#dde").attr("stroke-width",1).attr("stroke-dasharray","3,3");
}});

// X axis
svg.append("g").attr("transform",`translate(0,${{H}})`)
  .call(d3.axisBottom(x).ticks(5).tickFormat(d => d + "%"))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.76rem");

// Y axis
svg.append("g").call(d3.axisLeft(y).tickSize(0))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.76rem").attr("dx","-4");

// Stacked bars (using percentage data)
const stack = d3.stack().keys(keys);
svg.selectAll(".serie").data(stack(data)).enter().append("g")
  .attr("fill", d => colours[d.key])
  .selectAll("rect").data(d => d).enter().append("rect")
  .attr("y",  d => y(d.data.ward))
  .attr("x",  d => x(d[0]))
  .attr("width",  d => Math.max(0, x(d[1]) - x(d[0])))
  .attr("height", y.bandwidth())
  .style("cursor","pointer")
  .on("mouseover", function(event, d) {{
    const key  = d3.select(this.parentNode).datum().key;
    const absD = absData.find(a => a.ward === d.data.ward);
    showTip(event,
      `<strong>${{d.data.ward}}</strong><br/>
       ${{key}}: <b>${{absD[key].toLocaleString()}} votes</b> (${{d.data[key].toFixed(1)}}%)`);
  }})
  .on("mousemove", moveTip).on("mouseout", hideTip);

// 50% line drawn AFTER bars so it sits on top
svg.append("line")
  .attr("x1",x(50)).attr("x2",x(50)).attr("y1",0).attr("y2",H)
  .attr("stroke","#222").attr("stroke-width",2).attr("stroke-dasharray","6,3");
svg.append("text")
  .attr("x",x(50)).attr("y",-5)
  .attr("text-anchor","middle")
  .style("font-size","0.72rem").style("fill","#222").style("font-weight","700")
  .text("50%");

// Legend
const leg = svg.append("g").attr("transform",`translate(${{W+14}},10)`);
keys.forEach((k,i) => {{
  leg.append("rect").attr("y",i*26).attr("width",14).attr("height",14).attr("rx",2).attr("fill",colours[k]);
  leg.append("text").attr("x",20).attr("y",i*26+11).style("font-size","0.76rem").text(k);
}});
// Bloc labels in legend
leg.append("text").attr("x",0).attr("y", keys.length*26+18)
  .style("font-size","0.68rem").style("fill","#888").text("Right bloc ↑  |  Left bloc ↓");
"""

save("chart3_ward_stacked.html", html_wrap(
    "Party Votes by Ward — Sandwell 2026",
    "% of votes per ward · Sorted by Reform UK share · Dashed line = 50% · Hover for vote counts",
    js3, svg_width=900, svg_height=700,
    extra_css="#chart { overflow-x: auto; }"))

# ════════════════════════════════════════════════════════════════════════════
# CHART 4 — Grouped bars: Left bloc vs Reform per ward
# ════════════════════════════════════════════════════════════════════════════

js4 = f"""
const raw = {json.dumps(ward_data)};
const data = raw.map(d => ({{
  ward: d.ward,
  left: d.labour + d.green,
  reform: d.reform,
  split: d.split,
  reform_seats: d.reform_seats
}})).sort((a,b) => b.reform - a.reform);

const margin = {{top:10, right:40, bottom:40, left:220}};
const rowH = 30;
const W = 800 - margin.left - margin.right;
const H = data.length * rowH;

const svg = d3.select("#chart").append("svg")
  .attr("width", W + margin.left + margin.right)
  .attr("height", H + margin.top + margin.bottom)
  .append("g").attr("transform",`translate(${{margin.left}},${{margin.top}})`);

const xMax = d3.max(data, d => Math.max(d.left, d.reform)) * 1.06;
const x  = d3.scaleLinear().domain([0, xMax]).range([0, W]);
const y  = d3.scaleBand().domain(data.map(d => d.ward)).range([0, H]).padding(0.22);
const y1 = d3.scaleBand().domain(["reform","left"]).range([0, y.bandwidth()]).padding(0.06);

// X axis
svg.append("g").attr("transform",`translate(0,${{H}})`)
  .call(d3.axisBottom(x).ticks(5).tickFormat(d => (d/1000).toFixed(0)+"k"))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.76rem");

// Y axis
svg.append("g").call(d3.axisLeft(y).tickSize(0))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.76rem").attr("dx","-4");

// Gridlines
svg.append("g").selectAll("line").data(x.ticks(5)).enter().append("line")
  .attr("x1", d=>x(d)).attr("x2",d=>x(d)).attr("y1",0).attr("y2",H)
  .attr("stroke","#dde").attr("stroke-dasharray","3,3");

data.forEach(d => {{
  const gy = y(d.ward);

  // Reform bar
  svg.append("rect")
    .attr("x",0).attr("y", gy + y1("reform"))
    .attr("width", x(d.reform)).attr("height", y1.bandwidth())
    .attr("fill","#12B6CF").attr("opacity",0.88)
    .style("cursor","pointer")
    .on("mouseover", ev => showTip(ev,
      `<strong>${{d.ward}}</strong><br/>
       Reform UK: <b>${{d.reform.toLocaleString()}}</b><br/>
       Left (Lab+Green): <b>${{d.left.toLocaleString()}}</b><br/>
       ${{d.split ? "⚠ Left outvoted Reform here" : "Reform had more votes"}}<br/>
       Reform seats won: <b>${{d.reform_seats}}</b>`))
    .on("mousemove", moveTip).on("mouseout", hideTip);

  // Left bar
  svg.append("rect")
    .attr("x",0).attr("y", gy + y1("left"))
    .attr("width", x(d.left)).attr("height", y1.bandwidth())
    .attr("fill","#C0392B").attr("opacity",0.82)
    .style("cursor","pointer")
    .on("mouseover", ev => showTip(ev,
      `<strong>${{d.ward}}</strong><br/>
       Reform UK: <b>${{d.reform.toLocaleString()}}</b><br/>
       Left (Lab+Green): <b>${{d.left.toLocaleString()}}</b><br/>
       ${{d.split ? "⚠ Left outvoted Reform here" : "Reform had more votes"}}<br/>
       Reform seats won: <b>${{d.reform_seats}}</b>`))
    .on("mousemove", moveTip).on("mouseout", hideTip);

  // Split vote warning
  if (d.split) {{
    const xEnd = Math.max(x(d.left), x(d.reform));
    svg.append("text")
      .attr("x", xEnd + 6).attr("y", gy + y.bandwidth()/2 + 4)
      .style("font-size","0.68rem").style("fill","#E67E22").style("font-weight","700")
      .text("⚠ split");
  }}
}});

// Legend
const leg = svg.append("g").attr("transform",`translate(10,${{H+26}})`);
[["Reform UK","#12B6CF"],["Left bloc (Labour + Green)","#C0392B"]].forEach(([lbl,col],i) => {{
  leg.append("rect").attr("x",i*230).attr("width",14).attr("height",14).attr("rx",2).attr("fill",col);
  leg.append("text").attr("x",i*230+20).attr("y",11).style("font-size","0.78rem").text(lbl);
}});
"""

save("chart4_left_vs_reform.html", html_wrap(
    "Reform UK vs Left Bloc (Labour + Green) — By Ward",
    "⚠ marks wards where Labour + Green combined outvoted Reform, yet Reform still won seats",
    js4, svg_width=860, svg_height=760))

# ════════════════════════════════════════════════════════════════════════════
# CHART 5 — Lollipop: the 9 split-vote wards
# ════════════════════════════════════════════════════════════════════════════

split = [d for d in ward_data if d["split"]]
for d in split:
    d["left"]    = d["labour"] + d["green"]
    d["gap"]     = d["left"] - d["reform"]
    d["gap_pct"] = round(d["gap"] / (d["left"] + d["reform"]) * 100, 1)
split.sort(key=lambda d: d["gap"], reverse=True)

js5 = f"""
const data = {json.dumps(split)};

const margin = {{top:30, right:60, bottom:50, left:230}};
const rowH = 60;
const W = 740 - margin.left - margin.right;
const H = data.length * rowH;

const svg = d3.select("#chart").append("svg")
  .attr("width", W + margin.left + margin.right)
  .attr("height", H + margin.top + margin.bottom)
  .append("g").attr("transform",`translate(${{margin.left}},${{margin.top}})`);

const allVals = data.flatMap(d => [d.reform, d.left]);
const xMin = d3.min(allVals) * 0.88;
const xMax = d3.max(allVals) * 1.08;
const x = d3.scaleLinear().domain([xMin, xMax]).range([0, W]);
const y = d3.scaleBand().domain(data.map(d => d.ward)).range([0, H]).padding(0.4);

// X axis
svg.append("g").attr("transform",`translate(0,${{H}})`)
  .call(d3.axisBottom(x).ticks(5).tickFormat(d => d.toLocaleString()))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.76rem");

// Y axis
svg.append("g").call(d3.axisLeft(y).tickSize(0))
  .call(g => g.select(".domain").remove())
  .selectAll("text").style("font-size","0.8rem").attr("dx","-4");

// Gridlines
svg.append("g").selectAll("line").data(x.ticks(5)).enter().append("line")
  .attr("x1",d=>x(d)).attr("x2",d=>x(d)).attr("y1",0).attr("y2",H)
  .attr("stroke","#dde").attr("stroke-dasharray","3,3");

data.forEach(d => {{
  const cy = y(d.ward) + y.bandwidth() / 2;

  // Connecting line
  svg.append("line")
    .attr("x1",x(d.reform)).attr("x2",x(d.left))
    .attr("y1",cy).attr("y2",cy)
    .attr("stroke","#ccc").attr("stroke-width",2).attr("stroke-dasharray","5,3");

  // Gap label above the line
  svg.append("text")
    .attr("x",(x(d.reform)+x(d.left))/2).attr("y",cy-16)
    .attr("text-anchor","middle")
    .style("font-size","0.72rem").style("fill","#E67E22").style("font-weight","700")
    .text(`Left surplus: +${{d.gap.toLocaleString()}} (${{d.gap_pct}}%)`);

  // Reform dot
  svg.append("circle").attr("cx",x(d.reform)).attr("cy",cy).attr("r",10)
    .attr("fill","#12B6CF").style("cursor","pointer")
    .on("mouseover", ev => showTip(ev,
      `<strong>${{d.ward}}</strong><br/>
       Reform UK: <b>${{d.reform.toLocaleString()}}</b><br/>
       Labour: <b>${{d.labour.toLocaleString()}}</b><br/>
       Green: <b>${{d.green.toLocaleString()}}</b><br/>
       Left combined: <b>${{d.left.toLocaleString()}}</b><br/>
       Left surplus: <b>+${{d.gap.toLocaleString()}}</b><br/>
       Reform seats won: <b>${{d.reform_seats}}</b>`))
    .on("mousemove", moveTip).on("mouseout", hideTip);

  // Left dot
  svg.append("circle").attr("cx",x(d.left)).attr("cy",cy).attr("r",10)
    .attr("fill","#C0392B").style("cursor","pointer")
    .on("mouseover", ev => showTip(ev,
      `<strong>${{d.ward}}</strong><br/>
       Reform UK: <b>${{d.reform.toLocaleString()}}</b><br/>
       Labour: <b>${{d.labour.toLocaleString()}}</b><br/>
       Green: <b>${{d.green.toLocaleString()}}</b><br/>
       Left combined: <b>${{d.left.toLocaleString()}}</b><br/>
       Left surplus: <b>+${{d.gap.toLocaleString()}}</b><br/>
       Reform seats won: <b>${{d.reform_seats}}</b>`))
    .on("mousemove", moveTip).on("mouseout", hideTip);

  // Seats label below Reform dot
  svg.append("text")
    .attr("x",x(d.reform)).attr("y",cy+22)
    .attr("text-anchor","middle")
    .style("font-size","0.68rem").style("fill","#0e8fa3")
    .text(`${{d.reform_seats}} Reform seat${{d.reform_seats!==1?"s":""}}`);
}});

// Legend
const leg = svg.append("g").attr("transform",`translate(0,${{H+30}})`);
[["Reform UK","#12B6CF"],["Left bloc combined","#C0392B"]].forEach(([lbl,col],i)=>{{
  leg.append("circle").attr("cx",i*210+10).attr("cy",10).attr("r",9).attr("fill",col);
  leg.append("text").attr("x",i*210+26).attr("y",15).style("font-size","0.78rem").text(lbl);
}});
"""

save("chart5_split_vote.html", html_wrap(
    "The Split Vote: 9 Wards Where Left Outvoted Reform — But Reform Still Won Seats",
    "In each ward, the red dot (left combined) is further right than the blue (Reform), yet Reform took seats. Hover for full detail.",
    js5, svg_width=800, svg_height=620))

print(f"\nAll 5 charts saved to: {OUT_DIR}")
print("Open any .html file directly in Chrome or Edge — no internet needed.")

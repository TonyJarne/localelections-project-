import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

ward_data = [
    {"ward":"Bearwood",                          "reform":2222,"labour":4313,"green":4661,"con":805, "libdem":0,  "ind":0,   "other":58},
    {"ward":"Blackheath",                        "reform":5453,"labour":2851,"green":1461,"con":1202,"libdem":0,  "ind":100, "other":0},
    {"ward":"Bristnall",                         "reform":3562,"labour":3453,"green":1824,"con":1221,"libdem":0,  "ind":0,   "other":0},
    {"ward":"Charlemont & Grove Vale",           "reform":4429,"labour":2812,"green":1372,"con":2313,"libdem":182,"ind":0,   "other":0},
    {"ward":"Cradley Heath & Old Hill",          "reform":4304,"labour":2663,"green":1388,"con":1013,"libdem":230,"ind":0,   "other":0},
    {"ward":"Friar Park & Stone Cross",          "reform":4843,"labour":2191,"green": 836,"con":1025,"libdem":0,  "ind":0,   "other":38},
    {"ward":"Great Barr, Tamebridge & Yew Tree", "reform":3753,"labour":3950,"green":1363,"con":1487,"libdem":617,"ind":0,   "other":0},
    {"ward":"Great Bridge",                      "reform":4380,"labour":2247,"green":2067,"con": 739,"libdem":0,  "ind":0,   "other":0},
    {"ward":"Greets Green & Lyng",               "reform":2338,"labour":4312,"green":1424,"con": 677,"libdem":215,"ind":0,   "other":0},
    {"ward":"Hateley Heath",                     "reform":3277,"labour":3505,"green": 873,"con": 745,"libdem":80, "ind":0,   "other":0},
    {"ward":"Hill Top",                          "reform":3757,"labour":2964,"green":1000,"con":1031,"libdem":191,"ind":0,   "other":0},
    {"ward":"Langley",                           "reform":4078,"labour":2875,"green":1697,"con":1130,"libdem":205,"ind":0,   "other":0},
    {"ward":"Newton & Valley",                   "reform":3155,"labour":3400,"green":1600,"con":1110,"libdem":556,"ind":0,   "other":0},
    {"ward":"Old Warley",                        "reform":3956,"labour":3530,"green":1623,"con":1145,"libdem":258,"ind":0,   "other":0},
    {"ward":"Oldbury",                           "reform":2242,"labour":3371,"green":1341,"con": 955,"libdem":0,  "ind":0,   "other":0},
    {"ward":"Princes End",                       "reform":4435,"labour":1699,"green": 687,"con":1111,"libdem":126,"ind":0,   "other":113},
    {"ward":"Rowley",                            "reform":4695,"labour":2668,"green":1090,"con":1068,"libdem":0,  "ind":0,   "other":0},
    {"ward":"Smethwick",                         "reform":1778,"labour":3819,"green":1908,"con": 616,"libdem":0,  "ind":0,   "other":0},
    {"ward":"Soho & Victoria",                   "reform": 734,"labour":3811,"green":1870,"con": 661,"libdem":0,  "ind":0,   "other":0},
    {"ward":"St Paul's",                         "reform": 769,"labour":5015,"green":1625,"con": 539,"libdem":0,  "ind":474, "other":214},
    {"ward":"Tipton Green",                      "reform":3687,"labour":3001,"green":1444,"con": 801,"libdem":0,  "ind":1443,"other":0},
    {"ward":"Tividale",                          "reform":3913,"labour":3385,"green":1116,"con": 764,"libdem":148,"ind":0,   "other":0},
    {"ward":"Wednesbury",                        "reform":5150,"labour":2459,"green":1414,"con": 870,"libdem":394,"ind":106, "other":0},
    {"ward":"West Bromwich Central",             "reform":1631,"labour":3789,"green":1284,"con":1043,"libdem":0,  "ind":150, "other":0},
]

keys   = ["Reform UK", "Conservative", "Labour", "Green", "Lib Dem", "Other/Ind"]
colors = {"Reform UK":"#12B6CF","Conservative":"#0087DC","Labour":"#E4003B",
          "Green":"#02A95B","Lib Dem":"#FAA61A","Other/Ind":"#AAAAAA"}

# Build % rows, sort by Reform % desc
rows = []
for d in ward_data:
    tot = d["reform"]+d["labour"]+d["green"]+d["con"]+d["libdem"]+d["ind"]+d["other"]
    rows.append({
        "ward": d["ward"],
        "Reform UK":     d["reform"]/tot*100,
        "Conservative":  d["con"]/tot*100,
        "Labour":        d["labour"]/tot*100,
        "Green":         d["green"]/tot*100,
        "Lib Dem":       d["libdem"]/tot*100,
        "Other/Ind":     (d["ind"]+d["other"])/tot*100,
    })
rows.sort(key=lambda r: r["Reform UK"], reverse=True)

wards = [r["ward"] for r in rows]
n = len(wards)

fig, ax = plt.subplots(figsize=(13, 11))
fig.patch.set_facecolor("#f5f6fa")
ax.set_facecolor("#f5f6fa")

lefts = np.zeros(n)
y_pos = np.arange(n)

for key in keys:
    vals = np.array([r[key] for r in rows])
    ax.barh(y_pos, vals, left=lefts, color=colors[key], height=0.72, label=key)
    lefts += vals

# 50% line on top of bars
ax.axvline(50, color="#222", linewidth=1.8, linestyle="--", zorder=5)
ax.text(50, n - 0.1, "50%", ha="center", va="bottom", fontsize=8,
        fontweight="bold", color="#222", zorder=6)

# Axes
ax.set_yticks(y_pos)
ax.set_yticklabels(wards, fontsize=8.5)
ax.set_xlim(0, 100)
ax.set_xlabel("% of votes cast", fontsize=9)
ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x)}%"))
ax.tick_params(axis="x", labelsize=8.5)

# Light gridlines
for v in [25, 75]:
    ax.axvline(v, color="#dde", linewidth=0.8, linestyle="--", zorder=0)

ax.spines[["top","right","left","bottom"]].set_visible(False)
ax.tick_params(left=False)

# Legend
patches = [mpatches.Patch(color=colors[k], label=k) for k in keys]
ax.legend(handles=patches, loc="lower right", fontsize=8.5,
          framealpha=0.9, ncol=2)

ax.set_title("Party Votes by Ward — Sandwell 2026",
             fontsize=13, fontweight="bold", color="#1F3864", pad=12)
ax.set_xlabel("% of votes cast · sorted by Reform UK share (highest at top)", fontsize=8.5)

plt.tight_layout()
out = r"C:\Users\antoj\OneDrive\Escritorio\sandwell_project\chart3_preview.png"
plt.savefig(out, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
print(f"Saved: {out}")

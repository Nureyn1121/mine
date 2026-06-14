const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");

const { FaLeaf, FaMapMarkedAlt, FaClipboardList, FaDatabase, FaCheckCircle,
        FaUsers, FaChartBar, FaFlask, FaCogs, FaSearch, FaFileAlt,
        FaShieldAlt, FaTimes, FaSeedling, FaMountain, FaCloudRain,
        FaRoad, FaLayerGroup, FaGlobe } = require("react-icons/fa");

const C = {
  dark:    "1B3A2D",
  mid:     "2E7D52",
  light:   "52B788",
  accent:  "B7E4C7",
  pale:    "D8F3DC",
  white:   "FFFFFF",
  offwhite:"F8FAF9",
  gray:    "64748B",
  darkgray:"334155",
  gold:    "D4A017",
  red:     "C0392B",
};

async function icon(IconComp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const cardShadow = () => ({ type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.10 });

function addSlideTitle(s, title, subtitle) {
  s.addShape("rect", { x: 0, y: 0, w: 10, h: 1.15, fill: { color: C.dark } });
  s.addText(title, { x: 0.4, y: 0.12, w: 9.2, h: 0.7, fontSize: 24, bold: true, color: C.white, fontFace: "Cambria", margin: 0 });
  if (subtitle) {
    s.addText(subtitle, { x: 0.4, y: 0.82, w: 9.2, h: 0.28, fontSize: 10.5, color: C.accent, fontFace: "Calibri", italic: true, margin: 0 });
  }
}

function card(s, x, y, w, h, fillColor) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: fillColor || C.white },
    shadow: cardShadow(),
    line: { color: "E2E8F0", width: 0.5 }
  });
}

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Agricultural Suitability Project Management";

  // ── SLIDE 1: TITLE ──────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape("ellipse", { x: 7.2, y: -1.5, w: 4.2, h: 4.2, fill: { color: C.mid, transparency: 62 } });
    s.addShape("ellipse", { x: 7.8, y: -0.6, w: 2.6, h: 2.6, fill: { color: C.light, transparency: 52 } });
    s.addShape("rect", { x: 0, y: 4.85, w: 10, h: 0.775, fill: { color: C.mid, transparency: 30 } });

    const leafIco = await icon(FaSeedling, "#B7E4C7", 256);
    s.addImage({ data: leafIco, x: 0.4, y: 0.9, w: 0.62, h: 0.62 });

    s.addText("PROJECT MANAGEMENT FRAMEWORK", {
      x: 1.12, y: 0.88, w: 7.5, h: 0.5, fontSize: 12, bold: true, color: C.accent,
      fontFace: "Calibri", charSpacing: 3, margin: 0
    });
    s.addText("Identifying Suitable Areas\nfor Agricultural Development", {
      x: 0.4, y: 1.35, w: 8.5, h: 1.9, fontSize: 34, bold: true, color: C.white, fontFace: "Cambria", margin: 0
    });
    s.addText("A GIS & Remote Sensing Project Implementation Plan", {
      x: 0.4, y: 3.25, w: 8, h: 0.42, fontSize: 13, color: C.pale, fontFace: "Calibri", italic: true, margin: 0
    });
    s.addShape("line", { x: 0.4, y: 3.75, w: 2.6, h: 0, line: { color: C.light, width: 1.5 } });
    s.addText([
      { text: "Ardhi University  |  Dept. of Geomatics & Land Management", options: { breakLine: true } },
      { text: "Submission: 15 June 2026", options: {} }
    ], { x: 0.4, y: 4.9, w: 9, h: 0.56, fontSize: 10.5, color: C.accent, fontFace: "Calibri", margin: 0 });

    s.addNotes("Title slide. Project: Agricultural Suitability Analysis. Institution: Ardhi University, Tanzania. Submission Date: 15 June 2026.");
  }

  // ── SLIDE 2: OUTLINE ────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Presentation Outline", "8 Tasks · Full Project Lifecycle from Initiation to Closure");

    const items = [
      { n:"01", label:"Project Initiation" },
      { n:"02", label:"Project Planning" },
      { n:"03", label:"Data Requirements & Acquisition" },
      { n:"04", label:"GIS & RS Analysis" },
      { n:"05", label:"Field Verification" },
      { n:"06", label:"Deliverables & Reporting" },
      { n:"07", label:"Monitoring & QA" },
      { n:"08", label:"Project Closure" },
    ];

    const cw=2.2, ch=1.55, gx=0.1, gy=0.18, sx=0.35, sy=1.3;
    items.forEach((item, i) => {
      const col = i%4, row = Math.floor(i/4);
      const x = sx + col*(cw+gx), y = sy + row*(ch+gy);
      card(s, x, y, cw, ch, C.white);
      s.addShape("ellipse", { x:x+0.12, y:y+0.12, w:0.5, h:0.5, fill:{ color:C.mid } });
      s.addText(item.n, { x:x+0.12, y:y+0.12, w:0.5, h:0.5, fontSize:12, bold:true, color:C.white, fontFace:"Cambria", align:"center", valign:"middle", margin:0 });
      s.addText(item.label, { x:x+0.1, y:y+0.72, w:cw-0.2, h:0.72, fontSize:11.5, bold:true, color:C.dark, fontFace:"Cambria", align:"center", valign:"top", margin:0 });
    });
    s.addNotes("Overview of all 8 tasks in the project management framework.");
  }

  // ── SLIDE 3: TASK 1 – INITIATION ────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 1: Project Initiation", "Client Requirements · Stakeholders · ToR · Kick-off");

    const usersIco = await icon(FaUsers, "#FFFFFF", 256);
    const clipIco  = await icon(FaClipboardList, "#FFFFFF", 256);
    const globeIco = await icon(FaGlobe, "#FFFFFF", 256);

    const panels = [
      { ico: globeIco,  title:"Client Requirements", color:C.dark,
        pts:["Structured needs assessment","Target crop: Maize (district-wide)","Formal requirements document signed","Scale: 1:50,000 suitability map"] },
      { ico: usersIco,  title:"Stakeholders & ToR",  color:C.mid,
        pts:["Regional govt, farmers, NGOs, NBS","ToR: scope, payment, IP rights","Signed before technical work","7 stakeholder groups identified"] },
      { ico: clipIco,   title:"Kick-off & Scope",    color:"1B7A3E",
        pts:["Formal kick-off meeting: Week 1","Objectives & constraints agreed","CRS: WGS84 UTM Zone 37S","Minutes circulated within 24 hrs"] },
    ];

    panels.forEach((p, i) => {
      const x = 0.3 + i*3.23;
      card(s, x, 1.3, 3.05, 4.1, C.white);
      s.addShape("ellipse", { x:x+1.08, y:1.42, w:0.88, h:0.88, fill:{ color:p.color } });
      s.addImage({ data:p.ico, x:x+1.24, y:1.56, w:0.55, h:0.55 });
      s.addText(p.title, { x:x+0.1, y:2.42, w:2.85, h:0.44, fontSize:12.5, bold:true, color:C.dark, fontFace:"Cambria", align:"center", margin:0 });
      s.addText(p.pts.map(t=>({ text:t, options:{ bullet:true, breakLine:true, paraSpaceAfter:5 } })), {
        x:x+0.18, y:2.95, w:2.7, h:2.35, fontSize:10.5, color:C.darkgray, fontFace:"Calibri", margin:0
      });
    });
    s.addNotes("Initiation phase sets the project foundation. Key outputs: signed requirements doc and ToR, and a formal kick-off meeting.");
  }

  // ── SLIDE 4: TASK 2 – PLANNING ──────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 2: Project Planning", "WBS · Resources · Budget · Risks · Communication");

    card(s, 0.3, 1.28, 4.6, 4.1, C.white);
    s.addText("Work Breakdown Structure", { x:0.45, y:1.36, w:4.3, h:0.38, fontSize:12.5, bold:true, color:C.dark, fontFace:"Cambria", margin:0 });
    const wbs=[
      "1.0  Project Management & Initiation",
      "2.0  Data Acquisition",
      "3.0  Data Processing & Classification",
      "4.0  Suitability Analysis (MCDA)",
      "5.0  Field Verification",
      "6.0  Reporting & Delivery",
      "7.0  Project Closure",
    ];
    s.addText(wbs.map(t=>({ text:t, options:{ bullet:{ type:"number" }, breakLine:true, paraSpaceAfter:6 } })), {
      x:0.45, y:1.78, w:4.3, h:3.4, fontSize:10.5, color:C.darkgray, fontFace:"Calibri", margin:0
    });

    const stats=[
      { label:"Team Size",   value:"7",       unit:"specialists" },
      { label:"Duration",    value:"12",      unit:"weeks" },
      { label:"Budget",      value:"TZS 45M", unit:"total estimate" },
      { label:"Contingency", value:"15%",     unit:"budget reserve" },
    ];
    stats.forEach((st, i) => {
      const y = 1.28 + i*1.03;
      card(s, 5.15, y, 4.55, 0.9, i%2===0 ? C.pale : C.white);
      s.addText(st.value, { x:5.25, y:y+0.08, w:1.6, h:0.72, fontSize:26, bold:true, color:C.mid, fontFace:"Cambria", align:"center", valign:"middle", margin:0 });
      s.addText(st.label, { x:6.95, y:y+0.1,  w:2.55, h:0.36, fontSize:12.5, bold:true, color:C.dark, fontFace:"Cambria", margin:0 });
      s.addText(st.unit,  { x:6.95, y:y+0.48, w:2.55, h:0.3,  fontSize:10.5, color:C.gray, fontFace:"Calibri", margin:0 });
    });
    s.addNotes("7-member team over 12 weeks with TZS 45M budget. WBS decomposes the project into 7 major work packages.");
  }

  // ── SLIDE 5: GANTT CHART ────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Project Timeline — Gantt Chart", "12-Week Implementation Schedule");

    const activities = [
      { label:"Project Initiation & ToR",          weeks:[1,2] },
      { label:"Data Acquisition",                  weeks:[3,4,5] },
      { label:"Data Preprocessing",                weeks:[4,5,6] },
      { label:"LULC Mapping",                      weeks:[5,6,7] },
      { label:"Slope / Soil / Rainfall Analysis",  weeks:[6,7,8] },
      { label:"Accessibility & MCDA",              weeks:[7,8,9] },
      { label:"Field Verification",                weeks:[9,10] },
      { label:"Reporting & Map Production",        weeks:[10,11] },
      { label:"Client Review & Final Submission",  weeks:[11,12] },
    ];

    const chartX=2.7, chartY=1.3, chartW=7.0, chartH=4.05;
    const rowH = chartH/activities.length;
    const colW = chartW/12;

    card(s, 0.25, 1.22, 9.5, 4.25, C.white);

    // week headers
    for(let w=1;w<=12;w++){
      const x=chartX+(w-1)*colW;
      s.addShape("rect",{ x, y:chartY-0.3, w:colW, h:0.28, fill:{ color:C.dark }, line:{ color:C.dark, width:0.3 } });
      s.addText(`W${w}`,{ x, y:chartY-0.3, w:colW, h:0.28, fontSize:7.5, bold:true, color:C.white, fontFace:"Calibri", align:"center", valign:"middle", margin:0 });
    }

    activities.forEach((act, i) => {
      const y = chartY + i*rowH;
      const bg = i%2===0 ? C.pale : C.white;
      s.addShape("rect",{ x:0.28, y, w:2.38, h:rowH, fill:{ color:bg }, line:{ color:"E2E8F0", width:0.3 } });
      s.addText(act.label,{ x:0.32, y:y+0.02, w:2.3, h:rowH-0.04, fontSize:8.5, color:C.darkgray, fontFace:"Calibri", valign:"middle", margin:0 });
      s.addShape("rect",{ x:chartX, y, w:chartW, h:rowH, fill:{ color:bg }, line:{ color:"E2E8F0", width:0.3 } });

      const barStart=chartX+(act.weeks[0]-1)*colW;
      const barEnd=chartX+act.weeks[act.weeks.length-1]*colW;
      s.addShape("roundRect",{
        x:barStart+0.04, y:y+rowH*0.18, w:barEnd-barStart-0.08, h:rowH*0.64,
        rectRadius:0.04, fill:{ color:C.light },
        shadow:{ type:"outer", color:"000000", blur:3, offset:1, angle:45, opacity:0.12 }
      });
    });
    s.addNotes("12-week schedule. Activities overlap to maximise efficiency. Data acquisition starts Week 3 while initiation finishes.");
  }

  // ── SLIDE 6: DATA REQUIREMENTS ──────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 3: Data Requirements & Acquisition", "Multi-Source Spatial Datasets via GEE & Open Data");

    const datasets=[
      ["Sentinel-2 MSI",       "ESA / GEE",           "LULC mapping & classification",    "GeoTIFF"],
      ["SRTM DEM (30m)",       "NASA / USGS",         "Slope, drainage, terrain",         "GeoTIFF"],
      ["SoilGrids 250m",       "ISRIC / GEE",         "pH, SOC, CEC, clay suitability",   "GeoTIFF"],
      ["WorldClim / CHIRPS",   "WorldClim / UCSB",    "Rainfall & temperature",           "GeoTIFF"],
      ["ESA WorldCover 10m",   "ESA / FAO",           "Existing land use context",        "GeoTIFF"],
      ["Road Network (OSM)",   "OpenStreetMap",       "Market & road accessibility",      "Shapefile"],
      ["Admin. Boundaries",    "GADM / NBS",          "Spatial extent & reporting",       "Shapefile"],
      ["Protected Areas",      "UNEP WDPA / TAWA",    "Constraint masking / exclusions",  "Shapefile"],
    ];

    const cols=[2.3, 2.0, 3.0, 1.1];
    const hdrY=1.28;
    const totalW=cols.reduce((a,b)=>a+b,0);
    s.addShape("rect",{ x:0.3, y:hdrY, w:totalW, h:0.38, fill:{ color:C.dark } });
    const hdrs=["Dataset","Source","Purpose","Format"];
    let cx=0.3;
    hdrs.forEach((h,i)=>{
      s.addText(h,{ x:cx+0.06, y:hdrY+0.04, w:cols[i]-0.08, h:0.3, fontSize:10, bold:true, color:C.white, fontFace:"Cambria", valign:"middle", margin:0 });
      cx+=cols[i];
    });
    datasets.forEach((row, i)=>{
      const rowY=hdrY+0.38+i*0.44;
      s.addShape("rect",{ x:0.3, y:rowY, w:totalW, h:0.44, fill:{ color:i%2===0 ? C.pale : C.white }, line:{ color:"D1D5DB", width:0.3 } });
      let cx2=0.3;
      row.forEach((v,j)=>{
        s.addText(v,{ x:cx2+0.06, y:rowY+0.04, w:cols[j]-0.1, h:0.36,
          fontSize:j===0?9.5:9, bold:j===0, color:j===0?C.dark:C.darkgray, fontFace:"Calibri", valign:"middle", margin:0 });
        cx2+=cols[j];
      });
    });
    s.addNotes("All datasets are freely accessible. GEE used for raster processing. GADM and OSM for vector boundaries and roads.");
  }

  // ── SLIDE 7: ANALYSIS WORKFLOW ───────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 4: GIS & Remote Sensing Analysis Workflow", "14-Step Processing Pipeline");

    const steps=[
      { label:"Data\nAcquisition",       color:C.dark },
      { label:"Pre-\nprocessing",        color:C.mid },
      { label:"LULC\nClassification",    color:C.mid },
      { label:"Slope\nAnalysis",         color:C.mid },
      { label:"Soil\nSuitability",       color:C.light },
      { label:"Rainfall\nSuitability",   color:C.light },
      { label:"Accessibility\nAnalysis", color:C.light },
      { label:"Standardise\n(1–4 Scale)",color:"2C7873" },
      { label:"AHP\nWeighting",          color:"2C7873" },
      { label:"Weighted\nOverlay",       color:C.dark },
      { label:"Constraint\nMasking",     color:"8B2020" },
      { label:"Suitability\nClassify",   color:"4A148C" },
      { label:"Field\nVerification",     color:"C07000" },
      { label:"Final Map\n& Report",     color:C.dark },
    ];

    const bw=0.63, bh=0.88, gx=0.04, sx=0.22, sy=1.35, perRow=7;
    steps.forEach((st, i)=>{
      const row=Math.floor(i/perRow), col=i%perRow;
      const x=sx+col*(bw+gx), y=sy+row*(bh+0.46);
      s.addShape("roundRect",{ x, y, w:bw, h:bh, rectRadius:0.06, fill:{ color:st.color },
        shadow:{ type:"outer", color:"000000", blur:4, offset:2, angle:45, opacity:0.15 } });
      s.addText(st.label,{ x, y:y+0.1, w:bw, h:bh-0.1, fontSize:7.5, color:C.white, fontFace:"Calibri", align:"center", valign:"middle", margin:0 });
      s.addShape("ellipse",{ x:x+bw/2-0.14, y:y-0.17, w:0.28, h:0.28, fill:{ color:C.white } });
      s.addText(String(i+1),{ x:x+bw/2-0.14, y:y-0.17, w:0.28, h:0.28, fontSize:7, bold:true, color:C.dark, fontFace:"Calibri", align:"center", valign:"middle", margin:0 });
      if(col < perRow-1 && i < steps.length-1 && (i+1)%perRow !== 0){
        s.addShape("line",{ x:x+bw, y:y+bh/2, w:gx, h:0, line:{ color:C.gray, width:1 } });
      }
    });

    const legend=[
      { color:C.dark,   label:"Input/Output" },
      { color:C.mid,    label:"Preprocessing" },
      { color:C.light,  label:"Criterion Analysis" },
      { color:"2C7873", label:"MCDA/Weights" },
      { color:"8B2020", label:"Constraint Masking" },
      { color:"4A148C", label:"Classification" },
      { color:"C07000", label:"Validation" },
    ];
    let lx=0.22;
    legend.forEach(l=>{
      s.addShape("roundRect",{ x:lx, y:5.1, w:0.22, h:0.22, rectRadius:0.04, fill:{ color:l.color } });
      s.addText(l.label,{ x:lx+0.26, y:5.08, w:1.1, h:0.24, fontSize:7.5, color:C.darkgray, fontFace:"Calibri", margin:0 });
      lx+=1.4;
    });
    s.addNotes("14-step workflow from data acquisition to final map. Step 10 (Weighted Overlay) and Step 11 (Constraint Masking) are the critical analytical steps.");
  }

  // ── SLIDE 8: MCDA / AHP ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "MCDA — AHP Criterion Weights", "Analytic Hierarchy Process · Consistency Ratio ≤ 0.10");

    s.addChart(pres.charts.PIE, [{
      name:"Weight",
      labels:["Soil 35%","Rainfall 25%","Slope 15%","LULC 10%","Accessibility 8%","Flood 4%","Temp 3%"],
      values:[35,25,15,10,8,4,3]
    }], {
      x:0.3, y:1.25, w:4.5, h:4.15,
      chartColors:["1B3A2D","2E7D52","52B788","B7E4C7","D4A017","4A148C","8B2020"],
      showLabel:true, showPercent:false, showLegend:false,
      dataLabelColor:"FFFFFF", dataLabelFontSize:9, dataLabelFontBold:true,
      chartArea:{ fill:{ color:C.offwhite } },
    });

    const tableData=[
      [
        { text:"Criterion",   options:{ fill:{ color:C.dark }, color:C.white, bold:true, fontSize:10, fontFace:"Cambria" } },
        { text:"Weight",      options:{ fill:{ color:C.dark }, color:C.white, bold:true, fontSize:10, fontFace:"Cambria", align:"center" } },
        { text:"Priority",    options:{ fill:{ color:C.dark }, color:C.white, bold:true, fontSize:10, fontFace:"Cambria", align:"center" } },
      ],
      ...([
        ["Soil (pH, SOC, CEC, Texture)","35%","Highest"],
        ["Rainfall Adequacy","25%","High"],
        ["Slope / Terrain","15%","Moderate"],
        ["Land Use Compatibility","10%","Moderate"],
        ["Market & Road Accessibility","8%","Low"],
        ["Flood Risk","4%","Low"],
        ["Temperature Suitability","3%","Lowest"],
      ].map((row,i)=>row.map((cell,j)=>({
        text:cell,
        options:{ fill:{ color:i%2===0?C.pale:C.white }, color:C.darkgray, fontSize:10, fontFace:"Calibri", align:j>0?"center":"left" }
      }))))
    ];
    s.addTable(tableData,{ x:5.1, y:1.28, w:4.65, h:4.1, colW:[2.7,0.8,1.15], border:{ pt:0.5, color:"D1D5DB" }, margin:4 });
    s.addNotes("AHP weights: Soil 35%, Rainfall 25%, Slope 15%. These reflect the most critical factors for maize cultivation. CR must be ≤ 0.10.");
  }

  // ── SLIDE 9: SUITABILITY CLASSES ────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Suitability Classification", "FAO Framework — S1 · S2 · S3 · N");

    const classes=[
      { code:"S1", name:"Highly Suitable",      score:"≥ 75",  slope:"0–8°",   color:"1B7A3E", tc:C.white,  desc:"Excellent soil, adequate rainfall, flat terrain. Priority expansion areas for immediate investment." },
      { code:"S2", name:"Moderately Suitable",  score:"60–74", slope:"8–15°",  color:"52B788", tc:C.dark,   desc:"Good soil and climate with minor limitations. Suitable with small-scale inputs." },
      { code:"S3", name:"Marginally Suitable",  score:"40–59", slope:"15–30°", color:"D4A017", tc:C.dark,   desc:"Significant limitations in one or more factors. Requires soil amendments and investment." },
      { code:"N",  name:"Not Suitable",         score:"< 40",  slope:"> 30°",  color:C.red,    tc:C.white,  desc:"Poor soil or rainfall, steep terrain, or protected areas. Exclude from agricultural development." },
    ];

    classes.forEach((cl, i) => {
      const y=1.28+i*1.02;
      s.addShape("roundRect",{ x:0.3, y, w:9.4, h:0.95, rectRadius:0.07, fill:{ color:cl.color },
        shadow:{ type:"outer", color:"000000", blur:4, offset:2, angle:45, opacity:0.13 } });
      s.addText(cl.code,{ x:0.38, y:y+0.1, w:0.9, h:0.75, fontSize:26, bold:true, color:cl.tc, fontFace:"Cambria", align:"center", valign:"middle", margin:0 });
      s.addText(cl.name,{ x:1.35, y:y+0.08, w:2.4, h:0.38, fontSize:13.5, bold:true, color:cl.tc, fontFace:"Cambria", margin:0 });
      s.addText(`Score: ${cl.score}  |  Slope: ${cl.slope}`,{ x:1.35, y:y+0.5, w:2.8, h:0.32, fontSize:10, color:cl.tc, fontFace:"Calibri", margin:0 });
      s.addShape("line",{ x:4.3, y:y+0.18, w:0, h:0.6, line:{ color:cl.tc, width:0.5, transparency:55 } });
      s.addText(cl.desc,{ x:4.45, y:y+0.15, w:5.1, h:0.65, fontSize:10.5, color:cl.tc, fontFace:"Calibri", valign:"middle", margin:0 });
    });
    s.addNotes("Four FAO suitability classes. S1 and S2 are target areas. S3 may be developed with interventions. N areas are excluded.");
  }

  // ── SLIDE 10: FIELD VERIFICATION ───────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 5: Field Verification & Validation", "Ground-Truthing · GPS · Confusion Matrix · Accuracy ≥ 80%");

    const searchIco = await icon(FaSearch, "#FFFFFF", 256);
    const mapIco    = await icon(FaMapMarkedAlt, "#FFFFFF", 256);
    const checkIco  = await icon(FaCheckCircle, "#FFFFFF", 256);

    const panels=[
      { ico:searchIco, title:"Sampling Strategy",   color:C.dark,
        pts:["Stratified random sampling","Min. 50 pts/suitability class","~200–250 total sample points","Pre-loaded to GPS devices"] },
      { ico:mapIco,    title:"Data Collection",     color:C.mid,
        pts:["Garmin GPSMAP 64s GPS units","Kobo Collect / ODK mobile forms","3+ geo-tagged photos per point","Land use, soil, crop observations"] },
      { ico:checkIco,  title:"Accuracy Assessment", color:"1B7A3E",
        pts:["Confusion matrix generation","Overall Accuracy ≥ 80%","Kappa Coefficient ≥ 0.75","Iterative re-classification if needed"] },
    ];

    panels.forEach((p, i) => {
      const x=0.3+i*3.23;
      card(s, x, 1.3, 3.05, 4.1, C.white);
      s.addShape("ellipse",{ x:x+1.08, y:1.42, w:0.88, h:0.88, fill:{ color:p.color } });
      s.addImage({ data:p.ico, x:x+1.24, y:1.56, w:0.55, h:0.55 });
      s.addText(p.title,{ x:x+0.1, y:2.42, w:2.85, h:0.44, fontSize:12.5, bold:true, color:C.dark, fontFace:"Cambria", align:"center", margin:0 });
      s.addText(p.pts.map(t=>({ text:t, options:{ bullet:true, breakLine:true, paraSpaceAfter:5 } })), {
        x:x+0.18, y:2.95, w:2.7, h:2.35, fontSize:10.5, color:C.darkgray, fontFace:"Calibri", margin:0
      });
    });
    s.addNotes("Field verification validates model accuracy. Minimum OA 80% required. Kobo Collect minimises data entry errors.");
  }

  // ── SLIDE 11: DELIVERABLES ──────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 6: Deliverables & Reporting", "Project Outputs & Client Communication Strategy");

    const deliverables=[
      { ico:await icon(FaMapMarkedAlt, "#FFFFFF"), title:"Agricultural\nSuitability Map",   detail:"A1 at 1:50,000 | PDF + GeoTIFF", color:C.dark },
      { ico:await icon(FaDatabase, "#FFFFFF"),     title:"Geodatabase",                     detail:"File GDB/GeoPackage | All layers", color:C.mid },
      { ico:await icon(FaFileAlt, "#FFFFFF"),      title:"Technical\nReport",               detail:"60–80 pp | APA 7th | PDF+DOCX", color:"1B7A3E" },
      { ico:await icon(FaLayerGroup, "#FFFFFF"),   title:"Metadata\nDocumentation",         detail:"ISO 19115 compliant | XML+PDF", color:"2C7873" },
      { ico:await icon(FaGlobe, "#FFFFFF"),        title:"Web GIS\nDashboard (Optional)",    detail:"ArcGIS Online / GEE App | URL", color:"4A148C" },
      { ico:await icon(FaChartBar, "#FFFFFF"),     title:"Presentation\nPack",              detail:"Executive summary | PPTX+PDF", color:"C07000" },
    ];

    const cw=2.95, ch=1.65, sx=0.32, sy=1.3;
    deliverables.forEach((d, i) => {
      const col=i%3, row=Math.floor(i/3);
      const x=sx+col*(cw+0.1), y=sy+row*(ch+0.12);
      card(s, x, y, cw, ch, C.white);
      s.addShape("ellipse",{ x:x+0.18, y:y+0.15, w:0.65, h:0.65, fill:{ color:d.color } });
      s.addImage({ data:d.ico, x:x+0.3, y:y+0.26, w:0.42, h:0.42 });
      s.addText(d.title,{ x:x+0.92, y:y+0.16, w:1.9, h:0.7, fontSize:11.5, bold:true, color:C.dark, fontFace:"Cambria", margin:0 });
      s.addText(d.detail,{ x:x+0.92, y:y+0.9, w:1.9, h:0.6, fontSize:9.5, color:C.gray, fontFace:"Calibri", margin:0 });
    });
    s.addNotes("Six deliverables. Suitability map and technical report are primary outputs. Web GIS dashboard is optional but highly recommended.");
  }

  // ── SLIDE 12: MONITORING & QA ───────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 7: Monitoring & Quality Assurance", "RAG Status · Peer Review · Client Meetings · Change Management");

    card(s, 0.3, 1.28, 4.55, 4.1, C.white);
    s.addText("Quality Control Pillars",{ x:0.45, y:1.36, w:4.25, h:0.36, fontSize:12.5, bold:true, color:C.dark, fontFace:"Cambria", margin:0 });
    const qc=[
      "Data quality check (CRS, extent, attributes)",
      "Accuracy threshold: OA ≥ 80%, κ ≥ 0.75",
      "Peer review by independent senior analyst",
      "Cartographic checklist for all maps",
      "Version control: GitHub for all scripts",
      "Cloud backup: Google Drive / SharePoint",
    ];
    s.addText(qc.map(t=>({ text:t, options:{ bullet:true, breakLine:true, paraSpaceAfter:7 } })), {
      x:0.45, y:1.78, w:4.25, h:3.4, fontSize:10.5, color:C.darkgray, fontFace:"Calibri", margin:0
    });

    s.addText("Progress Monitoring — RAG System",{ x:5.1, y:1.28, w:4.55, h:0.36, fontSize:12.5, bold:true, color:C.dark, fontFace:"Cambria", margin:0 });
    const rag=[
      { color:"1B7A3E", label:"GREEN", desc:"On track — no action required" },
      { color:"E07B00", label:"AMBER", desc:"Minor delay — mitigation underway" },
      { color:C.red,    label:"RED",   desc:"Significant delay — escalate to PM" },
    ];
    rag.forEach((r, i) => {
      const y=1.78+i*1.1;
      card(s, 5.1, y, 4.55, 0.95, C.white);
      s.addShape("ellipse",{ x:5.2, y:y+0.17, w:0.62, h:0.62, fill:{ color:r.color } });
      s.addText(r.label,{ x:5.92, y:y+0.08, w:1.2, h:0.38, fontSize:13.5, bold:true, color:r.color, fontFace:"Cambria", margin:0 });
      s.addText(r.desc,{ x:5.92, y:y+0.52, w:3.55, h:0.32, fontSize:10, color:C.darkgray, fontFace:"Calibri", margin:0 });
    });
    s.addText("Client Reviews: Week 3 · Week 8 · Week 12",{ x:5.1, y:5.12, w:4.55, h:0.24, fontSize:9.5, italic:true, color:C.gray, fontFace:"Calibri", margin:0 });
    s.addNotes("RAG system for progress tracking. Three scheduled client review meetings. All scope changes require a signed Change Request Form.");
  }

  // ── SLIDE 13: PROJECT CLOSURE ────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offwhite };
    addSlideTitle(s, "Task 8: Project Closure", "Handover · Archiving · Lessons Learned · Project Evaluation");

    const steps=[
      { n:"1", label:"Final Report\nSubmission",   detail:"Deliverables packaged. PDF + hard drive hand-delivery." },
      { n:"2", label:"Client\nAcceptance",         detail:"5-day review period. Acceptance Certificate signed." },
      { n:"3", label:"Data\nArchiving",            detail:"ISO 19115 metadata. Two copies: server + client HDD." },
      { n:"4", label:"Lessons\nLearned",           detail:"Team session within 1 week. Stored in knowledge system." },
      { n:"5", label:"Project\nEvaluation",        detail:"Quality, schedule, budget & client satisfaction assessed." },
    ];

    steps.forEach((st, i) => {
      const x=0.3+i*1.9;
      card(s, x, 1.28, 1.78, 4.1, C.white);
      s.addShape("ellipse",{ x:x+0.54, y:1.42, w:0.7, h:0.7, fill:{ color:C.mid } });
      s.addText(st.n,{ x:x+0.54, y:1.42, w:0.7, h:0.7, fontSize:22, bold:true, color:C.white, fontFace:"Cambria", align:"center", valign:"middle", margin:0 });
      s.addText(st.label,{ x:x+0.08, y:2.28, w:1.62, h:0.68, fontSize:11.5, bold:true, color:C.dark, fontFace:"Cambria", align:"center", margin:0 });
      s.addShape("line",{ x:x+0.35, y:3.05, w:1.0, h:0, line:{ color:C.accent, width:1 } });
      s.addText(st.detail,{ x:x+0.1, y:3.15, w:1.6, h:2.1, fontSize:9.5, color:C.darkgray, fontFace:"Calibri", margin:0 });
    });
    s.addNotes("Five closure steps. The signed Acceptance Certificate triggers the final payment. Lessons learned feed future projects.");
  }

  // ── SLIDE 14: CONCLUSION ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape("ellipse",{ x:7.2, y:2.6, w:3.5, h:3.5, fill:{ color:C.mid, transparency:62 } });
    s.addShape("ellipse",{ x:-0.8, y:-0.5, w:3, h:3, fill:{ color:C.light, transparency:70 } });

    const leafIco2 = await icon(FaSeedling, "#B7E4C7", 256);
    s.addImage({ data:leafIco2, x:0.42, y:0.72, w:0.62, h:0.62 });

    s.addText("KEY TAKEAWAYS", { x:1.14, y:0.72, w:7, h:0.45, fontSize:11.5, bold:true, color:C.accent, fontFace:"Calibri", charSpacing:3, margin:0 });
    s.addText("A Scientifically Sound,\nCost-Effective Agricultural\nSuitability Framework", {
      x:0.4, y:1.15, w:7.5, h:1.9, fontSize:30, bold:true, color:C.white, fontFace:"Cambria", margin:0
    });

    const tks=[
      "GEE-based satellite analysis eliminates expensive data purchases",
      "AHP/MCDA ensures transparent, defensible weighting of criteria",
      "FAO framework provides internationally recognised suitability classes",
      "Field verification (OA ≥ 80%) validates the model objectively",
      "ISO 19115 geodatabase ensures data longevity and future reuse",
    ];
    s.addText(tks.map(t=>({ text:t, options:{ bullet:true, breakLine:true, paraSpaceAfter:6 } })), {
      x:0.4, y:3.1, w:6.8, h:2.2, fontSize:11.5, color:C.pale, fontFace:"Calibri", margin:0
    });
    s.addShape("line",{ x:0.4, y:5.2, w:3, h:0, line:{ color:C.light, width:1 } });
    s.addText("Ardhi University  ·  GIS & Remote Sensing Dept.  ·  June 2026", {
      x:0.4, y:5.28, w:9, h:0.28, fontSize:10, color:C.accent, fontFace:"Calibri", italic:true, margin:0
    });
    s.addNotes("Closing slide. Emphasise scientific rigour and cost-effectiveness. Open for questions.");
  }

  // Output path - save to current directory
  const outputPath = path.join(process.cwd(), "Agricultural_Suitability_Presentation.pptx");
  await pres.writeFile({ fileName: outputPath });
  console.log(`✅ PPT created successfully!`);
  console.log(`📁 File saved to: ${outputPath}`);
}

build().catch(err => { console.error("Error:", err); process.exit(1); });

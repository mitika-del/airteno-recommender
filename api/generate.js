'use strict';

// ============================================================================
// AIRTENO SITE RECOMMENDER
// ============================================================================
// POST /api/generate  ← called by Google Apps Script on form submit
//
// Flow: form data → Claude (placement logic) → PptxGenJS (2-slide deck)
//       → Resend email to submitter, CC Mitika always, CC Nilesh if senior review
//
// GOOGLE FORM FIELD ORDER (v[0] = timestamp, v[1] = field 1, ...):
//   1  submitter_name
//   2  submitter_email
//   3  site_type           (Residential Apartment / Villa / Playschool / School / Other Commercial)
//   4  site_name
//   5  address
//   6  carpet_area         (number, sq ft)
//   7  floors              (1 / 2 / 3+)
//   8  num_rooms
//   9  room_size           (Under 200 sqft / 200-400 sqft / Over 400 sqft)
//  10  other_spaces        (checkboxes, comma-separated)
//  11  install_location
//  12  filter_access       (Easy - door available / Moderate - needs hatch / Difficult - needs assessment)
//  13  ac_type             (Central AC / Split units only / No AC)
//  14  power_available     (Yes - point available / No - needs extension / Needs to be checked)
//  15  construction        (RCC sealed / Older brick / Mix)
//  16  kitchen_present     (Yes / No)
//  17  pollution_proximity (Yes / No)  — main road OR construction site nearby
//  18  site_notes          (free text)
//  -- QA CHECK FIELDS --
//  19  rooms_covered       (checkboxes: Living room, Dining, Master bedroom, Bedroom 2, Bedroom 3, Kitchen, Study, Others)
//  20  layout_available    (Customer provided floor plan / Hand drawing made on-site / Not done yet)
//  21  ceiling_height      (ft, short answer)
//  22  airflow_path        (paragraph — how air bubble will travel through the space)
//  23  obstructions        (paragraph — furniture, walls, partitions blocking airflow)
//  24  doors_to_open       (short answer — which doors must stay open)
//  25  existing_purifiers  (short answer — count and make of room purifiers customer already has)
//  26  mount_position      (Roof / terrace mount / Ground / floor stand / Wall mount / False ceiling / To be assessed)
//  27  inlet_method        (Wall hole required / Glass replacement with vent / Existing vent available / To be assessed)
//  28  rain_protection     (Yes - needs weatherproofing / No - location is sheltered / Needs assessment)
//  29  construction_status (Yes - inside the house / Yes - nearby within 500m / No)
// ============================================================================

const Anthropic = require('@anthropic-ai/sdk');
const PptxGenJS = require('pptxgenjs');
const { Resend } = require('resend');
const fs = require('fs');
const path = require('path');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const resend = new Resend(process.env.RESEND_API_KEY);

// ── Layout analysis prompt ────────────────────────────────────────────────────

const LAYOUT_SYSTEM_PROMPT = `You are analyzing a floor plan image for an airTENO positive pressure air system installation.

Extract:
1. Total carpet area in square feet — use any scale markings shown; if none, estimate from standard room dimensions
2. Area of each room or space in square feet
3. Brief layout notes relevant to positive pressure airflow (open corridors, sealed rooms, central zones, multi-floor)

Return ONLY valid JSON, no markdown, no commentary:
{
  "total_carpet_area_sqft": 1200,
  "rooms": [
    {"name": "Living Room", "area_sqft": 300},
    {"name": "Master Bedroom", "area_sqft": 180}
  ],
  "layout_notes": "Open plan between living and dining. Two bedrooms accessed via single corridor."
}`;

// ── System prompt ─────────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `You are an airTENO technical installation expert. airTENO is a positive pressure fresh air system — one central unit pulls outdoor air, filters it through H14 HEPA, and pushes clean air outward through the building. No ducting required in every room. Positive pressure spreads clean air through gaps under doors and corridors.

Site types handled: Residential Apartment, Villa, Playschool, Other Commercial.

Generate three outputs:

1. MODEL_RECOMMENDATION
Always: airTENO Max. Determine unit count:

Apartment / Villa:
- Up to 1500 sq ft: 1 unit
- Over 1500 sq ft: assess layout. Open plan or single corridor: 1 unit, set senior_review_required true. Multiple sealed wings: 2 units, set senior_review_required true.

Playschool:
- Do NOT apply the 1500 sq ft rule. Playschools typically have a high-ceiling central open area — positive pressure distributes effectively through this volume.
- Single-floor with central open zone: 1 unit, likely sufficient regardless of total area. Set senior_review_required true only if total area exceeds 5000 sq ft.
- Multi-floor OR classrooms fully sealed from central area: set senior_review_required true, note that unit count needs on-site confirmation.

Other Commercial: apply apartment logic, set senior_review_required true.

2. IDEAL_CONDITIONS
Exactly 5 to 7 bullet points. Write as direct instructions. No generic filler. Use site-specific data (obstructions, doors_to_open, existing_purifiers, health_conditions) to make these specific to this site.

Residential:
- Use doors_to_open field to name exact doors
- Reference obstructions if anything was flagged
- If existing_purifiers has a value: note they can be switched off once airTENO is running
- Balcony and main door closure at night
- AC co-use guidance (keep AC on, airTENO and AC complement each other)
- If health_conditions mentions asthma, COPD, allergies, elderly, or children: add a specific note about priority zones or higher ACH for those occupants

Playschool / Commercial:
- Door management during operating hours (keep 3-4 inches ajar)
- Corridor management for air distribution
- AC co-use guidance
- If health_conditions mentions children: note that airTENO directly reduces the exposure risk for the stated age group

3. INSTALLATION_NOTES
Exactly 4 to 6 bullet points. Use site data to be specific:
- Mount position and exact proposed location (use mount_position field)
- Inlet method: how outdoor air enters (use inlet_method field)
- Filter access: confirmed accessible or needs ladder / special access (use filter_access field)
- Power supply status (use power_available field)
- Rain protection: if rain_protection is "Yes", flag weatherproofing requirement
- If pollution_proximity has a value OR construction_status is not "No construction": flag elevated PM2.5 / dust load, recommend quarterly filter check instead of standard schedule
- If ceiling_height is above 12 ft, note it as advantageous for positive pressure distribution

Hard rules:
- Never suggest per-room units unless site is severely compartmentalized with no connecting corridor.
- If filter_access contains "Difficult": include "SITE VISIT REQUIRED before installation can be confirmed" as the first installation note.
- If pollution_proximity has a value OR construction_status is not "No construction": always include filter load flag in installation_notes.
- If existing_purifiers has a value: always address them in ideal_conditions.
- If health_conditions has a value: always incorporate it into ideal_conditions — name the condition and tailor the guidance.

Return ONLY valid JSON, no markdown, no commentary:
{
  "model_recommendation": "airTENO Max",
  "unit_count": 1,
  "ideal_conditions": ["...", "..."],
  "installation_notes": ["...", "..."],
  "flags": ["...", "..."],
  "senior_review_required": false
}`;

// ── Layout analysis (vision) ─────────────────────────────────────────────────

async function analyzeLayout(imageBase64, mimeType) {
  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 1024,
    system: LAYOUT_SYSTEM_PROMPT,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'image',
          source: { type: 'base64', media_type: mimeType, data: imageBase64 }
        },
        { type: 'text', text: 'Analyze this floor plan. Extract total carpet area and area by room in square feet.' }
      ]
    }]
  });

  const text = response.content[0].text.trim()
    .replace(/^```(?:json)?\n?/, '')
    .replace(/\n?```$/, '');

  return JSON.parse(text);
}

// ── Claude call ───────────────────────────────────────────────────────────────

async function getRecommendation(d) {
  const prompt = `Site Assessment:

Date of Visit: ${d.date_of_visit || 'Not specified'}
Surveyor: ${d.surveyor_name || 'Not specified'}
Site Type: ${d.site_type}
Site / Client Name: ${d.site_name}
Address: ${d.address}
Total Carpet Area: ${d.carpet_area} sq ft
Number of Floors: ${d.floors}
Number of Rooms / Classrooms: ${d.num_rooms}
Average Room Size: ${d.room_size}
Ceiling Height: ${d.ceiling_height ? d.ceiling_height + ' ft' : 'Not measured'}
Existing AC Type: ${d.ac_type || 'Not specified'}
Other Spaces to Cover: ${d.other_spaces || 'None specified'}

--- Installation Assessment ---
Candidate Install Location: ${d.install_location || 'Not specified'}
Filter Access: ${d.filter_access || 'Not assessed'}
Air Inlet Method: ${d.inlet_method || 'Not assessed'}
Power at Install Location: ${d.power_available || 'Not specified'}
Near Main Road or Active Construction: ${d.pollution_proximity || 'None noted'}
Rain Protection Needed: ${d.rain_protection || 'Not assessed'}
Obstructions to Airflow: ${d.obstructions || 'None noted'}
Doors to Keep Open: ${d.doors_to_open || 'Not specified'}
Installation Mount Position: ${d.mount_position || 'Not assessed'}
Site Notes: ${d.site_notes || 'None'}

--- User Profiling ---
Existing Room Purifiers: ${d.existing_purifiers || 'None'}
Construction / Renovation Status: ${d.construction_status || 'Not specified'}
Health Conditions in Household: ${d.health_conditions || 'None reported'}`;

  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 1024,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: prompt }]
  });

  const text = response.content[0].text.trim()
    .replace(/^```(?:json)?\n?/, '')
    .replace(/\n?```$/, '');

  return JSON.parse(text);
}

// ── PPTX builder ──────────────────────────────────────────────────────────────

async function buildPptx(d, rec, layoutAnalysis) {
  const DARK   = '002210';
  const GREEN  = '16BD95';
  const GRAY   = '4A4A4A';
  const LGRAY  = 'E0E3E3';
  const WHITE  = 'FFFFFF';
  const RED    = 'CC3300';
  const AMBER  = 'B87800';
  const BGBOX  = 'F5F6F6';

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 10" x 7.5"

  const assetsDir = path.join(process.cwd(), 'assets');
  const logoB64 = fs.readFileSync(path.join(assetsDir, 'logo.png')).toString('base64');
  const unitB64 = fs.readFileSync(path.join(assetsDir, 'unit.png')).toString('base64');

  const dateStr = new Date().toLocaleDateString('en-IN', {
    day: 'numeric', month: 'long', year: 'numeric'
  });
  const footer = `Prepared by ${d.surveyor_name || 'airTENO'}  |  ${d.date_of_visit || dateStr}  |  Internal working document — subject to site confirmation`;

  // ── SLIDE 1: Recommendation ───────────────────────────────────────────────

  const s1 = pptx.addSlide();

  // Header bar
  s1.addShape('rect', { x: 0, y: 0, w: 10, h: 0.65, fill: { color: DARK }, line: { color: DARK } });
  s1.addImage({ data: `data:image/png;base64,${logoB64}`, x: 0.2, y: 0.07, w: 1.4, h: 0.5 });
  s1.addText('Technical Site Recommendation', {
    x: 5.5, y: 0.18, w: 4.3, h: 0.3,
    fontSize: 9, color: LGRAY, fontFace: 'Open Sans', align: 'right'
  });

  // ── Area Assessment (top left) ──

  const siteRows = [
    ['Site Type',     d.site_type],
    ['Client / Site', d.site_name],
    ['Address',       d.address],
    ['Carpet Area',   `${d.carpet_area} sq ft`],
    ['Floors',        d.floors],
    ['Rooms',         d.num_rooms],
    ['Ceiling Ht',    d.ceiling_height ? `${d.ceiling_height} ft` : '—'],
    ['Surveyor',      d.surveyor_name || '—'],
    ['Date',          d.date_of_visit || '—'],
  ];
  // Expand Area Assessment box height to fit extra rows
  s1.addShape('rect', {
    x: 0.2, y: 0.8, w: 4.5, h: 3.2,
    fill: { color: BGBOX }, line: { color: LGRAY, width: 0.75 }
  });
  s1.addText('Area Assessment', {
    x: 0.35, y: 0.9, w: 4.2, h: 0.28,
    fontSize: 9.5, bold: true, color: DARK, fontFace: 'Open Sans'
  });
  siteRows.forEach(([label, val], i) => {
    const y = 1.24 + i * 0.26;
    s1.addText(label + ':', { x: 0.35, y, w: 1.5, h: 0.24, fontSize: 7.5, bold: true, color: GRAY, fontFace: 'Open Sans' });
    s1.addText(String(val || '-'), { x: 1.9, y, w: 2.7, h: 0.24, fontSize: 7.5, color: DARK, fontFace: 'Open Sans' });
  });

  // ── Model Recommendation (top right) ──
  s1.addShape('rect', {
    x: 5.0, y: 0.8, w: 4.8, h: 3.2,
    fill: { color: DARK }, line: { color: DARK }
  });
  s1.addText('Model Recommendation', {
    x: 5.15, y: 0.92, w: 2.8, h: 0.28,
    fontSize: 8.5, color: GREEN, fontFace: 'Open Sans'
  });
  s1.addText(`- ${rec.model_recommendation}`, {
    x: 5.15, y: 1.22, w: 2.8, h: 0.55,
    fontSize: 20, bold: true, color: WHITE, fontFace: 'Open Sans'
  });
  s1.addText(`${rec.unit_count} ${rec.unit_count === 1 ? 'Unit' : 'Units'}`, {
    x: 5.15, y: 1.82, w: 2.5, h: 0.32,
    fontSize: 13, color: GREEN, fontFace: 'Open Sans'
  });
  if (rec.senior_review_required) {
    s1.addText('! Senior review required — Mitika + Nilesh', {
      x: 5.15, y: 2.2, w: 4.55, h: 0.26,
      fontSize: 7.5, bold: true, color: 'FFB800', fontFace: 'Open Sans'
    });
  }
  // Health context (if any) — shown in the recommendation box
  if (d.health_conditions && d.health_conditions.trim()) {
    s1.addText('Health Context', {
      x: 5.15, y: 2.6, w: 4.55, h: 0.22,
      fontSize: 7.5, bold: true, color: GREEN, fontFace: 'Open Sans'
    });
    s1.addText(d.health_conditions, {
      x: 5.15, y: 2.84, w: 4.55, h: 0.9,
      fontSize: 7.5, color: LGRAY, fontFace: 'Open Sans', wrap: true, valign: 'top'
    });
  }
  s1.addImage({ data: `data:image/png;base64,${unitB64}`, x: 7.55, y: 0.85, w: 2.1, h: 2.2 });

  // ── Divider ──
  s1.addShape('rect', { x: 0.2, y: 4.1, w: 9.6, h: 0.02, fill: { color: LGRAY }, line: { color: LGRAY } });

  // ── Ideal Conditions (bottom left) ──
  s1.addText('Ideal Conditions for Optimal Performance', {
    x: 0.2, y: 4.18, w: 4.7, h: 0.3,
    fontSize: 9.5, bold: true, color: DARK, fontFace: 'Open Sans'
  });
  s1.addShape('rect', { x: 0.2, y: 4.5, w: 2.8, h: 0.025, fill: { color: GREEN }, line: { color: GREEN } });

  const conditionsText = rec.ideal_conditions.map(c => `- ${c}`).join('\n');
  s1.addText(conditionsText, {
    x: 0.2, y: 4.55, w: 4.6, h: 2.55,
    fontSize: 8.5, color: GRAY, fontFace: 'Open Sans',
    valign: 'top', wrap: true, paraSpaceAfter: 5
  });

  // ── Installation Notes (bottom right) ──
  s1.addText('Installation Notes', {
    x: 5.15, y: 4.18, w: 4.65, h: 0.3,
    fontSize: 9.5, bold: true, color: DARK, fontFace: 'Open Sans'
  });
  s1.addShape('rect', { x: 5.15, y: 4.5, w: 2.0, h: 0.025, fill: { color: GREEN }, line: { color: GREEN } });

  const notesText = rec.installation_notes.map(n => `- ${n}`).join('\n');
  s1.addText(notesText, {
    x: 5.15, y: 4.55, w: 4.6, h: 2.0,
    fontSize: 8.5, color: GRAY, fontFace: 'Open Sans',
    valign: 'top', wrap: true, paraSpaceAfter: 5
  });

  // Flags
  if (rec.flags && rec.flags.length > 0) {
    const flagText = rec.flags.map(f => `>> ${f}`).join('\n');
    s1.addText(flagText, {
      x: 5.15, y: 6.6, w: 4.6, h: 0.5,
      fontSize: 7.5, color: RED, fontFace: 'Open Sans',
      valign: 'top', wrap: true
    });
  }

  // Footer
  s1.addText(footer, {
    x: 0.2, y: 7.22, w: 9.6, h: 0.22,
    fontSize: 6.5, color: 'AAAAAA', fontFace: 'Open Sans', align: 'center'
  });

  // ── SLIDE 2: Layout Analysis (conditional) ───────────────────────────────

  if (layoutAnalysis) {
    const sl = pptx.addSlide();

    sl.addShape('rect', { x: 0, y: 0, w: 10, h: 0.65, fill: { color: DARK }, line: { color: DARK } });
    sl.addImage({ data: `data:image/png;base64,${logoB64}`, x: 0.2, y: 0.07, w: 1.4, h: 0.5 });
    sl.addText('Layout Analysis', {
      x: 5.5, y: 0.18, w: 4.3, h: 0.3,
      fontSize: 9, color: LGRAY, fontFace: 'Open Sans', align: 'right'
    });

    sl.addText(`${d.site_name}  —  ${d.address}`, {
      x: 0.3, y: 0.78, w: 9.4, h: 0.3,
      fontSize: 9.5, bold: true, color: DARK, fontFace: 'Open Sans'
    });

    // Total carpet area highlight bar
    sl.addShape('rect', { x: 0.2, y: 1.15, w: 4.5, h: 0.62, fill: { color: DARK }, line: { color: DARK } });
    sl.addText('Total Carpet Area', {
      x: 0.35, y: 1.22, w: 2.0, h: 0.22,
      fontSize: 8, color: GREEN, fontFace: 'Open Sans'
    });
    sl.addText(`${Number(layoutAnalysis.total_carpet_area_sqft).toLocaleString('en-IN')} sq ft`, {
      x: 2.35, y: 1.18, w: 2.2, h: 0.38,
      fontSize: 20, bold: true, color: WHITE, fontFace: 'Open Sans'
    });

    // Room breakdown table
    sl.addText('Area by Room', {
      x: 0.2, y: 1.9, w: 4.5, h: 0.28,
      fontSize: 9.5, bold: true, color: DARK, fontFace: 'Open Sans'
    });
    sl.addShape('rect', { x: 0.2, y: 2.2, w: 2.0, h: 0.025, fill: { color: GREEN }, line: { color: GREEN } });

    // Table header row
    sl.addShape('rect', { x: 0.2, y: 2.28, w: 4.5, h: 0.28, fill: { color: DARK }, line: { color: DARK } });
    sl.addText('Room / Space', {
      x: 0.35, y: 2.31, w: 3.0, h: 0.22,
      fontSize: 7.5, bold: true, color: WHITE, fontFace: 'Open Sans'
    });
    sl.addText('Sq ft', {
      x: 3.85, y: 2.31, w: 0.7, h: 0.22,
      fontSize: 7.5, bold: true, color: WHITE, fontFace: 'Open Sans', align: 'right'
    });

    let rowY = 2.56;
    layoutAnalysis.rooms.forEach((room, i) => {
      const bg = i % 2 === 0 ? BGBOX : WHITE;
      sl.addShape('rect', { x: 0.2, y: rowY, w: 4.5, h: 0.27, fill: { color: bg }, line: { color: LGRAY, width: 0.5 } });
      sl.addText(room.name, {
        x: 0.35, y: rowY + 0.03, w: 3.0, h: 0.22,
        fontSize: 8, color: GRAY, fontFace: 'Open Sans'
      });
      sl.addText(String(room.area_sqft), {
        x: 3.85, y: rowY + 0.03, w: 0.7, h: 0.22,
        fontSize: 8, bold: true, color: DARK, fontFace: 'Open Sans', align: 'right'
      });
      rowY += 0.27;
    });

    // Total row
    sl.addShape('rect', { x: 0.2, y: rowY, w: 4.5, h: 0.3, fill: { color: GREEN }, line: { color: GREEN } });
    sl.addText('TOTAL', {
      x: 0.35, y: rowY + 0.04, w: 3.0, h: 0.22,
      fontSize: 8.5, bold: true, color: WHITE, fontFace: 'Open Sans'
    });
    sl.addText(String(layoutAnalysis.total_carpet_area_sqft), {
      x: 3.85, y: rowY + 0.04, w: 0.7, h: 0.22,
      fontSize: 8.5, bold: true, color: WHITE, fontFace: 'Open Sans', align: 'right'
    });

    // Layout notes
    if (layoutAnalysis.layout_notes) {
      sl.addText('Layout Notes', {
        x: 0.2, y: rowY + 0.4, w: 4.5, h: 0.25,
        fontSize: 8.5, bold: true, color: DARK, fontFace: 'Open Sans'
      });
      sl.addText(layoutAnalysis.layout_notes, {
        x: 0.2, y: rowY + 0.68, w: 4.5, h: 1.2,
        fontSize: 8, color: GRAY, fontFace: 'Open Sans', wrap: true, valign: 'top'
      });
    }

    // Right side: floor plan image or placeholder
    if (d.layout_image_b64) {
      const mime = d.layout_image_type || 'image/jpeg';
      sl.addImage({
        data: `data:${mime};base64,${d.layout_image_b64}`,
        x: 5.0, y: 1.15, w: 4.8, h: 5.85,
        sizing: { type: 'contain', w: 4.8, h: 5.85 }
      });
    } else {
      sl.addShape('rect', {
        x: 5.0, y: 1.15, w: 4.8, h: 5.85,
        fill: { color: 'F0F2F2' },
        line: { color: LGRAY, width: 0.75, dashType: 'dash' }
      });
      sl.addText('+', {
        x: 5.0, y: 3.3, w: 4.8, h: 0.6,
        fontSize: 28, color: 'CCCCCC', fontFace: 'Open Sans', align: 'center'
      });
      sl.addText('Floor Plan', {
        x: 5.0, y: 4.0, w: 4.8, h: 0.28,
        fontSize: 8, color: 'AAAAAA', fontFace: 'Open Sans', align: 'center'
      });
    }

    sl.addText(footer, {
      x: 0.2, y: 7.22, w: 9.6, h: 0.22,
      fontSize: 6.5, color: 'AAAAAA', fontFace: 'Open Sans', align: 'center'
    });
  }

  // ── SLIDE 3: Reference Images placeholder ─────────────────────────────────

  const s2 = pptx.addSlide();

  s2.addShape('rect', { x: 0, y: 0, w: 10, h: 0.65, fill: { color: DARK }, line: { color: DARK } });
  s2.addImage({ data: `data:image/png;base64,${logoB64}`, x: 0.2, y: 0.07, w: 1.4, h: 0.5 });
  s2.addText('Reference Images — Site Visit', {
    x: 4.8, y: 0.18, w: 5.0, h: 0.3,
    fontSize: 9, color: LGRAY, fontFace: 'Open Sans', align: 'right'
  });

  s2.addText(`${d.site_name} — ${d.address}`, {
    x: 0.3, y: 0.82, w: 9.4, h: 0.32,
    fontSize: 10, bold: true, color: DARK, fontFace: 'Open Sans'
  });

  const imgSlots = [
    { x: 0.2,  y: 1.25, label: 'Installation Area' },
    { x: 3.5,  y: 1.25, label: 'Vent / Inlet Position' },
    { x: 6.8,  y: 1.25, label: 'Floor Plan' },
    { x: 0.2,  y: 4.15, label: 'Site Photo' },
    { x: 3.5,  y: 4.15, label: 'Site Photo' },
    { x: 6.8,  y: 4.15, label: 'Site Photo' },
  ];

  imgSlots.forEach(({ x, y, label }) => {
    s2.addShape('rect', {
      x, y, w: 3.1, h: 2.75,
      fill: { color: 'F0F2F2' },
      line: { color: LGRAY, width: 0.75, dashType: 'dash' }
    });
    s2.addText('+', {
      x, y: y + 0.85, w: 3.1, h: 0.6,
      fontSize: 28, color: 'CCCCCC', fontFace: 'Open Sans', align: 'center'
    });
    s2.addText(label, {
      x, y: y + 1.55, w: 3.1, h: 0.3,
      fontSize: 7.5, color: 'AAAAAA', fontFace: 'Open Sans', align: 'center'
    });
  });

  s2.addText(footer, {
    x: 0.2, y: 7.22, w: 9.6, h: 0.22,
    fontSize: 6.5, color: 'AAAAAA', fontFace: 'Open Sans', align: 'center'
  });

  return await pptx.write({ outputType: 'nodebuffer' });
}

// ── Email sender ──────────────────────────────────────────────────────────────

async function sendEmail(d, rec, pptxBuffer) {
  const filename = `${d.site_name.replace(/[^a-zA-Z0-9 ]/g, '').trim()}_airTENO_Recommendation.pptx`;

  // Primary recipient is always Mitika. Surveyor and Nilesh (if senior review) are CCed.
  const to = process.env.MITIKA_EMAIL || 'mitika@afferent.in';
  const cc = [];
  if (d.surveyor_email) cc.push(d.surveyor_email);
  if (rec.senior_review_required && process.env.NILESH_EMAIL) cc.push(process.env.NILESH_EMAIL);

  const reviewNote = rec.senior_review_required
    ? '\n\nSENIOR REVIEW FLAGGED — confirm unit count with Nilesh before sharing with client.'
    : '';

  const layoutNote = rec.layout_analysis
    ? `\n\nLayout analysis complete — ${rec.layout_analysis.total_carpet_area_sqft} sq ft total, ${rec.layout_analysis.rooms.length} rooms identified. See Slide 2.`
    : '';

  await resend.emails.send({
    from: process.env.FROM_EMAIL || 'airTENO <onboarding@resend.dev>',
    to,
    cc: cc.filter(Boolean),
    subject: `Tech Recommendation — ${d.site_name}`,
    text: `Tech recommendation for ${d.site_name} is attached.\n\nSurveyor: ${d.surveyor_name}  |  Date: ${d.date_of_visit}\nModel: ${rec.model_recommendation}  |  Units: ${rec.unit_count}${layoutNote}${reviewNote}\n\nEdit per field reality before sharing with client. Last slide has image placeholders — add site photos before sending.\n\nMitika`,
    attachments: [{ filename, content: pptxBuffer.toString('base64') }]
  });
}

// ── Handler ───────────────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const d = req.body;
  const required = ['surveyor_name', 'site_type', 'site_name', 'carpet_area'];
  const missing = required.filter(f => !d[f]);
  if (missing.length) {
    return res.status(400).json({ error: `Missing fields: ${missing.join(', ')}` });
  }

  try {
    console.log(`[recommender] Starting: ${d.site_name} (${d.site_type}) — surveyor: ${d.surveyor_name}`);

    const [rec, layoutAnalysis] = await Promise.all([
      getRecommendation(d),
      d.layout_image_b64
        ? analyzeLayout(d.layout_image_b64, d.layout_image_type || 'image/jpeg')
            .catch(err => { console.warn('[recommender] Layout analysis failed:', err.message); return null; })
        : Promise.resolve(null)
    ]);

    console.log(`[recommender] Claude: ${rec.unit_count} unit(s), senior_review=${rec.senior_review_required}`);
    if (layoutAnalysis) console.log(`[recommender] Layout: ${layoutAnalysis.total_carpet_area_sqft} sq ft, ${layoutAnalysis.rooms.length} rooms`);

    const pptxBuffer = await buildPptx(d, rec, layoutAnalysis);
    console.log(`[recommender] PPTX: ${Math.round(pptxBuffer.length / 1024)}KB`);

    await sendEmail(d, rec, pptxBuffer);
    console.log(`[recommender] Email sent to ${d.submitter_email}`);

    return res.json({
      success: true,
      unit_count: rec.unit_count,
      senior_review_required: rec.senior_review_required,
      flags: rec.flags,
      layout_analysis: layoutAnalysis
    });
  } catch (err) {
    console.error('[recommender] Error:', err.message);
    return res.status(500).json({ error: err.message });
  }
};

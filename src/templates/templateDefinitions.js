/**
 * Layout Toolkit - Template Definitions
 * Compatibility: Word for Mac (Microsoft 365) JS API
 * Avoids: Word.BuiltInTableStyle, table.style assignments
 */

const thumbnails = {
  triFold: `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="116" height="76" rx="3" fill="#fff" stroke="#C8D0DA" stroke-width="1"/><rect x="6" y="6" width="34" height="10" rx="1" fill="#1B3A5C"/><rect x="6" y="20" width="34" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="25" width="34" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="30" width="28" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="38" width="34" height="20" rx="1" fill="#EBF5FB"/><line x1="42" y1="4" x2="42" y2="76" stroke="#E1E5EA" stroke-width="0.5" stroke-dasharray="2"/><rect x="46" y="6" width="34" height="10" rx="1" fill="#C0392B"/><rect x="46" y="20" width="34" height="3" rx="1" fill="#D5DCE4"/><rect x="46" y="25" width="34" height="3" rx="1" fill="#D5DCE4"/><rect x="46" y="30" width="34" height="3" rx="1" fill="#D5DCE4"/><line x1="82" y1="4" x2="82" y2="76" stroke="#E1E5EA" stroke-width="0.5" stroke-dasharray="2"/><rect x="86" y="6" width="28" height="28" rx="1" fill="#F0F3F7"/><rect x="86" y="38" width="28" height="10" rx="1" fill="#1B3A5C"/><rect x="86" y="52" width="28" height="3" rx="1" fill="#D5DCE4"/></svg>`,
  biFold: `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="116" height="76" rx="3" fill="#fff" stroke="#C8D0DA" stroke-width="1"/><rect x="6" y="6" width="52" height="14" rx="1" fill="#1B3A5C"/><rect x="8" y="8" width="30" height="5" rx="1" fill="rgba(255,255,255,0.3)"/><rect x="6" y="24" width="52" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="29" width="52" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="34" width="52" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="48" width="52" height="24" rx="2" fill="#EBF5FB"/><line x1="60" y1="4" x2="60" y2="76" stroke="#E1E5EA" stroke-width="0.5" stroke-dasharray="2"/><rect x="64" y="6" width="50" height="30" rx="2" fill="#F0F3F7"/><rect x="64" y="40" width="50" height="8" rx="1" fill="#C0392B"/><rect x="64" y="52" width="50" height="3" rx="1" fill="#D5DCE4"/><rect x="64" y="57" width="50" height="3" rx="1" fill="#D5DCE4"/></svg>`,
  newsletter: `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="116" height="76" rx="3" fill="#fff" stroke="#C8D0DA" stroke-width="1"/><rect x="6" y="4" width="108" height="16" rx="2" fill="#1B3A5C"/><rect x="10" y="7" width="50" height="6" rx="1" fill="rgba(255,255,255,0.3)"/><line x1="6" y1="22" x2="114" y2="22" stroke="#C0392B" stroke-width="1.5"/><rect x="6" y="26" width="52" height="5" rx="1" fill="#2C3E50"/><rect x="6" y="34" width="52" height="3" rx="1" fill="#D5DCE4"/><rect x="6" y="39" width="52" height="3" rx="1" fill="#D5DCE4"/><rect x="64" y="26" width="50" height="24" rx="1" fill="#EBF5FB"/><rect x="64" y="54" width="50" height="4" rx="1" fill="#2C3E50"/><rect x="64" y="61" width="50" height="3" rx="1" fill="#D5DCE4"/></svg>`,
  eventProgram: `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="116" height="76" rx="3" fill="#fff" stroke="#C8D0DA" stroke-width="1"/><rect x="20" y="5" width="80" height="12" rx="2" fill="#1B3A5C"/><rect x="30" y="8" width="60" height="5" rx="1" fill="rgba(255,255,255,0.3)"/><rect x="35" y="20" width="50" height="4" rx="1" fill="#C0392B"/><rect x="10" y="28" width="20" height="3" rx="1" fill="#8E44AD"/><rect x="34" y="28" width="76" height="3" rx="1" fill="#D5DCE4"/><rect x="10" y="34" width="20" height="3" rx="1" fill="#8E44AD"/><rect x="34" y="34" width="76" height="3" rx="1" fill="#D5DCE4"/><rect x="10" y="40" width="20" height="3" rx="1" fill="#8E44AD"/><rect x="34" y="40" width="76" height="3" rx="1" fill="#D5DCE4"/></svg>`,
  businessCard: `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg"><rect x="15" y="10" width="90" height="55" rx="4" fill="#fff" stroke="#C8D0DA" stroke-width="1"/><rect x="20" y="16" width="10" height="10" rx="2" fill="#1B3A5C"/><rect x="34" y="16" width="40" height="5" rx="1" fill="#1B3A5C"/><rect x="34" y="23" width="30" height="3" rx="1" fill="#C0392B"/><line x1="20" y1="33" x2="100" y2="33" stroke="#E1E5EA" stroke-width="0.5"/><rect x="20" y="38" width="50" height="3" rx="1" fill="#D5DCE4"/><rect x="20" y="43" width="45" height="3" rx="1" fill="#D5DCE4"/></svg>`,
  flyer: `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="116" height="76" rx="3" fill="#fff" stroke="#C8D0DA" stroke-width="1"/><rect x="4" y="4" width="112" height="30" rx="2" fill="#1B3A5C"/><rect x="20" y="10" width="80" height="8" rx="1" fill="rgba(255,255,255,0.3)"/><rect x="15" y="40" width="90" height="5" rx="1" fill="#C0392B"/><rect x="10" y="50" width="100" height="3" rx="1" fill="#D5DCE4"/><rect x="10" y="55" width="100" height="3" rx="1" fill="#D5DCE4"/><rect x="25" y="68" width="70" height="8" rx="4" fill="#C0392B"/></svg>`,
};

const TEMPLATES = [
  { id: "tri-fold-brochure", name: "Tri-Fold Brochure", description: "Classic three-panel layout for marketing and information materials", category: "brochure", thumbnail: thumbnails.triFold, build: buildTriFoldBrochure },
  { id: "bi-fold-brochure", name: "Bi-Fold Brochure", description: "Two-panel spread ideal for programs and short features", category: "brochure", thumbnail: thumbnails.biFold, build: buildBiFoldBrochure },
  { id: "newsletter-classic", name: "Classic Newsletter", description: "Two-column newsletter with masthead, articles, and sidebar", category: "newsletter", thumbnail: thumbnails.newsletter, build: buildNewsletter },
  { id: "event-program", name: "Event Program", description: "Structured program with schedule blocks and speaker notes", category: "program", thumbnail: thumbnails.eventProgram, build: buildEventProgram },
  { id: "business-card", name: "Business Card", description: "Standard 3.5 x 2 in. card layout with contact details", category: "card", thumbnail: thumbnails.businessCard, build: buildBusinessCard },
  { id: "event-flyer", name: "Event Flyer", description: "Bold single-page flyer with headline, details, and call to action", category: "brochure", thumbnail: thumbnails.flyer, build: buildFlyer },
];

// Helper: create a borderless layout table (Mac-compatible)
function makeLayoutTable(body, rows, cols, data) {
  const table = body.insertTable(rows, cols, Word.InsertLocation.end, data);
  var sides = [
    Word.BorderLocation.top,
    Word.BorderLocation.bottom,
    Word.BorderLocation.left,
    Word.BorderLocation.right,
    Word.BorderLocation.insideHorizontal,
    Word.BorderLocation.insideVertical,
  ];
  sides.forEach(function(side) {
    try {
      var border = table.getBorder(side);
      border.type = Word.BorderType.none;
    } catch (e) { /* safe to ignore */ }
  });
  return table;
}

async function buildTriFoldBrochure(context, brand) {
  var body = context.document.body;
  body.clear();
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  var section = sections.items[0];
  section.pageSetup.orientation = Word.PageOrientation.landscape;
  section.pageSetup.topMargin = 36;
  section.pageSetup.bottomMargin = 36;
  section.pageSetup.leftMargin = 36;
  section.pageSetup.rightMargin = 36;
  await context.sync();

  var title = body.insertParagraph(brand.name || "Organization Name", Word.InsertLocation.end);
  title.font.set({ name: brand.headingFont, size: 18, bold: true, color: brand.primary });
  title.spaceAfter = 2;

  var sub = body.insertParagraph("Your Tagline or Mission Statement Goes Here", Word.InsertLocation.end);
  sub.font.set({ name: brand.bodyFont, size: 10, italic: true, color: brand.accent });
  sub.spaceAfter = 8;
  await context.sync();

  var table = makeLayoutTable(body, 1, 3, [["", "", ""]]);

  var c1 = table.getCell(0, 0);
  var c1t = c1.body.insertParagraph("Welcome", Word.InsertLocation.end);
  c1t.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.primary });
  c1t.spaceAfter = 6;
  var c1b = c1.body.insertParagraph("Introduce your organization, mission, and purpose here. This panel is the first thing readers see when they open the brochure. Keep the language inviting and concise.", Word.InsertLocation.end);
  c1b.font.set({ name: brand.bodyFont, size: 9, color: "#333333" });
  c1b.spaceAfter = 8;
  var c1i = c1.body.insertParagraph("[Insert image or graphic here]", Word.InsertLocation.end);
  c1i.font.set({ name: brand.bodyFont, size: 8, italic: true, color: "#999999" });

  var c2 = table.getCell(0, 1);
  var c2t = c2.body.insertParagraph("What We Do", Word.InsertLocation.end);
  c2t.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.accent });
  c2t.spaceAfter = 6;
  ["1. Program or Service One: Brief description of what you offer and who it serves.",
   "2. Program or Service Two: Explain the value and impact of this offering.",
   "3. Program or Service Three: Highlight outcomes and community benefit."
  ].forEach(function(feat) {
    var fp = c2.body.insertParagraph(feat, Word.InsertLocation.end);
    fp.font.set({ name: brand.bodyFont, size: 9, color: "#333333" });
    fp.spaceAfter = 5;
  });

  var c3 = table.getCell(0, 2);
  var c3t = c3.body.insertParagraph("Get Involved", Word.InsertLocation.end);
  c3t.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.primary });
  c3t.spaceAfter = 6;
  ["Phone: (555) 123-4567", "Email: info@example.org", "Web: www.example.org", "Address: 123 Main St, Suite 100"].forEach(function(line) {
    var cp = c3.body.insertParagraph(line, Word.InsertLocation.end);
    cp.font.set({ name: brand.bodyFont, size: 9, color: "#333333" });
    cp.spaceAfter = 3;
  });
  var c3cta = c3.body.insertParagraph("Follow us on social media!", Word.InsertLocation.end);
  c3cta.font.set({ name: brand.bodyFont, size: 9, bold: true, color: brand.accent });
  c3cta.spaceBefore = 8;

  await context.sync();
}

async function buildBiFoldBrochure(context, brand) {
  var body = context.document.body;
  body.clear();
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  var section = sections.items[0];
  section.pageSetup.orientation = Word.PageOrientation.landscape;
  section.pageSetup.topMargin = 54;
  section.pageSetup.bottomMargin = 54;
  section.pageSetup.leftMargin = 54;
  section.pageSetup.rightMargin = 54;
  await context.sync();

  var header = body.insertParagraph(brand.name || "Organization Name", Word.InsertLocation.end);
  header.font.set({ name: brand.headingFont, size: 20, bold: true, color: brand.primary });
  header.alignment = Word.Alignment.center;
  header.spaceAfter = 2;

  var tagline = body.insertParagraph("Subtitle or Event Name", Word.InsertLocation.end);
  tagline.font.set({ name: brand.bodyFont, size: 11, italic: true, color: brand.accent });
  tagline.alignment = Word.Alignment.center;
  tagline.spaceAfter = 12;
  await context.sync();

  var table = makeLayoutTable(body, 1, 2, [["", ""]]);

  var left = table.getCell(0, 0);
  var lt = left.body.insertParagraph("About This Publication", Word.InsertLocation.end);
  lt.font.set({ name: brand.headingFont, size: 14, bold: true, color: brand.primary });
  lt.spaceAfter = 8;
  var lb1 = left.body.insertParagraph("Use this space to introduce the purpose of this brochure. Describe your organization, an upcoming event, or a program you want to promote. A bi-fold is ideal for concise, impactful messaging with room for visuals on the right panel.", Word.InsertLocation.end);
  lb1.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });
  lb1.spaceAfter = 8;
  var lb2 = left.body.insertParagraph("Add supporting details here: dates, locations, eligibility criteria, or any additional context that helps the reader take the next step.", Word.InsertLocation.end);
  lb2.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });

  var right = table.getCell(0, 1);
  var ri = right.body.insertParagraph("[Insert featured image or infographic here]", Word.InsertLocation.end);
  ri.font.set({ name: brand.bodyFont, size: 9, italic: true, color: "#999999" });
  ri.alignment = Word.Alignment.center;
  ri.spaceAfter = 12;
  var rt = right.body.insertParagraph("Key Highlights", Word.InsertLocation.end);
  rt.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.accent });
  rt.spaceAfter = 6;
  ["\u2022 Highlight one: A brief, compelling point about your work or event.",
   "\u2022 Highlight two: Another key benefit or feature for your audience.",
   "\u2022 Highlight three: A call to action or invitation to participate."
  ].forEach(function(h) {
    var hp = right.body.insertParagraph(h, Word.InsertLocation.end);
    hp.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });
    hp.spaceAfter = 5;
  });

  await context.sync();
}

async function buildNewsletter(context, brand) {
  var body = context.document.body;
  body.clear();
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  var section = sections.items[0];
  section.pageSetup.orientation = Word.PageOrientation.portrait;
  section.pageSetup.topMargin = 54;
  section.pageSetup.bottomMargin = 54;
  section.pageSetup.leftMargin = 54;
  section.pageSetup.rightMargin = 54;
  await context.sync();

  var masthead = body.insertParagraph(brand.name || "The Organization Newsletter", Word.InsertLocation.end);
  masthead.font.set({ name: brand.headingFont, size: 24, bold: true, color: brand.primary });
  masthead.alignment = Word.Alignment.center;
  masthead.spaceAfter = 2;

  var dateline = body.insertParagraph("Volume 1, Issue 1  |  Spring 2026", Word.InsertLocation.end);
  dateline.font.set({ name: brand.bodyFont, size: 9, color: brand.accent });
  dateline.alignment = Word.Alignment.center;
  dateline.spaceAfter = 4;

  var div1 = body.insertParagraph("\u2500".repeat(70), Word.InsertLocation.end);
  div1.font.set({ color: brand.accent, size: 7 });
  div1.alignment = Word.Alignment.center;
  div1.spaceAfter = 10;

  var headline = body.insertParagraph("Lead Story Headline Goes Here", Word.InsertLocation.end);
  headline.font.set({ name: brand.headingFont, size: 16, bold: true, color: "#1A1A2E" });
  headline.spaceAfter = 4;

  var byline = body.insertParagraph("By Author Name  |  Date", Word.InsertLocation.end);
  byline.font.set({ name: brand.bodyFont, size: 9, italic: true, color: "#888888" });
  byline.spaceAfter = 8;

  var lead1 = body.insertParagraph("Write your lead article here. This should be the most important or timely piece of content in this issue. Aim for 150-300 words that capture the reader's attention and convey the essential information.", Word.InsertLocation.end);
  lead1.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });
  lead1.spaceAfter = 8;

  var lead2 = body.insertParagraph("Continue the article with additional details. Include quotes from stakeholders, data points that support the story, and any relevant background.", Word.InsertLocation.end);
  lead2.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });
  lead2.spaceAfter = 10;

  var div2 = body.insertParagraph("\u2500".repeat(70), Word.InsertLocation.end);
  div2.font.set({ color: "#E1E5EA", size: 5 });
  div2.alignment = Word.Alignment.center;
  div2.spaceAfter = 8;
  await context.sync();

  var table = makeLayoutTable(body, 1, 2, [["", ""]]);

  var lc = table.getCell(0, 0);
  var lct = lc.body.insertParagraph("Secondary Story", Word.InsertLocation.end);
  lct.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.primary });
  lct.spaceAfter = 5;
  var lcb = lc.body.insertParagraph("A shorter feature or update. This could cover a recent event, an upcoming initiative, a profile of a community member, or a policy development relevant to your audience.", Word.InsertLocation.end);
  lcb.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });

  var rc = table.getCell(0, 1);
  var rct = rc.body.insertParagraph("Announcements", Word.InsertLocation.end);
  rct.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.accent });
  rct.spaceAfter = 5;
  ["\u25A0 Upcoming Event: Date and brief description.",
   "\u25A0 New Program: What it is and how to enroll.",
   "\u25A0 Volunteer Opportunities: How to get involved.",
   "\u25A0 Important Deadline: Mark your calendar."
  ].forEach(function(a) {
    var ap = rc.body.insertParagraph(a, Word.InsertLocation.end);
    ap.font.set({ name: brand.bodyFont, size: 9, color: "#333333" });
    ap.spaceAfter = 5;
  });

  await context.sync();
}

async function buildEventProgram(context, brand) {
  var body = context.document.body;
  body.clear();
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  var section = sections.items[0];
  section.pageSetup.orientation = Word.PageOrientation.portrait;
  section.pageSetup.topMargin = 54;
  section.pageSetup.bottomMargin = 54;
  section.pageSetup.leftMargin = 72;
  section.pageSetup.rightMargin = 72;
  await context.sync();

  var title = body.insertParagraph("Event Name", Word.InsertLocation.end);
  title.font.set({ name: brand.headingFont, size: 24, bold: true, color: brand.primary });
  title.alignment = Word.Alignment.center;
  title.spaceAfter = 3;

  var sub = body.insertParagraph("Hosted by " + (brand.name || "Organization Name"), Word.InsertLocation.end);
  sub.font.set({ name: brand.bodyFont, size: 11, italic: true, color: brand.accent });
  sub.alignment = Word.Alignment.center;
  sub.spaceAfter = 3;

  var edate = body.insertParagraph("Saturday, Month Day, Year  |  Location Name", Word.InsertLocation.end);
  edate.font.set({ name: brand.bodyFont, size: 10, color: "#555555" });
  edate.alignment = Word.Alignment.center;
  edate.spaceAfter = 10;

  var d1 = body.insertParagraph("\u2500".repeat(50), Word.InsertLocation.end);
  d1.font.set({ color: brand.accent, size: 6 });
  d1.alignment = Word.Alignment.center;
  d1.spaceAfter = 10;

  var st = body.insertParagraph("Program Schedule", Word.InsertLocation.end);
  st.font.set({ name: brand.headingFont, size: 14, bold: true, color: brand.primary });
  st.alignment = Word.Alignment.center;
  st.spaceAfter = 8;
  await context.sync();

  var schedData = [
    ["Time", "Activity", "Speaker / Notes"],
    ["9:00 AM", "Registration & Welcome Coffee", "Main Lobby"],
    ["9:30 AM", "Opening Remarks", "Speaker Name, Title"],
    ["10:00 AM", "Keynote Address", "Keynote Speaker Name"],
    ["11:00 AM", "Panel Discussion", "Panelist 1, Panelist 2, Panelist 3"],
    ["12:00 PM", "Lunch Break", "Dining Hall"],
    ["1:00 PM", "Breakout Session A", "Room 101"],
    ["1:00 PM", "Breakout Session B", "Room 202"],
    ["2:30 PM", "Closing Remarks & Awards", "Main Hall"],
    ["3:00 PM", "Networking Reception", "Courtyard"],
  ];

  var table = body.insertTable(schedData.length, 3, Word.InsertLocation.end, schedData);
  var headerRow = table.getRow(0);
  headerRow.font.set({ name: brand.headingFont, bold: true, size: 10, color: "#FFFFFF" });
  headerRow.shadingColor = brand.primary;

  for (var i = 1; i < schedData.length; i++) {
    var row = table.getRow(i);
    row.font.set({ name: brand.bodyFont, size: 9, color: "#333333" });
    if (i % 2 === 0) { row.shadingColor = "#F5F7FA"; }
  }
  for (var j = 1; j < schedData.length; j++) {
    var cell = table.getCell(j, 0);
    cell.body.font.set({ bold: true, color: brand.accent });
  }
  await context.sync();

  var fn = body.insertParagraph("Schedule is subject to change. For questions, contact info@example.org.", Word.InsertLocation.end);
  fn.font.set({ name: brand.bodyFont, size: 8, italic: true, color: "#888888" });
  fn.alignment = Word.Alignment.center;
  fn.spaceBefore = 10;
  await context.sync();
}

async function buildBusinessCard(context, brand) {
  var body = context.document.body;
  body.clear();
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  var section = sections.items[0];
  section.pageSetup.orientation = Word.PageOrientation.landscape;
  section.pageSetup.pageWidth = 252;
  section.pageSetup.pageHeight = 144;
  section.pageSetup.topMargin = 14;
  section.pageSetup.bottomMargin = 14;
  section.pageSetup.leftMargin = 18;
  section.pageSetup.rightMargin = 18;
  await context.sync();

  var logo = body.insertParagraph("[Logo]", Word.InsertLocation.end);
  logo.font.set({ name: brand.headingFont, size: 8, bold: true, color: brand.primary });
  logo.spaceAfter = 2;

  var name = body.insertParagraph("Full Name", Word.InsertLocation.end);
  name.font.set({ name: brand.headingFont, size: 12, bold: true, color: brand.primary });
  name.spaceAfter = 1;

  var role = body.insertParagraph("Title / Position", Word.InsertLocation.end);
  role.font.set({ name: brand.bodyFont, size: 8, italic: true, color: brand.accent });
  role.spaceAfter = 4;

  var d = body.insertParagraph("\u2500".repeat(40), Word.InsertLocation.end);
  d.font.set({ color: "#E1E5EA", size: 4 });
  d.spaceAfter = 4;

  ["phone: (555) 123-4567", "email: name@example.org", "web: www.example.org"].forEach(function(line) {
    var cp = body.insertParagraph(line, Word.InsertLocation.end);
    cp.font.set({ name: brand.bodyFont, size: 7, color: "#555555" });
    cp.spaceAfter = 1;
  });
  await context.sync();
}

async function buildFlyer(context, brand) {
  var body = context.document.body;
  body.clear();
  var sections = context.document.sections;
  sections.load("items");
  await context.sync();

  var section = sections.items[0];
  section.pageSetup.orientation = Word.PageOrientation.portrait;
  section.pageSetup.topMargin = 54;
  section.pageSetup.bottomMargin = 54;
  section.pageSetup.leftMargin = 54;
  section.pageSetup.rightMargin = 54;
  await context.sync();

  var spacer = body.insertParagraph("", Word.InsertLocation.end);
  spacer.spaceAfter = 36;

  var en = body.insertParagraph("EVENT NAME", Word.InsertLocation.end);
  en.font.set({ name: brand.headingFont, size: 32, bold: true, color: brand.primary });
  en.alignment = Word.Alignment.center;
  en.spaceAfter = 6;

  var sub = body.insertParagraph("A Brief, Compelling Subtitle", Word.InsertLocation.end);
  sub.font.set({ name: brand.bodyFont, size: 14, italic: true, color: brand.accent });
  sub.alignment = Word.Alignment.center;
  sub.spaceAfter = 16;

  var img = body.insertParagraph("[Insert hero image or graphic here]", Word.InsertLocation.end);
  img.font.set({ name: brand.bodyFont, size: 10, italic: true, color: "#AAAAAA" });
  img.alignment = Word.Alignment.center;
  img.spaceAfter = 16;

  ["Date: Saturday, Month Day, Year", "Time: 6:00 PM - 9:00 PM", "Location: Venue Name, Address"].forEach(function(d) {
    var dp = body.insertParagraph(d, Word.InsertLocation.end);
    dp.font.set({ name: brand.bodyFont, size: 12, color: "#333333" });
    dp.alignment = Word.Alignment.center;
    dp.spaceAfter = 3;
  });

  var sp2 = body.insertParagraph("", Word.InsertLocation.end);
  sp2.spaceAfter = 12;

  var desc = body.insertParagraph("A short description of the event (2-3 sentences). What will attendees experience? Why should they come? Keep it energetic and direct.", Word.InsertLocation.end);
  desc.font.set({ name: brand.bodyFont, size: 11, color: "#555555" });
  desc.alignment = Word.Alignment.center;
  desc.spaceAfter = 16;

  var cta = body.insertParagraph("RSVP at www.example.org/events", Word.InsertLocation.end);
  cta.font.set({ name: brand.headingFont, size: 13, bold: true, color: brand.accent });
  cta.alignment = Word.Alignment.center;
  cta.spaceAfter = 8;

  var contact = body.insertParagraph("Questions? Contact info@example.org  |  (555) 123-4567", Word.InsertLocation.end);
  contact.font.set({ name: brand.bodyFont, size: 9, color: "#888888" });
  contact.alignment = Word.Alignment.center;
  await context.sync();
}

export { TEMPLATES };

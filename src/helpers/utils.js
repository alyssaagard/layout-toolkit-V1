/**
 * Layout Toolkit - Helper Utilities
 *
 * Brand profile storage, status messaging, and document helpers.
 */

// ============================================
// STATUS MESSAGING
// ============================================

const statusEl = () => document.getElementById("status-message");

function setStatus(message, type = "idle") {
  const el = statusEl();
  if (!el) return;
  el.textContent = message;
  el.className = `status-${type}`;
}

function setWorking(message) {
  setStatus(message, "working");
}

function setSuccess(message) {
  setStatus(message, "success");
  setTimeout(() => setStatus("Ready", "idle"), 4000);
}

function setError(message) {
  setStatus(message, "error");
  setTimeout(() => setStatus("Ready", "idle"), 6000);
}

// ============================================
// BRAND PROFILE MANAGEMENT
// ============================================

const STORAGE_KEY = "layoutToolkit_brandProfiles";
const ACTIVE_KEY = "layoutToolkit_activeBrand";

const DEFAULT_BRAND = {
  name: "",
  primary: "#1B3A5C",
  accent: "#C0392B",
  headingFont: "Georgia",
  bodyFont: "Calibri",
};

/**
 * Get current brand settings from the form inputs.
 */
function getBrandFromForm() {
  return {
    name: document.getElementById("brand-name")?.value || "",
    primary: document.getElementById("brand-primary")?.value || DEFAULT_BRAND.primary,
    accent: document.getElementById("brand-accent")?.value || DEFAULT_BRAND.accent,
    headingFont: document.getElementById("brand-heading-font")?.value || DEFAULT_BRAND.headingFont,
    bodyFont: document.getElementById("brand-body-font")?.value || DEFAULT_BRAND.bodyFont,
  };
}

/**
 * Load brand settings into the form inputs.
 */
function loadBrandToForm(brand) {
  const fields = {
    "brand-name": brand.name,
    "brand-primary": brand.primary,
    "brand-accent": brand.accent,
    "brand-heading-font": brand.headingFont,
    "brand-body-font": brand.bodyFont,
  };

  Object.entries(fields).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el) el.value = value;
  });

  // Update color hex displays
  updateColorHex("brand-primary");
  updateColorHex("brand-accent");
}

/**
 * Update the hex display span for a color input.
 */
function updateColorHex(inputId) {
  const input = document.getElementById(inputId);
  const span = document.querySelector(`.color-hex[data-for="${inputId}"]`);
  if (input && span) {
    span.textContent = input.value.toUpperCase();
  }
}

/**
 * Save a brand profile to localStorage.
 */
function saveBrandProfile(profileName, brand) {
  const profiles = getAllBrandProfiles();
  profiles[profileName] = brand;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
  localStorage.setItem(ACTIVE_KEY, profileName);
}

/**
 * Get all saved brand profiles.
 */
function getAllBrandProfiles() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {};
  } catch {
    return {};
  }
}

/**
 * Delete a brand profile by name.
 */
function deleteBrandProfile(profileName) {
  const profiles = getAllBrandProfiles();
  delete profiles[profileName];
  localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
}

/**
 * Get the active brand settings (from form or default).
 */
function getActiveBrand() {
  const formBrand = getBrandFromForm();
  // If the form has a name, use form values; otherwise use defaults
  if (formBrand.name) return formBrand;
  return { ...DEFAULT_BRAND, ...formBrand };
}

// ============================================
// DOCUMENT HELPERS
// ============================================

/**
 * Apply page orientation and paper size to the active section.
 */
async function applyPageSetup(orientation, paperSize) {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    const section = sections.items[0];

    // Orientation
    section.pageSetup.orientation =
      orientation === "Landscape"
        ? Word.PageOrientation.landscape
        : Word.PageOrientation.portrait;

    // Paper size (dimensions in points: 1 inch = 72 points)
    const sizes = {
      letter:  { width: 612, height: 792 },
      legal:   { width: 612, height: 1008 },
      tabloid: { width: 792, height: 1224 },
      a4:      { width: 595, height: 842 },
      a5:      { width: 420, height: 595 },
    };

    const size = sizes[paperSize] || sizes.letter;

    if (orientation === "Landscape") {
      section.pageSetup.pageWidth = size.height;
      section.pageSetup.pageHeight = size.width;
    } else {
      section.pageSetup.pageWidth = size.width;
      section.pageSetup.pageHeight = size.height;
    }

    await context.sync();
  });
}

/**
 * Apply margins to the active section.
 */
async function applyMargins(top, bottom, left, right) {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    const section = sections.items[0];
    section.pageSetup.topMargin = top * 72;
    section.pageSetup.bottomMargin = bottom * 72;
    section.pageSetup.leftMargin = left * 72;
    section.pageSetup.rightMargin = right * 72;

    await context.sync();
  });
}

/**
 * Insert a styled text placeholder at the cursor.
 */
async function insertTextPlaceholder() {
  await Word.run(async (context) => {
    const brand = getActiveBrand();
    const range = context.document.getSelection();

    const ph = range.insertParagraph("[Enter your text here]", Word.InsertLocation.after);
    ph.font.set({ name: brand.bodyFont, size: 11, italic: true, color: "#999999" });
    ph.spaceAfter = 10;

    await context.sync();
  });
}

/**
 * Insert a horizontal divider at the cursor.
 */
async function insertDivider() {
  await Word.run(async (context) => {
    const brand = getActiveBrand();
    const range = context.document.getSelection();

    const div = range.insertParagraph("\u2500".repeat(80), Word.InsertLocation.after);
    div.font.set({ size: 6, color: brand.primary });
    div.alignment = Word.Alignment.center;
    div.spaceAfter = 10;
    div.spaceBefore = 10;

    await context.sync();
  });
}

/**
 * Insert a pull quote block at the cursor.
 */
async function insertPullQuote() {
  await Word.run(async (context) => {
    const brand = getActiveBrand();
    const range = context.document.getSelection();

    const openQuote = range.insertParagraph("\u201C", Word.InsertLocation.after);
    openQuote.font.set({ name: brand.headingFont, size: 36, color: brand.accent });
    openQuote.alignment = Word.Alignment.center;
    openQuote.spaceAfter = 0;

    const quoteText = openQuote.insertParagraph(
      "Insert a compelling quote or statistic here that you want to highlight for the reader.",
      Word.InsertLocation.after
    );
    quoteText.font.set({ name: brand.headingFont, size: 14, italic: true, color: brand.primary });
    quoteText.alignment = Word.Alignment.center;
    quoteText.spaceAfter = 2;

    const attribution = quoteText.insertParagraph("\u2014 Attribution", Word.InsertLocation.after);
    attribution.font.set({ name: brand.bodyFont, size: 10, color: "#888888" });
    attribution.alignment = Word.Alignment.center;
    attribution.spaceAfter = 14;

    await context.sync();
  });
}

/**
 * Insert a sidebar/callout box using a single-cell table.
 */
async function insertSidebar() {
  await Word.run(async (context) => {
    const brand = getActiveBrand();
    const range = context.document.getSelection();

    // Insert a single-cell table as a "box"
    const table = range.insertTable(1, 1, Word.InsertLocation.after, [
      [""],
    ]);

    // Style the box
    table.getBorder(Word.BorderLocation.all).color = brand.primary;
    table.getBorder(Word.BorderLocation.all).width = 1.5;
    table.getCell(0, 0).shadingColor = "#F5F7FA";

    const cell = table.getCell(0, 0);
    const sideTitle = cell.body.insertParagraph("Sidebar Title", Word.InsertLocation.end);
    sideTitle.font.set({ name: brand.headingFont, size: 12, bold: true, color: brand.primary });
    sideTitle.spaceAfter = 6;

    const sideBody = cell.body.insertParagraph(
      "Use this box for supplementary information, tips, contact details, or callouts that should stand apart from the main content flow.",
      Word.InsertLocation.end
    );
    sideBody.font.set({ name: brand.bodyFont, size: 10, color: "#333333" });

    await context.sync();
  });
}

export {
  setStatus,
  setWorking,
  setSuccess,
  setError,
  getBrandFromForm,
  loadBrandToForm,
  updateColorHex,
  saveBrandProfile,
  getAllBrandProfiles,
  deleteBrandProfile,
  getActiveBrand,
  applyPageSetup,
  applyMargins,
  insertTextPlaceholder,
  insertDivider,
  insertPullQuote,
  insertSidebar,
  DEFAULT_BRAND,
};

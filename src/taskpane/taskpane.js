/**
 * Layout Toolkit - Main Taskpane Controller
 *
 * Initializes Office.js, renders the template gallery, and binds
 * all UI events to their corresponding Word document operations.
 */

import "./taskpane.css";
import { TEMPLATES } from "../templates/templateDefinitions.js";
import {
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
} from "../helpers/utils.js";

// ============================================
// OFFICE INITIALIZATION
// ============================================

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initApp();
  } else {
    document.getElementById("app").innerHTML =
      '<p style="padding:20px;color:#C0392B;">This add-in requires Microsoft Word.</p>';
  }
});

function initApp() {
  renderTemplateGrid();
  bindTabNavigation();
  bindFilterButtons();
  bindLayoutControls();
  bindBrandControls();
  loadSavedBrand();
  renderBrandProfiles();
}

// ============================================
// TAB NAVIGATION
// ============================================

function bindTabNavigation() {
  const tabs = document.querySelectorAll(".tab-btn");
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      // Deactivate all
      tabs.forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));

      // Activate selected
      tab.classList.add("active");
      const targetId = `tab-${tab.dataset.tab}`;
      document.getElementById(targetId)?.classList.add("active");
    });
  });
}

// ============================================
// TEMPLATE GALLERY
// ============================================

function renderTemplateGrid(filterCategory = "all") {
  const grid = document.getElementById("template-grid");
  if (!grid) return;

  const filtered =
    filterCategory === "all"
      ? TEMPLATES
      : TEMPLATES.filter((t) => t.category === filterCategory);

  grid.innerHTML = filtered
    .map(
      (t) => `
    <div class="template-card" data-template-id="${t.id}">
      <div class="template-card-preview">${t.thumbnail}</div>
      <div class="template-card-info">
        <h4>${t.name}</h4>
        <p>${t.description}</p>
        <span class="template-card-badge badge-${t.category}">${t.category}</span>
      </div>
    </div>
  `
    )
    .join("");

  // Bind click handlers
  grid.querySelectorAll(".template-card").forEach((card) => {
    card.addEventListener("click", () => {
      const templateId = card.dataset.templateId;
      confirmAndBuildTemplate(templateId);
    });
  });
}

function bindFilterButtons() {
  const filters = document.querySelectorAll(".filter-btn");
  filters.forEach((btn) => {
    btn.addEventListener("click", () => {
      filters.forEach((f) => f.classList.remove("active"));
      btn.classList.add("active");
      renderTemplateGrid(btn.dataset.category);
    });
  });
}

/**
 * Show a confirmation modal before building a template
 * (since it clears the document).
 */
function confirmAndBuildTemplate(templateId) {
  const template = TEMPLATES.find((t) => t.id === templateId);
  if (!template) return;

  // Create modal
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.innerHTML = `
    <div class="modal-content">
      <h3>${template.name}</h3>
      <p>This will replace the current document content with the "${template.name}" layout. Your brand settings will be applied.</p>
      <div class="modal-actions">
        <button class="btn btn-secondary" id="modal-cancel">Cancel</button>
        <button class="btn btn-primary" id="modal-confirm">Build Layout</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  document.getElementById("modal-cancel").addEventListener("click", () => {
    overlay.remove();
  });

  document.getElementById("modal-confirm").addEventListener("click", async () => {
    overlay.remove();
    await buildTemplate(template);
  });

  // Close on overlay background click
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });
}

async function buildTemplate(template) {
  setWorking(`Building ${template.name}...`);

  try {
    await Word.run(async (context) => {
      const brand = getActiveBrand();
      await template.build(context, brand);
    });
    setSuccess(`${template.name} created successfully!`);
  } catch (err) {
    console.error("Template build error:", err);
    setError(`Error: ${err.message || "Failed to build template."}`);
  }
}

// ============================================
// LAYOUT CONTROLS
// ============================================

function bindLayoutControls() {
  // Page Setup
  document.getElementById("btn-apply-page")?.addEventListener("click", async () => {
    const orientation = document.getElementById("orientation-select").value;
    const paperSize = document.getElementById("paper-select").value;

    setWorking("Applying page setup...");
    try {
      await applyPageSetup(orientation, paperSize);
      setSuccess("Page setup applied.");
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });

  // Columns (applied via OOXML injection for proper column support)
  document.getElementById("btn-apply-columns")?.addEventListener("click", async () => {
    const colCount = parseInt(document.getElementById("col-count").value, 10);
    const spacing = parseFloat(document.getElementById("col-spacing").value);
    const showLine = document.getElementById("col-line").checked;

    setWorking("Applying columns...");
    try {
      await applyColumns(colCount, spacing, showLine);
      setSuccess(`${colCount}-column layout applied.`);
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });

  // Margins
  document.getElementById("btn-apply-margins")?.addEventListener("click", async () => {
    const top = parseFloat(document.getElementById("margin-top").value);
    const bottom = parseFloat(document.getElementById("margin-bottom").value);
    const left = parseFloat(document.getElementById("margin-left").value);
    const right = parseFloat(document.getElementById("margin-right").value);

    setWorking("Applying margins...");
    try {
      await applyMargins(top, bottom, left, right);
      setSuccess("Margins applied.");
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });

  // Insert elements
  document.getElementById("btn-insert-textbox")?.addEventListener("click", async () => {
    setWorking("Inserting placeholder...");
    try {
      await insertTextPlaceholder();
      setSuccess("Text placeholder inserted.");
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });

  document.getElementById("btn-insert-divider")?.addEventListener("click", async () => {
    setWorking("Inserting divider...");
    try {
      await insertDivider();
      setSuccess("Divider inserted.");
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });

  document.getElementById("btn-insert-pull-quote")?.addEventListener("click", async () => {
    setWorking("Inserting pull quote...");
    try {
      await insertPullQuote();
      setSuccess("Pull quote inserted.");
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });

  document.getElementById("btn-insert-sidebar")?.addEventListener("click", async () => {
    setWorking("Inserting sidebar...");
    try {
      await insertSidebar();
      setSuccess("Sidebar box inserted.");
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  });
}

/**
 * Apply columns using OOXML section properties.
 * The Word JS API does not expose columns directly, so we use
 * OOXML insertion to set column properties on the current section.
 */
async function applyColumns(count, spacingInches, showLine) {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    // Build OOXML column definition
    const spacingTwips = Math.round(spacingInches * 1440); // 1 inch = 1440 twips
    const sep = showLine ? ' w:sep="1"' : "";

    let colsXml;
    if (count === 1) {
      colsXml = `<w:cols w:space="${spacingTwips}"/>`;
    } else {
      colsXml = `<w:cols w:num="${count}" w:space="${spacingTwips}"${sep} w:equalWidth="1"/>`;
    }

    // We apply this by inserting an OOXML section break with column defs.
    // For simplicity, we insert a paragraph with section properties.
    const body = context.document.body;
    const ooxml = `
      <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
          <pkg:xmlData>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:pPr>
                    <w:sectPr>
                      ${colsXml}
                    </w:sectPr>
                  </w:pPr>
                </w:p>
              </w:body>
            </w:document>
          </pkg:xmlData>
        </pkg:part>
      </pkg:package>
    `.trim();

    body.insertOoxml(ooxml, Word.InsertLocation.end);
    await context.sync();
  });
}

// ============================================
// BRAND CONTROLS
// ============================================

function bindBrandControls() {
  // Color hex display sync
  ["brand-primary", "brand-accent"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", () => {
      updateColorHex(id);
    });
  });

  // Save brand
  document.getElementById("btn-save-brand")?.addEventListener("click", () => {
    const brand = getBrandFromForm();
    const profileName = brand.name || "Default";

    saveBrandProfile(profileName, brand);
    setSuccess(`Brand profile "${profileName}" saved.`);
    renderBrandProfiles();
  });
}

function loadSavedBrand() {
  const profiles = getAllBrandProfiles();
  const keys = Object.keys(profiles);

  if (keys.length > 0) {
    // Load the most recently saved profile
    const lastName = keys[keys.length - 1];
    loadBrandToForm(profiles[lastName]);
  } else {
    loadBrandToForm(DEFAULT_BRAND);
  }
}

function renderBrandProfiles() {
  const container = document.getElementById("brand-profiles-list");
  if (!container) return;

  const profiles = getAllBrandProfiles();
  const keys = Object.keys(profiles);

  if (keys.length === 0) {
    container.innerHTML = '<p class="empty-state">No saved profiles yet.</p>';
    return;
  }

  container.innerHTML = keys
    .map(
      (name) => `
    <div class="profile-item" data-profile="${name}">
      <div class="profile-item-info">
        <div class="profile-swatch" style="background:${profiles[name].primary}"></div>
        <span class="profile-item-name">${name}</span>
      </div>
      <div class="profile-item-actions">
        <button class="profile-load" data-profile="${name}">Load</button>
        <button class="profile-delete" data-profile="${name}">Delete</button>
      </div>
    </div>
  `
    )
    .join("");

  // Bind load / delete buttons
  container.querySelectorAll(".profile-load").forEach((btn) => {
    btn.addEventListener("click", () => {
      const name = btn.dataset.profile;
      const profile = profiles[name];
      if (profile) {
        loadBrandToForm(profile);
        setSuccess(`Loaded "${name}" brand profile.`);
      }
    });
  });

  container.querySelectorAll(".profile-delete").forEach((btn) => {
    btn.addEventListener("click", () => {
      const name = btn.dataset.profile;
      deleteBrandProfile(name);
      renderBrandProfiles();
      setSuccess(`Deleted "${name}" brand profile.`);
    });
  });
}

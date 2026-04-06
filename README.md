# Layout Toolkit V.1

A free, open-source word processing add-in that brings desktop publishing features to your documents. Create brochures, newsletters, event programs, business cards, and flyers with professional templates, brand customization, and built-in folding guides.

Built as an open-source alternative for desktop publishing workflows in word processors.

## Features

**Template Gallery** - Six ready-to-use layouts: tri-fold brochure, bi-fold brochure, classic newsletter, event program, business card, and event flyer. Each generates a complete Word document with placeholder content, formatted with your brand settings.

**Layout Tools** - Direct control over page orientation, paper size, columns (1-4 with configurable gutter spacing), margins, and insertable elements including text placeholders, horizontal dividers, pull quotes, and sidebar callout boxes.

**Brand Profiles** - Save and switch between organization-specific color schemes and font pairings. Brand settings are applied automatically when generating any template.

**Folding & Print Guide** - Built-in visual diagrams showing how to fold tri-fold and bi-fold brochures, along with paper selection and printing tips.

## Installation

### Prerequisites

- Microsoft Word (desktop app for Windows or Mac, or the free web version at office.com) (This version is most compatable for the Word web application.)
- For the desktop add in, see 
- For self-hosting during development: Node.js 18+

### Quick Start (Users)

If someone is hosting this add-in for you (e.g., your IT department or organization), they will provide a manifest.xml file or a direct install link. Follow the platform-specific sideloading steps below.

### For Developers / Self-Hosting

```bash
git clone https://github.com/alyssaagard/layout-toolkit.git
cd layout-toolkit
npm install
npm run dev
```

This starts a local HTTPS server on port 3000. Then sideload the manifest into Word using the instructions below.

For a production build:

```bash
npm run build
```

The `dist/` folder can be deployed to any static host (GitHub Pages, Netlify, your own server). Update the URLs in `manifest.xml` to point to your production host.

### Sideloading on Windows

1. Open Word and create a new document.
2. Click **Insert** in the ribbon.
3. Click **My Add-ins** (or **Get Add-ins**, then **My Add-ins** tab).
4. Click **Upload My Add-in**.
5. Browse to `manifest.xml` and click **Upload**.
6. The "Layout Toolkit" tab will appear in the ribbon.

### Sideloading on Mac

1. Open Finder. Press **Shift + Cmd + G** and go to:
   ```
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```
   If the `wef` folder does not exist, create it:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```
2. Copy `manifest.xml` into that folder:
   ```bash
   cp /path/to/layout-toolkit/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
3. Quit Word completely (Cmd + Q) and reopen it.
4. Click **Add-ins** in the ribbon. Layout Toolkit will appear under **My Add-ins**.

### Sideloading on Word for Web (office.com)

1. Open Word at office.com and create a new document.
2. Click **Insert** > **Add-ins** > **Upload My Add-in**.
3. Browse to `manifest.xml` and upload.

### Organizational Deployment

IT administrators can deploy the add-in to all users via the **Microsoft 365 admin center** under Settings > Integrated Apps. This eliminates the need for individual sideloading.

## Project Structure

```
layout-toolkit/
├── manifest.xml                    # Office Add-in manifest
├── package.json                    # Dependencies and scripts
├── webpack.config.js               # Build configuration
├── assets/                         # Add-in icons
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html           # UI: 4-tab layout (Templates, Layout, Branding, Guide)
│   │   ├── taskpane.css            # Stylesheet with CSS custom properties
│   │   └── taskpane.js             # Controller: Office init, events, template orchestration
│   ├── templates/
│   │   └── templateDefinitions.js  # Template metadata, SVG thumbnails, Word JS API builders
│   └── helpers/
│       └── utils.js                # Brand storage, status messaging, document helpers
├── LICENSE                         # MIT License
└── README.md
```

## Compatibility

Tested and working on:

| Platform | Status |
|---|---|
| Word for Mac (Microsoft 365) | Tested |
| Word for Windows (Microsoft 365) | Supported |
| Word for Windows (Office 2021) | Supported |
| Word for Web (office.com) | Supported (free) |
| Word for iPad | Partial (task pane add-ins supported) |

The codebase is identical across all platforms. Platform-specific API differences (e.g., border rendering on Mac) are handled with try/catch guards.

## Licensing

This project is licensed under the MIT License. No special license from Microsoft is required to create, distribute, or use a Word Add-in. Office.js is loaded from Microsoft's public CDN at no cost.

End users need a copy of Microsoft Word. The free web version at office.com is sufficient.

## Contributing

Contributions welcome. Please open an issue before submitting a pull request for major changes.

Areas where help is especially useful:
- Additional templates (booklets, certificates, letterheads, postcards)
- Accessibility improvements
- Localization / internationalization
- Testing on Word for Windows and Word for iPad
- A Google Docs (Apps Script) port

## Author

Developed by Alyssa I. Agard.

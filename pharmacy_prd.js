const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  TabStopPosition, PageBreak
} = require('docx');
const fs = require('fs');

// ─── Color Palette ───────────────────────────────────────────────
const COLORS = {
  primary:   "1A5276",   // deep navy blue
  accent:    "1ABC9C",   // teal/green
  lightBlue: "D6EAF8",   // table header fill
  lightGray: "F2F3F4",   // alt row fill
  darkText:  "1C2833",   // near-black
  white:     "FFFFFF",
  rule:      "1A5276",
};

// ─── Helpers ──────────────────────────────────────────────────────
const border = (color = "CCCCCC") => ({
  top:    { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left:   { style: BorderStyle.SINGLE, size: 1, color },
  right:  { style: BorderStyle.SINGLE, size: 1, color },
});

const noBorder = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function rule(color = COLORS.primary) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } },
    spacing: { after: 200 },
    children: [],
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 30, bold: true, color: COLORS.primary })],
    spacing: { before: 440, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.primary, space: 4 } },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: COLORS.accent })],
    spacing: { before: 320, after: 100 },
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: COLORS.primary })],
    spacing: { before: 200, after: 80 },
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: COLORS.darkText, ...opts })],
  });
}

function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: COLORS.darkText, bold })],
  });
}

function subbullet(text) {
  return new Paragraph({
    numbering: { reference: "subbullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20, color: COLORS.darkText })],
  });
}

function numbered(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: COLORS.darkText, bold })],
  });
}

function spacer(pts = 160) {
  return new Paragraph({ spacing: { after: pts }, children: [] });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── Info Box (highlighted callout) ──────────────────────────────
function infoBox(title, lines, fillColor = COLORS.lightBlue) {
  const rows = [];
  // Title row
  rows.push(new TableRow({
    children: [new TableCell({
      borders: border(COLORS.primary),
      shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: COLORS.white })] })],
    })],
  }));
  // Content rows
  lines.forEach(line => {
    rows.push(new TableRow({
      children: [new TableCell({
        borders: border(COLORS.primary),
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 160, right: 160 },
        children: [new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 20, color: COLORS.darkText })] })],
      })],
    }));
  });
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows,
  });
}

// ─── Two-column table ─────────────────────────────────────────────
function twoColTable(headers, rows) {
  const COL = [4680, 4680];
  const tableRows = [];
  // Header
  tableRows.push(new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: border(COLORS.primary),
      shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      width: { size: COL[i], type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Arial", size: 22, bold: true, color: COLORS.white })] })],
    })),
  }));
  // Data rows
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? COLORS.white : COLORS.lightGray;
    tableRows.push(new TableRow({
      children: row.map((cell, i) => new TableCell({
        borders: border("CCCCCC"),
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        width: { size: COL[i], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 20, color: COLORS.darkText })] })],
      })),
    }));
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: COL, rows: tableRows });
}

// ─── Three-column table ───────────────────────────────────────────
function threeColTable(headers, rows) {
  const COL = [2200, 4000, 3160];
  const tableRows = [];
  tableRows.push(new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: border(COLORS.primary),
      shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 120, right: 120 },
      width: { size: COL[i], type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Arial", size: 21, bold: true, color: COLORS.white })] })],
    })),
  }));
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? COLORS.white : COLORS.lightGray;
    tableRows.push(new TableRow({
      children: row.map((cell, i) => new TableCell({
        borders: border("CCCCCC"),
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: COL[i], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 20, color: COLORS.darkText })] })],
      })),
    }));
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: COL, rows: tableRows });
}

// ─── Phase card ───────────────────────────────────────────────────
function phaseCard(phase, title, duration, features) {
  const COL = [2400, 6960];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: COL,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: border(COLORS.primary),
            shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 160, right: 160 },
            verticalAlign: VerticalAlign.CENTER,
            rowSpan: features.length + 1,
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: phase, font: "Arial", size: 32, bold: true, color: COLORS.white })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: COLORS.lightBlue })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: duration, font: "Arial", size: 20, color: COLORS.lightBlue })] }),
            ],
          }),
          new TableCell({
            borders: border(COLORS.primary),
            shading: { fill: COLORS.lightBlue, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [new Paragraph({ children: [new TextRun({ text: "Key Deliverables", font: "Arial", size: 22, bold: true, color: COLORS.primary })] })],
          }),
        ],
      }),
      ...features.map((f, fi) => new TableRow({
        children: [
          new TableCell({
            borders: border("CCCCCC"),
            shading: { fill: fi % 2 === 0 ? COLORS.white : COLORS.lightGray, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 160, right: 160 },
            children: [new Paragraph({ children: [new TextRun({ text: `\u2022 ${f}`, font: "Arial", size: 20, color: COLORS.darkText })] })],
          }),
        ],
      })),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────
// DOCUMENT BUILD
// ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",    levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 300 } } } }] },
      { reference: "subbullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 900, hanging: 300 } } } }] },
      { reference: "numbers",    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 300 } } } }] },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: COLORS.darkText } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: "Arial", color: COLORS.primary }, paragraph: { spacing: { before: 440, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Arial", color: COLORS.accent }, paragraph: { spacing: { before: 320, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: "Arial", color: COLORS.primary }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.primary, space: 6 } },
            spacing: { after: 0 },
            children: [
              new TextRun({ text: "RetailRx Pharmacy Platform", font: "Arial", size: 18, bold: true, color: COLORS.primary }),
              new TextRun({ text: "\tProject Requirements Document", font: "Arial", size: 18, color: COLORS.accent }),
            ],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: COLORS.primary, space: 4 } },
            spacing: { before: 80 },
            children: [
              new TextRun({ text: "CONFIDENTIAL \u2014 For Internal Use Only", font: "Arial", size: 16, color: "888888" }),
              new TextRun({ text: "\tPage ", font: "Arial", size: 16, color: "888888" }),
              new PageNumber({ font: "Arial", size: 16, color: "888888" }),
            ],
          }),
        ],
      }),
    },
    children: [

      // ══════════════════════════════════════════════════════════
      // COVER PAGE
      // ══════════════════════════════════════════════════════════
      spacer(2400),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [
        new TextRun({ text: "RETAILRX", font: "Arial", size: 80, bold: true, color: COLORS.primary }),
      ]}),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [
        new TextRun({ text: "PHARMACY E-COMMERCE PLATFORM", font: "Arial", size: 40, bold: true, color: COLORS.accent }),
      ]}),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: COLORS.primary, space: 4 } },
        spacing: { after: 240 },
        children: [],
      }),
      spacer(120),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [
        new TextRun({ text: "Project Requirements & Technical Specification Document", font: "Arial", size: 30, color: COLORS.darkText }),
      ]}),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [
        new TextRun({ text: "Featuring Wellgo & Pediazone Brand Catalog Integration", font: "Arial", size: 26, italics: true, color: COLORS.accent }),
      ]}),
      spacer(240),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [
        new TextRun({ text: "Version 1.0  |  June 2026  |  CONFIDENTIAL", font: "Arial", size: 22, color: "888888" }),
      ]}),
      spacer(2000),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
        new TextRun({ text: "Prepared for: Pharmacy Operations & Technology Leadership", font: "Arial", size: 22, color: COLORS.darkText }),
      ]}),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [
        new TextRun({ text: "Classification: Internal / Restricted", font: "Arial", size: 20, color: "888888" }),
      ]}),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // EXECUTIVE SUMMARY
      // ══════════════════════════════════════════════════════════
      heading1("Executive Summary"),
      para("This document defines the complete functional, technical, security, and operational requirements for the development of RetailRx, a professional e-commerce platform for a licensed retail pharmacy. The platform is purpose-built to provide a seamless, secure, and clinically trustworthy digital experience for customers seeking to browse, evaluate, and purchase health products online."),
      spacer(80),
      para("The platform will prominently feature curated product catalogs from two flagship brands:"),
      bullet("Wellgo \u2014 a broad-spectrum adult wellness and nutritional supplement brand"),
      bullet("Pediazone \u2014 a specialized pediatric and family health brand"),
      spacer(120),
      para("Three strategic pillars underpin every design and engineering decision in this project:"),
      spacer(80),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({ children: [
            ["Trust & Professionalism", "The Pharmacist Factor", "Local SEO"].map((h, i) => new TableCell({
              borders: border(COLORS.primary), shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 140, right: 140 },
              width: { size: 3120, type: WidthType.DXA },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Arial", size: 22, bold: true, color: COLORS.white })] })],
            })),
          ]}),
          new TableRow({ children: [
            ["Clinical color palette (blues, whites, greens) and accurate medical information on every product page.", "Professional verification workflows and pharmacist-reviewed badges are a legal and trust requirement.", "Prominent business NAP (Name, Address, Phone) with Schema.org markup on every page for Google local ranking."].map((t, i) => new TableCell({
              borders: border("CCCCCC"), shading: { fill: COLORS.lightBlue, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              width: { size: 3120, type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 20, color: COLORS.darkText })] })],
            })),
          ]}),
        ],
      }),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 1
      // ══════════════════════════════════════════════════════════
      heading1("1. Technical & Security Foundation"),

      heading2("1.1 Legal & Regulatory Compliance"),
      para("The platform must comply with all applicable health-data, pharmacy, and consumer-protection regulations prior to launch. Compliance is non-negotiable and must be validated by qualified legal counsel before go-live."),
      spacer(80),
      heading3("1.1.1 Data Protection Frameworks"),
      bullet("HIPAA (US): If the platform stores, processes, or transmits Protected Health Information (PHI), a full HIPAA compliance program is mandatory. This includes a signed Business Associate Agreement (BAA) with all cloud and SaaS vendors, encryption at rest and in transit, audit logs for all PHI access, and a documented Breach Notification Policy."),
      bullet("GDPR (EU/UK users): Obtain explicit, granular consent before any data collection. Implement a compliant Cookie Consent banner. Honour Data Subject Access Requests (DSAR) within 30 days. Appoint a Data Protection Officer (DPO) if processing at scale."),
      bullet("PDPA / Local Laws: If operating in regions with local health-data laws (e.g., PDPA in Thailand/Singapore, POPIA in South Africa), map all data flows to each applicable law and document compliance controls."),
      spacer(80),
      heading3("1.1.2 Pharmacy Licensing & Product Regulations"),
      bullet("Display a valid pharmacy license number prominently in the site footer and on the About Us page."),
      bullet("Categorise all products by prescription status: OTC (Over-the-Counter), Prescription-Only (Rx), and Controlled Substances if applicable."),
      bullet("Integrate with the relevant national medicine regulatory database (e.g., FDA NDC Database, MHRA, or local equivalent) to verify product listings."),
      bullet("Publish a clear returns, disposal, and recall notification policy in compliance with local pharmacy board regulations."),
      bullet("Ensure pharmacist-of-record information is displayed for the dispensing pharmacy, including licensure details."),
      spacer(80),
      heading3("1.1.3 Accessibility & Consumer Protection"),
      bullet("WCAG 2.1 AA Accessibility: The platform must meet Level AA standards to comply with the ADA (US), EAA (EU), and equivalent legislation in other regions."),
      bullet("Age verification gate for any products restricted to adults (e.g., certain OTC medications, supplements with age restrictions)."),
      bullet("Full terms of service, privacy policy, and cookie policy must be reviewed by legal counsel and versioned in a document management system."),
      spacer(120),

      heading2("1.2 Data Security Architecture"),
      spacer(80),
      heading3("1.2.1 Transport & Encryption"),
      twoColTable(["Security Control", "Requirement"],
        [
          ["TLS/SSL Certificate", "TLS 1.3 enforced site-wide. HTTP Strict Transport Security (HSTS) header with minimum 1-year max-age. Automated certificate renewal via Let\u2019s Encrypt or a commercial CA."],
          ["Data at Rest", "AES-256 encryption for all database fields containing PHI, PII, prescription images, and payment tokens. Encryption keys managed in a dedicated KMS (e.g., AWS KMS, Google Cloud KMS)."],
          ["Data in Transit", "All internal service-to-service communication encrypted. No plaintext credentials over any network path."],
          ["Password Storage", "Passwords hashed with bcrypt (cost factor \u226512) or Argon2id. Salted per-user. No reversible password storage."],
        ]
      ),
      spacer(120),
      heading3("1.2.2 Payment Security (PCI-DSS Level 1)"),
      bullet("Integrate exclusively with a PCI-DSS Level 1 certified payment gateway such as Stripe, Braintree, or Adyen."),
      bullet("Never store raw card numbers, CVV, or full magnetic stripe data on platform servers. Use tokenisation exclusively."),
      bullet("Implement 3D Secure 2 (3DS2) for additional cardholder authentication on high-value or high-risk transactions."),
      bullet("Maintain a quarterly ASV (Approved Scanning Vendor) scan and annual PCI-DSS Self-Assessment Questionnaire (SAQ)."),
      bullet("Provide Apple Pay, Google Pay, and Buy Now Pay Later (BNPL) options via the gateway\u2019s hosted fields \u2014 these do not increase PCI scope."),
      spacer(80),
      heading3("1.2.3 Application Security"),
      bullet("OWASP Top 10: All code reviewed against OWASP Top 10 vulnerabilities. Mandatory penetration testing before launch and annually thereafter."),
      bullet("Web Application Firewall (WAF): Deploy a WAF (e.g., Cloudflare WAF, AWS WAF) in front of all public endpoints to block SQLi, XSS, and CSRF attacks."),
      bullet("Rate Limiting: Apply per-IP and per-account rate limiting on login, signup, password reset, prescription upload, and checkout endpoints."),
      bullet("Multi-Factor Authentication (MFA): MFA required for all admin panel and pharmacist dashboard logins. Offer optional MFA to end customers."),
      bullet("Dependency Scanning: Integrate Snyk or Dependabot into the CI/CD pipeline to flag vulnerable npm/pip packages before deployment."),
      bullet("Content Security Policy (CSP): Strict CSP headers to prevent XSS via injected scripts."),
      spacer(80),
      heading3("1.2.4 Prescription Data Storage"),
      bullet("Prescription images stored in an isolated, access-controlled cloud bucket (not publicly accessible URLs)."),
      bullet("Each prescription file encrypted at rest and accessible only by authenticated pharmacists with an audit log entry on every access."),
      bullet("Automatic deletion of prescription images after a configurable retention period (e.g., 3 years) in compliance with local record-keeping laws."),
      bullet("Virus/malware scanning of all uploaded files before storage (e.g., ClamAV integration)."),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 2
      // ══════════════════════════════════════════════════════════
      heading1("2. Core Functional Requirements"),

      heading2("2.1 User Experience (UX) & Design System"),
      para("The platform\u2019s design must communicate clinical authority, trustworthiness, and ease of use. Customers must never feel uncertain about the legitimacy of the pharmacy or the accuracy of product information."),
      spacer(80),
      heading3("2.1.1 Design Principles & Visual Identity"),
      bullet("Color Palette: Primary navy blue (#1A5276), supporting teal (#1ABC9C), white backgrounds, and clean grey dividers. Avoid warm or playful color schemes that reduce clinical credibility."),
      bullet("Typography: Clean, legible sans-serif fonts (Inter, Source Sans Pro, or Roboto). Minimum 16px body text for readability. High-contrast text (\u22654.5:1 ratio) throughout."),
      bullet("Verified Pharmacist Badge: A prominent, consistent badge displayed on all reviewed products, fulfilled orders, and pharmacist-verified prescriptions. This is a legal trust signal, not merely decorative."),
      bullet("Mobile-First Responsive Design: Design for 375px mobile width first, scaling to tablet (768px) and desktop (1280px+). All UI interactions must be fully functional on touch devices."),
      spacer(80),
      heading3("2.1.2 Navigation Architecture"),
      bullet("Global Persistent Header: Logo, Smart Search Bar, Account Icon (login/profile), Cart Icon with item count, and a Pharmacist Help CTA."),
      bullet("Mega Menu Navigation: Top-level categories (Wellgo, Pediazone, Vitamins & Supplements, Cold & Flu, Baby & Child, Prescription Services, Health Blog). Each category reveals brand, subcategory, and featured product tiles on hover/tap."),
      bullet("Breadcrumb Trail: Present on all product, category, and content pages for orientation and SEO benefit."),
      bullet("Footer (Every Page): Full business NAP (Name, Address, Phone, Opening Hours), pharmacy license number, quick links, social media, and Schema.org LocalBusiness markup."),
      spacer(80),
      heading3("2.1.3 Smart Search Bar"),
      bullet("Powered by Elasticsearch or Algolia with typo-tolerance, synonym matching, and fuzzy search."),
      bullet("Autocomplete with product images, brand labels, and category suggestions appearing as the user types."),
      bullet("Search by product name, active ingredient, brand name (Wellgo, Pediazone), health concern (e.g., \u201Ccough relief\u201D), or symptom."),
      bullet("Zero-results page includes suggested alternatives, a \u201CAsk a Pharmacist\u201D CTA, and popular products."),
      bullet("Search analytics dashboard in the admin panel to identify high-volume search terms with no results \u2014 these are catalog gaps."),
      spacer(80),
      heading3("2.1.4 Filtered Browsing & Category Pages"),
      bullet("Left-rail filter panel (desktop) / bottom-sheet filter drawer (mobile) with the following filter dimensions:"),
      subbullet("Brand: Wellgo, Pediazone, and all other stocked brands"),
      subbullet("Health Concern / Symptom: Organised into a taxonomy (e.g., Digestive Health, Immunity, Paediatric Care, Vitamins & Minerals)"),
      subbullet("Price Range: Slider control with min/max input"),
      subbullet("Product Form: Tablet, Capsule, Syrup, Drops, Topical"),
      subbullet("Age Suitability: Infant, Child, Adult, Senior"),
      subbullet("Availability: In Stock, Pre-Order"),
      subbullet("Prescription Status: OTC Only"),
      bullet("Sorting options: Relevance (default), Newest, Price Low-High, Price High-Low, Best Reviewed, Most Purchased."),
      bullet("Filter selections shown as removable chips above the product grid. \u201CClear All\u201D control available."),
      spacer(120),

      heading2("2.2 Prescription Management Workflow"),
      para("This workflow is a legally sensitive process. Every step must be logged with a timestamp, user ID, and IP address."),
      spacer(80),

      threeColTable(["Step", "Customer Action", "System / Pharmacist Action"],
        [
          ["1. Upload", "Selects product tagged Prescription Required. Prompted to upload prescription via camera capture or file upload (JPG, PNG, PDF, max 10MB).", "Virus scan file. Encrypt and store in isolated bucket. Create Prescription record with status: Pending Review."],
          ["2. Confirmation", "Receives email and on-screen confirmation with a reference number and estimated review time (e.g., 2\u20134 business hours).", "Notify available pharmacists via dashboard alert."],
          ["3. Pharmacist Review", "Views status tracker in My Account > Prescriptions.", "Pharmacist reviews prescription validity, prescribing doctor details, dosage, and product match. Can Approve, Reject, or Request More Info."],
          ["4. Verification", "If approved: order proceeds to checkout. If rejected: customer is notified with reason and recommended next step.", "Pharmacist logs decision with notes. System applies Verified by Pharmacist badge to the order."],
          ["5. Refill Reminder", "Receives automated reminder X days before estimated supply runs out (configurable).", "System checks refill eligibility based on prescription validity date. Flags expired prescriptions for renewal."],
        ]
      ),
      spacer(120),

      heading2("2.3 Account Management Features"),
      heading3("2.3.1 Authentication & Security"),
      bullet("Secure registration via email/password or social OAuth (Google, Apple). Email verification required before placing first order."),
      bullet("Password strength enforcement (minimum 12 characters, mixed case, numbers, symbols). Breach detection via HaveIBeenPwned API integration."),
      bullet("Optional TOTP-based MFA (Google Authenticator / Authy). Recovery codes generated and shown once at enrollment."),
      bullet("Session management: JWT access tokens (15-minute expiry) with rotating refresh tokens. Force logout from all sessions option."),
      spacer(80),
      heading3("2.3.2 My Account Dashboard"),
      bullet("Order History: Full order list with status, items, invoice download, and one-click Reorder functionality."),
      bullet("Reorder: Adds all items from a previous order to the cart in one action, checking current availability and price."),
      bullet("Prescription Vault: Secure list of all uploaded prescriptions with status indicator, approval history, and refill countdown."),
      bullet("Refill Reminders: Customer can configure automated email/SMS/push notifications for each active prescription."),
      bullet("Saved Addresses: Multiple delivery addresses with a default selection."),
      bullet("Saved Payment Methods: Tokenised card display (last 4 digits, expiry) via PCI-compliant vault. Delete at any time."),
      bullet("Health Profile (Optional): Age range, known allergies, chronic conditions \u2014 used to surface relevant product recommendations and safety warnings. Stored as PHI with full encryption."),
      bullet("Loyalty / Points: If a rewards program is implemented, points balance and redemption history displayed here."),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 3
      // ══════════════════════════════════════════════════════════
      heading1("3. Product Catalog Strategy"),

      heading2("3.1 Brand Spotlight: Wellgo & Pediazone"),
      para("Wellgo and Pediazone are the platform\u2019s flagship brands and must receive elevated visual treatment and discoverability throughout the site."),
      spacer(80),

      twoColTable(["Wellgo", "Pediazone"],
        [
          ["Dedicated brand landing page with hero banner, brand story, and product showcase.", "Dedicated brand landing page with child-safe, friendly design elements within the clinical framework."],
          ["Featured placement in the homepage hero carousel and \u2018Top Brands\u2019 section.", "Featured placement in the \u2018Baby & Child Health\u2019 homepage section with age-range selectors."],
          ["Wellgo filter shortcut in the global navigation mega menu.", "Pediazone filter shortcut accessible from Baby & Child category and the main navigation."],
          ["Wellgo badge displayed on all product cards and product detail pages.", "Pediazone badge with a \u2018Pediatrician Recommended\u2019 callout where applicable."],
          ["Email campaign integration: Wellgo new arrivals and promotions segment.", "Email campaign: Pediazone growth-stage reminders (e.g., newborn, 6 months, 1 year, 3 years)."],
        ]
      ),
      spacer(120),

      heading2("3.2 Product Catalog Structure & Categories"),
      para("The following top-level category taxonomy applies to the full catalog. Wellgo and Pediazone products are tagged to one or more of these categories."),
      spacer(80),

      twoColTable(["Category", "Sub-categories (Examples)"],
        [
          ["Vitamins & Supplements", "Multivitamins, Vitamin C, Vitamin D, Omega-3, Iron, Calcium, Probiotics"],
          ["Baby & Child Health (Pediazone)", "Infant Vitamins, Teething Relief, Pediatric Cold & Flu, Baby Skin Care, Growth Supplements"],
          ["Adult Wellness (Wellgo)", "Immune Support, Energy & Vitality, Joint Health, Digestive Wellness, Men\u2019s Health, Women\u2019s Health"],
          ["Cold, Flu & Allergy", "Antihistamines, Decongestants, Cough Syrups, Lozenges, Nasal Sprays"],
          ["Skin, Hair & Beauty", "Topical Creams, Sunscreen, Wound Care, Dermatological Products"],
          ["Chronic Disease Management", "Diabetes Care, Hypertension Support, Cholesterol Management"],
          ["Prescription Services", "Upload & Manage Prescriptions, Repeat Prescriptions, Specialist Referrals"],
          ["Medical Devices & Equipment", "Blood Pressure Monitors, Glucometers, Thermometers, Pulse Oximeters"],
          ["Personal Protective & Hygiene", "Hand Sanitisers, Masks, Gloves, Disinfectants"],
        ]
      ),
      spacer(120),

      heading2("3.3 Product Detail Page (PDP) Requirements"),
      para("Every product detail page must meet the following information standard. Incomplete product data must be flagged in the admin panel and not published until complete."),
      spacer(80),

      bullet("Product Identity: Full product name, brand (Wellgo/Pediazone/other), SKU, barcode (EAN/UPC), and manufacturer."),
      bullet("Rich Media: Minimum 3 high-resolution product images (front, back, side/ingredient view). 360\u00B0 image viewer recommended for premium products."),
      bullet("Medical Information Panel:"),
      subbullet("Active Ingredients with amounts per serving/dose"),
      subbullet("Other / Inactive Ingredients"),
      subbullet("Recommended Dosage by age group where applicable"),
      subbullet("Contraindications and known Drug Interactions"),
      subbullet("Common Side Effects"),
      subbullet("Storage Instructions"),
      subbullet("Warnings (pregnancy, lactation, allergy alerts)"),
      bullet("Prescription Status Badge: OTC, Rx-Required, or Pharmacist Recommended."),
      bullet("Verified by Pharmacist badge where applicable."),
      bullet("Stock Status: In Stock, Low Stock (show count if \u226410 units), Out of Stock with restock notification signup."),
      bullet("Customer Reviews: Star rating, verified purchase badge, and ability to filter reviews by rating."),
      bullet("Related Products: \u201CFrequently Bought Together\u201D and \u201CPeople also viewed\u201D carousels."),
      bullet("Accordion Sections: Full product description, FAQs, and Manufacturer Information."),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 4
      // ══════════════════════════════════════════════════════════
      heading1("4. Operational Dashboard (Admin Panel)"),

      heading2("4.1 Role-Based Access Control (RBAC)"),
      spacer(80),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2400, 3480, 3480],
        rows: [
          new TableRow({ children: [
            ["Role", "Permissions", "Restrictions"].map((h, i) => new TableCell({
              borders: border(COLORS.primary), shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 120, right: 120 },
              width: { size: [2400, 3480, 3480][i], type: WidthType.DXA },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Arial", size: 22, bold: true, color: COLORS.white })] })],
            })),
          ]}),
          new TableRow({ children: [
            ["Super Admin", "Full access: all modules, user management, system configuration, financial reports.", "N/A \u2014 all access. Requires hardware MFA."].map((t, i) => new TableCell({
              borders: border("CCCCCC"), shading: { fill: COLORS.white, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: [2400, 3480, 3480][i], type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 20, color: COLORS.darkText })] })],
            })),
          ]}),
          new TableRow({ children: [
            ["Store Manager", "Inventory management, order processing, product catalog edits, customer service, content publishing.", "Cannot access pharmacist prescription queue or raw PHI. Cannot change user roles."].map((t, i) => new TableCell({
              borders: border("CCCCCC"), shading: { fill: COLORS.lightGray, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: [2400, 3480, 3480][i], type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 20, color: COLORS.darkText })] })],
            })),
          ]}),
          new TableRow({ children: [
            ["Licensed Pharmacist", "Prescription queue review, approval/rejection, pharmacist note entry, patient health profile read access, clinical product data editing.", "Cannot access financial reports, change pricing, or manage non-clinical product data."].map((t, i) => new TableCell({
              borders: border("CCCCCC"), shading: { fill: COLORS.white, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: [2400, 3480, 3480][i], type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 20, color: COLORS.darkText })] })],
            })),
          ]}),
          new TableRow({ children: [
            ["Customer Support Agent", "View order history, initiate refunds/returns per policy, manage customer inquiries, read-only product catalog.", "No access to prescriptions, PHI, or financial data. Cannot modify orders beyond refund/return initiation."].map((t, i) => new TableCell({
              borders: border("CCCCCC"), shading: { fill: COLORS.lightGray, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: [2400, 3480, 3480][i], type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 20, color: COLORS.darkText })] })],
            })),
          ]}),
          new TableRow({ children: [
            ["Content Editor", "Blog/Health Resource Centre authoring and publishing, brand page editing, promotional banner management.", "No access to orders, customer data, prescriptions, or inventory."].map((t, i) => new TableCell({
              borders: border("CCCCCC"), shading: { fill: COLORS.white, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: [2400, 3480, 3480][i], type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 20, color: COLORS.darkText })] })],
            })),
          ]}),
        ],
      }),
      spacer(120),

      heading2("4.2 Inventory Management"),
      bullet("Real-time stock level display with per-SKU count, reserved (in-cart) quantity, and available quantity."),
      bullet("Low-Stock Alerts: Configurable threshold per SKU. Email and dashboard alert to Store Manager when stock falls below threshold."),
      bullet("Out-of-Stock Actions: Automatic product visibility toggle (hide or show with \u2018Notify Me\u2019 button) when stock reaches zero."),
      bullet("Batch Import/Export: CSV and XLSX bulk product upload for Wellgo and Pediazone catalog ingestion. Template provided with field validation."),
      bullet("Supplier & Purchase Order Module: Track supplier lead times, raise purchase orders, and record goods receipt."),
      bullet("Expiry Date Tracking: For perishable health products, track expiry dates and generate pick-lists for near-expiry items."),
      spacer(80),

      heading2("4.3 Order Management"),
      bullet("Unified order queue with status pipeline: Pending Payment \u2192 Payment Confirmed \u2192 Prescription Pending (if applicable) \u2192 Processing \u2192 Dispatched \u2192 Delivered \u2192 Completed."),
      bullet("Manual order creation by Customer Support for phone orders."),
      bullet("Partial fulfilment: Ability to dispatch in-stock items and back-order the rest within a single order."),
      bullet("Integrated shipping label printing (DHL, FedEx, Royal Mail, or local carrier APIs)."),
      bullet("Return & Refund Workflow: Initiate return, generate returns label, track received status, trigger refund to original payment method."),
      spacer(80),

      heading2("4.4 Analytics & Reporting"),
      bullet("Sales Reports: Daily, weekly, monthly, and custom date range. Filterable by brand (Wellgo/Pediazone), category, and SKU."),
      bullet("Inventory Reports: Stock valuation, slow-moving items, stock turns, and reorder point recommendations."),
      bullet("Customer Reports: New vs. returning customer ratio, average order value, customer lifetime value (CLV)."),
      bullet("Search Analytics: Top search terms, zero-result searches (catalog gap report), conversion rate from search."),
      bullet("Prescription Metrics: Volume of prescriptions received, average review time, approval/rejection rate."),
      bullet("All reports exportable to CSV and PDF. Scheduled email delivery available for recurring reports."),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 5
      // ══════════════════════════════════════════════════════════
      heading1("5. Trust, Content & Engagement"),

      heading2("5.1 Trust-Building Page Structure"),
      heading3("5.1.1 About Us Page"),
      bullet("Pharmacy History & Mission: Story of the pharmacy, founding values, and commitment to patient care."),
      bullet("Pharmacist Team Profiles: Photo, full name, license number, and bio for each qualified pharmacist. This directly builds customer trust."),
      bullet("Accreditations & Certifications: Display pharmacy board certification, NABP Accreditation (or local equivalent), PCI-DSS compliance badge, and SSL seal."),
      bullet("Physical Location(s): Embedded Google Map with directions, store photos, and parking information."),
      spacer(80),
      heading3("5.1.2 Health Resource Centre (Blog)"),
      bullet("Medically Reviewed Articles: All health articles must carry a \u2018Reviewed by [Pharmacist Name]\u2019 attribution with a link to their profile."),
      bullet("Content Categories: Nutrition & Supplements, Baby & Child Health, Chronic Condition Management, Seasonal Health (Flu Season, Allergy Season), Drug Interaction Guides."),
      bullet("Brand-Linked Content: Articles that feature or recommend Wellgo or Pediazone products must link to product pages with a clear editorial vs. promotional distinction."),
      bullet("SEO Optimisation: Each article follows a proper heading structure (H1, H2, H3), includes schema.org Article markup, meta description, and open graph tags for social sharing."),
      bullet("Email Newsletter: Opt-in for Health Tips newsletter. Segmented by customer health interest profile."),
      spacer(80),
      heading3("5.1.3 Contact & Store Locator"),
      bullet("Contact page with: phone number (click-to-call on mobile), email contact form with category selector (Order Enquiry, Prescription Query, Product Information, General), and WhatsApp integration if applicable."),
      bullet("Store Locator with map and structured hours, including holiday hours management via the admin panel."),
      bullet("Footer on Every Page: Business name, address, phone, opening hours, license number, and Schema.org LocalBusiness JSON-LD markup. This is critical for local SEO and Google Business Profile alignment."),
      spacer(120),

      heading2("5.2 Customer Support Integration"),
      bullet("Live Chat Widget: Integrate a HIPAA-aware live chat tool (e.g., Intercom, Freshchat, or Zendesk) with business-hours routing and after-hours chatbot fallback."),
      bullet("Ask a Pharmacist: A dedicated secure messaging feature allowing customers to submit clinical questions to the pharmacist team, with a defined SLA (e.g., response within 4 hours during business hours)."),
      bullet("FAQ System: Searchable FAQ covering Ordering & Shipping, Prescriptions, Returns, Account & Privacy, and Brand-Specific questions for Wellgo and Pediazone. Admin-editable without a code deployment."),
      bullet("Proactive Order Notifications: Automated SMS and email at each order status change (Confirmed, Dispatched, Delivered)."),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 6
      // ══════════════════════════════════════════════════════════
      heading1("6. Technology Stack Recommendation"),

      para("The following stack is recommended for a secure, scalable, and maintainable pharmacy e-commerce platform. All selections prioritise HIPAA/GDPR compatibility, strong security communities, and proven production scale."),
      spacer(80),

      twoColTable(["Layer", "Recommended Technology & Rationale"],
        [
          ["Frontend Framework", "Next.js 14 (React) \u2014 Server-Side Rendering (SSR) for SEO, static generation for content pages, and React for interactive UI. TypeScript enforced throughout."],
          ["UI Component Library", "Tailwind CSS + shadcn/ui \u2014 Utility-first CSS for rapid, consistent design. Custom clinical design tokens (brand colors, typography) applied globally."],
          ["Backend / API", "Node.js with Express or NestJS \u2014 NestJS preferred for its modular architecture, built-in TypeScript, and decorator-based RBAC support."],
          ["Database (Primary)", "PostgreSQL (managed via AWS RDS or Supabase) \u2014 ACID-compliant relational database. All PHI fields encrypted at the column level using pgcrypto."],
          ["Database (Search)", "Elasticsearch or Algolia \u2014 Powers the smart search bar. Algolia recommended for simpler implementation; Elasticsearch for on-premise data residency requirements."],
          ["File Storage", "AWS S3 (or Google Cloud Storage) with server-side encryption (SSE-KMS). Presigned URLs with short expiry for prescription image access."],
          ["CDN", "Cloudflare \u2014 Global CDN, WAF, DDoS protection, and bot management. Cloudflare Images for optimised product photo delivery."],
          ["Payment Gateway", "Stripe (PCI-DSS Level 1) \u2014 Stripe Elements for hosted card fields (no raw card data touches the server). Stripe Billing for subscription refill orders."],
          ["Authentication", "NextAuth.js (frontend) + JWT + Refresh Token rotation (backend). Auth0 or Cognito as a managed identity provider for enterprise deployments."],
          ["Email / SMS", "AWS SES for transactional email. Twilio for SMS notifications and OTP delivery."],
          ["CMS (Health Blog)", "Contentful or Sanity.io (headless CMS) \u2014 Allows content editors and pharmacists to publish articles without code deployments."],
          ["Monitoring & Logging", "Datadog or AWS CloudWatch for application performance monitoring. Sentry for error tracking. Audit logs stored in an immutable append-only log (e.g., AWS CloudTrail)."],
          ["CI/CD", "GitHub Actions for automated testing, linting, security scanning (Snyk), and deployment pipelines to staging and production environments."],
          ["Hosting", "AWS (preferred for HIPAA BAA availability) or Google Cloud. Container-based deployment via ECS (Fargate) or GKE for horizontal scalability."],
        ]
      ),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 7
      // ══════════════════════════════════════════════════════════
      heading1("7. Phased Development Roadmap"),

      para("The project is structured across four phases. Phases 1 and 2 constitute the Minimum Viable Product (MVP). Phases 3 and 4 deliver advanced features, automation, and scale optimisation."),
      spacer(120),

      phaseCard("Phase 1", "Foundation & MVP Launch", "Duration: Months 1\u20133",
        [
          "Domain, hosting, SSL certificate, and CI/CD pipeline setup",
          "Core authentication (signup, login, email verification, password reset)",
          "Product catalog (Wellgo & Pediazone) with categories and full PDP",
          "Smart search (Algolia integration) with brand and category filters",
          "Shopping cart, checkout, and Stripe payment integration",
          "Basic prescription upload workflow (upload, store, email notification)",
          "Customer account dashboard (orders, addresses, password management)",
          "Homepage, About Us, Contact, and Store Locator pages",
          "Footer with NAP and Schema.org LocalBusiness markup",
          "Admin panel: inventory view, order management, basic user management",
          "WCAG 2.1 AA accessibility audit and remediation",
          "OWASP penetration test before go-live",
        ]
      ),
      spacer(120),

      phaseCard("Phase 2", "Trust & Compliance Layer", "Duration: Months 4\u20135",
        [
          "Pharmacist dashboard with prescription review queue and RBAC",
          "Verified by Pharmacist badge system on products and orders",
          "Prescription status tracker in customer My Account",
          "Full HIPAA / GDPR compliance review and legal sign-off",
          "MFA implementation for admin, pharmacist, and optional customer MFA",
          "Automated refill reminder emails and SMS (Twilio)",
          "Customer review and rating system on PDPs",
          "Health Resource Centre (Blog) with Contentful CMS integration",
          "Live chat integration (Intercom or Zendesk)",
          "Low-stock alerts, expiry tracking, and inventory reports",
          "PCI-DSS SAQ completion and ASV scan",
        ]
      ),
      spacer(120),

      phaseCard("Phase 3", "Engagement & Growth Features", "Duration: Months 6\u20138",
        [
          "Wellgo and Pediazone dedicated brand landing pages with editorial content",
          "Advanced search: symptom-based search, \u2018Find by Condition\u2019 guided flow",
          "Loyalty points and rewards program",
          "Personalised product recommendations (based on order history and health profile)",
          "Email marketing automation: welcome series, abandonment flows, refill campaigns",
          "Promotional engine: discount codes, bundle pricing, brand-specific promotions",
          "Ask a Pharmacist secure messaging feature",
          "Order tracking integration with shipping carriers (DHL/FedEx/local API)",
          "Mobile app (React Native \u2014 iOS & Android) using shared API",
          "Apple Pay and Google Pay checkout",
        ]
      ),
      spacer(120),

      phaseCard("Phase 4", "Scale, Analytics & Automation", "Duration: Months 9\u201312",
        [
          "Advanced analytics dashboard: CLV, cohort analysis, brand performance",
          "Search analytics and zero-result catalog gap automation alerts",
          "AI-powered drug interaction checker on cart/checkout (third-party API integration)",
          "Subscription / Auto-Refill orders managed via Stripe Billing",
          "Multi-location inventory management (if pharmacy expands to additional branches)",
          "ERP / POS integration (e.g., RxNT, QS/1, or local pharmacy management system)",
          "Performance optimisation: Core Web Vitals scores \u226590 across all pages",
          "Third-party marketplace integration (Google Shopping, Meta Catalogue) for Wellgo/Pediazone",
          "Annual security penetration test and HIPAA re-assessment",
        ]
      ),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // SECTION 8 — Local SEO
      // ══════════════════════════════════════════════════════════
      heading1("8. Local SEO & Schema.org Implementation"),

      para("Local SEO is a critical acquisition channel for a retail pharmacy. The following requirements must be implemented from Phase 1 and maintained as a standing operational process."),
      spacer(80),

      heading2("8.1 On-Site Requirements"),
      bullet("NAP Consistency: Business Name, Address, and Phone number must be identical across the website footer, Google Business Profile, and all directory listings."),
      bullet("Schema.org JSON-LD: A LocalBusiness schema block (extending Pharmacy type) must be injected in the <head> of every page. This must include name, address (PostalAddress), telephone, openingHoursSpecification, geo coordinates, url, and logo."),
      bullet("Google Business Profile: Claim and fully complete the GBP listing with product categories (Pharmacy, Health Supplement Store), photos, Q&A, and regular post updates."),
      bullet("Location Landing Page: If the pharmacy has multiple branches, each location requires a unique URL, unique content, and its own Schema.org block."),
      spacer(80),

      infoBox("Sample Schema.org JSON-LD (Pharmacy)", [
        "{",
        "  \"@context\": \"https://schema.org\",",
        "  \"@type\": \"Pharmacy\",",
        "  \"name\": \"RetailRx Pharmacy\",",
        "  \"address\": {",
        "    \"@type\": \"PostalAddress\",",
        "    \"streetAddress\": \"123 Health Street\",",
        "    \"addressLocality\": \"Your City\",",
        "    \"postalCode\": \"XX000\",",
        "    \"addressCountry\": \"XX\"",
        "  },",
        "  \"telephone\": \"+1-800-000-0000\",",
        "  \"openingHoursSpecification\": [",
        "    { \"@type\": \"OpeningHoursSpecification\",",
        "      \"dayOfWeek\": [\"Monday\",\"Tuesday\",\"Wednesday\",\"Thursday\",\"Friday\"],",
        "      \"opens\": \"08:00\", \"closes\": \"19:00\" }",
        "  ],",
        "  \"url\": \"https://www.retailrx.com\",",
        "  \"logo\": \"https://www.retailrx.com/logo.png\"",
        "}",
      ], "F4F6F7"),
      spacer(80),
      heading2("8.2 Ongoing SEO Operations"),
      bullet("Monthly review of Google Search Console for crawl errors, Core Web Vitals issues, and new ranking opportunities."),
      bullet("Quarterly content audit of the Health Resource Centre to update, expand, or retire articles based on search performance."),
      bullet("Structured review solicitation post-purchase to build Google rating volume."),

      pageBreak(),

      // ══════════════════════════════════════════════════════════
      // APPENDIX
      // ══════════════════════════════════════════════════════════
      heading1("Appendix: Key Metrics & Success Criteria"),

      twoColTable(["KPI", "Target (12 months post-launch)"],
        [
          ["Site Uptime", "99.9% (measured monthly)"],
          ["Page Load Speed (LCP)", "< 2.5 seconds on mobile (Core Web Vitals \u2018Good\u2019)"],
          ["Search Response Time", "< 200ms for autocomplete, < 500ms for full results"],
          ["Prescription Review SLA", "95% of prescriptions reviewed within 4 business hours"],
          ["Cart Abandonment Rate", "< 65% (industry average ~70\u201375%)"],
          ["Customer Account Adoption", "> 60% of repeat purchases from logged-in accounts"],
          ["Security Incidents", "Zero critical/high severity incidents. Annual pen test with no Critical findings."],
          ["WCAG Accessibility Score", "AA compliance on all pages. Zero critical accessibility failures."],
          ["Google Local Ranking", "Top 3 results for \u201Cpharmacy near me\u201D in target service area"],
          ["Net Promoter Score (NPS)", "> 45 (collected via post-delivery survey)"],
        ]
      ),

      spacer(200),
      para("Document Control:", { bold: true }),
      twoColTable(["Field", "Value"],
        [
          ["Document Version", "1.0"],
          ["Date", "June 2026"],
          ["Next Review Date", "December 2026"],
          ["Document Owner", "Head of Digital / CTO"],
          ["Approval Required From", "CEO, Head of Pharmacy Operations, Legal Counsel, CISO"],
          ["Classification", "Confidential \u2014 Internal Use Only"],
        ]
      ),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/RetailRx_Pharmacy_PRD.docx", buffer);
  console.log("Done.");
}).catch(err => { console.error(err); process.exit(1); });

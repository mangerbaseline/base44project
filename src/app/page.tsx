"use client";

import { useRef, useState, useEffect } from "react";
import Image from "next/image";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from "docx";
import { saveAs } from "file-saver";

const screenshots = [
  { src: "/image%20(59).png", caption: "Dashboard Overview – Revenue & Analytics" },
  { src: "/image%20(60).png", caption: "Jobs & Service Management Interface" },
  { src: "/image%20(4).jfif", caption: "Installation Module & Job Tracking" },
  { src: "/image%20(5).jfif", caption: "Commission & Invoice Management" },
];

export default function Home() {
  const reportRef = useRef<HTMLDivElement>(null);
  const [exporting, setExporting] = useState<string | null>(null);
  const [lightboxImg, setLightboxImg] = useState<string | null>(null);

  const exportPDF = async () => {
    if (!reportRef.current) return;
    setExporting("pdf");
    try {
      const canvas = await html2canvas(reportRef.current, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: "#ffffff",
      });
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF("p", "mm", "a4");
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
      let heightLeft = pdfHeight;
      let position = 0;

      pdf.addImage(imgData, "PNG", 0, position, pdfWidth, pdfHeight);
      heightLeft -= pdf.internal.pageSize.getHeight();

      while (heightLeft > 0) {
        position = heightLeft - pdfHeight;
        pdf.addPage();
        pdf.addImage(imgData, "PNG", 0, position, pdfWidth, pdfHeight);
        heightLeft -= pdf.internal.pageSize.getHeight();
      }
      pdf.save("Aircon_Project_Report.pdf");
    } catch (err) {
      console.error("PDF export error:", err);
    } finally {
      setExporting(null);
    }
  };

  const exportDOCX = async () => {
    setExporting("docx");
    try {
      const doc = new Document({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({
                text: "Aircon · Project Report",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
              }),
              new Paragraph({
                text: "Base44 · Installation & Services Management",
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
              }),
              new Paragraph({
                text: "1. Executive Summary",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({
                text: "This report details the development of an Aircon Installation & Services Management System built on the Base44 platform. The system automates the extraction and management of customer data, job sheets, and financial records from Google Sheets into a structured, user-friendly web application. It features role-based access control (Admin, Accounting, Service, Install), real-time dashboards with revenue analytics, and efficient data handling with skeleton loading and pagination. The project streamlines operations by separating Installation and Service jobs, managing service contracts, and tracking invoices—all while leveraging Base44's built-in authentication and role management.",
                spacing: { after: 200 },
              }),
              new Paragraph({
                text: "2. Introduction & Objective",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({
                text: "Background: Managing aircon operations manually through Google Sheets is prone to data duplication, errors, and lack of real-time visibility. This project bridges the gap by integrating Google Sheets with Base44 to create a centralized management dashboard.",
                spacing: { after: 100 },
              }),
              new Paragraph({
                children: [new TextRun({ text: "Objectives:", bold: true })],
                spacing: { before: 100 },
              }),
              new Paragraph({ text: "1. Auto-extract & sync data from Google Sheets to Base44." }),
              new Paragraph({ text: "2. Dynamically separate Installation & Service jobs." }),
              new Paragraph({ text: "3. Role-specific dashboards with revenue, GST, job analytics." }),
              new Paragraph({ text: "4. Skeleton loading + pagination for large datasets." }),
              new Paragraph({ text: "5. Manage service contracts via keyword-based conditions.", spacing: { after: 200 } }),
              new Paragraph({
                text: "3. Technology Stack",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "Base44, Google Sheets API, HTML5/CSS3/JS, Workflow Automation, Pagination (10/page), Skeleton Loading" }),
              new Paragraph({ text: "One-command deploy: base44 deploy", spacing: { after: 200 } }),
              new Paragraph({
                text: "4. System Architecture & Workflow",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "High-Level: [Google Sheets] ↔ [Base44 Integration Layer] ↔ [Base44 Database] ↔ [Frontend UI]" }),
              new Paragraph({ text: "Data Extraction, Data Segregation, Auto-Creation, Manual Operations, User Interaction.", spacing: { after: 200 } }),
              new Paragraph({
                text: "5. Authentication & Role Assignment",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "Base44 provides a fixed, uneditable login/signup system. Users access Base44 login → credentials validated against internal DB → session management handles auth." }),
              new Paragraph({ text: "Roles: Admin, Accounting, Service, Install", spacing: { after: 200 } }),
              new Paragraph({
                text: "6. Integration",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "Google Sheets Integration: Real-time extraction, syncing of customer/job/financial data." }),
              new Paragraph({ text: "Base44 Platform: Built-in DB, native UI components, workflow automation." }),
              new Paragraph({ text: "Future scope: Payment/Email gateways.", spacing: { after: 200 } }),
              new Paragraph({
                text: "7. Features & Modules",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "Dashboard - First page after login. Shows: Total Revenue, GST, Customers, Upcoming Dates, Overdue Invoices." }),
              new Paragraph({ text: "Installation - Displays all installation jobs extracted from Google Sheets." }),
              new Paragraph({ text: "Jobs - Shows three tables: Installation Jobs, Services, and Follow-ups." }),
              new Paragraph({ text: "Commission - Shows commission details — paid and pending." }),
              new Paragraph({ text: "Invoices - Displays extracted invoice data from Google Sheets." }),
              new Paragraph({ text: "Service Contracts - Extracts and manages maintenance services (AMC, Annual, Quarterly)." }),
              new Paragraph({ text: "Teams - Lists all team members (extracted + manually added)." }),
              new Paragraph({ text: "Charts/Analytics - Visual representation of job completion status.", spacing: { after: 200 } }),
              new Paragraph({
                text: "8. Data Handling & Performance Optimizations",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "Skeleton Loading: Placeholder UI while data is fetched." }),
              new Paragraph({ text: "Pagination & Scrolling: 10 rows per page + scrollbar." }),
              new Paragraph({ text: "Conditional Extraction Logic: Keywords separate jobs and contracts.", spacing: { after: 200 } }),
              new Paragraph({
                text: "9. Role-Based Access Control",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "Admin - All pages (Dashboard, Installation, Jobs, Commission, Invoices, Service Contracts, Teams, Charts)" }),
              new Paragraph({ text: "Accounting - Total Revenue Dashboard, Commission Page, Job Tables" }),
              new Paragraph({ text: "Service - Dashboard (service details), Service Job Sheets, Service Contract Page, Teams Page" }),
              new Paragraph({ text: "Install - Installation Page, Install Analytics, Teams Page", spacing: { after: 200 } }),
              new Paragraph({
                text: "10. Production Ready in Base44",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300 },
              }),
              new Paragraph({ text: "One-Command Deployment: base44 deploy — deploy backend (entities, functions, auth) and frontend site." }),
              new Paragraph({ text: "Integrated Service: Base44 handles data management, authentication, server-side logic all in one place. Includes automatic HTTPS and support for custom domains.", spacing: { after: 200 } }),
              new Paragraph({
                text: "Aircon Installation & Services · Built on Base44 | June 2026 | 4 roles · 10+ modules",
                alignment: AlignmentType.CENTER,
                spacing: { before: 400 },
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, "Aircon_Project_Report.docx");
    } catch (err) {
      console.error("DOCX export error:", err);
    } finally {
      setExporting(null);
    }
  };

  return (
    <>
      {/* Export Toolbar */}
      <div className="max-w-[1440px] mx-auto mb-4 flex flex-wrap items-center justify-between gap-3 px-2">
        <div className="text-sm text-gray-500">
          <i className="fas fa-file-alt mr-2"></i>
          Project Report · Export as DOCX or PDF
        </div>
        <div className="flex gap-3">
          <button
            onClick={exportDOCX}
            disabled={exporting === "docx"}
            className="flex items-center gap-2 px-5 py-2.5 rounded-full text-sm font-semibold transition-all duration-200 bg-blue-600 hover:bg-blue-700 text-white shadow-md hover:shadow-lg disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <i className="fas fa-file-word"></i>
            {exporting === "docx" ? "Exporting..." : "Export DOCX"}
          </button>
          <button
            onClick={exportPDF}
            disabled={exporting === "pdf"}
            className="flex items-center gap-2 px-5 py-2.5 rounded-full text-sm font-semibold transition-all duration-200 bg-red-600 hover:bg-red-700 text-white shadow-md hover:shadow-lg disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <i className="fas fa-file-pdf"></i>
            {exporting === "pdf" ? "Exporting..." : "Export PDF"}
          </button>
        </div>
      </div>

      {/* Report Content */}
      <div ref={reportRef} className="report-wrap" style={{ fontFamily: "'Inter', sans-serif" }}>
        {/* HEADER */}
        <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <h1><i className="fas fa-snowflake" style={{ color: "#2563eb", marginRight: "12px" }}></i> Aircon · Project Report</h1>
            <div className="subhead"><i className="fas fa-circle" style={{ color: "#22c55e", fontSize: "0.6rem", marginRight: "8px" }}></i> Base44 · Installation & Services Management</div>
          </div>
          <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            <span className="badge-role"><i className="fas fa-check-circle"></i> v2.0</span>
            <span className="badge-role" style={{ background: "#d1fae5", color: "#0b5e42" }}><i className="fas fa-database"></i> Google Sheets + Base44</span>
          </div>
        </div>

        {/* 1. EXECUTIVE SUMMARY */}
        <div className="card" style={{ background: "linear-gradient(135deg, #f0f7ff 0%, #ffffff 100%)", borderLeft: "5px solid #2563eb", marginTop: "1.5rem" }}>
          <div className="card-header"><i className="fas fa-bolt"></i> 1. Executive Summary</div>
          <p style={{ fontWeight: 400, color: "#1e293b" }}>This report details the development of an <strong>Aircon Installation & Services Management System</strong> built on the Base44 platform. The system automates the extraction and management of customer data, job sheets, and financial records from Google Sheets into a structured, user-friendly web application. It features <strong>role-based access control</strong> (Admin, Accounting, Service, Install), real-time dashboards with revenue analytics, and efficient data handling with skeleton loading and pagination. The project streamlines operations by separating Installation and Service jobs, managing service contracts, and tracking invoices—all while leveraging Base44's built-in authentication and role management.</p>
        </div>

        {/* 2. INTRODUCTION & OBJECTIVE + 3. TECH STACK */}
        <div className="grid-2">
          <div className="card">
            <div className="card-header"><i className="fas fa-bullseye"></i> 2. Introduction & Objective</div>
            <p><strong>Background:</strong> Managing aircon operations manually through Google Sheets is prone to data duplication, errors, and lack of real-time visibility. This project bridges the gap by integrating Google Sheets with Base44 to create a centralized management dashboard.</p>
            <div style={{ marginTop: "0.8rem" }}><strong>Objectives:</strong></div>
            <ul style={{ listStyle: "none", paddingLeft: 0, fontSize: "0.92rem" }}>
              <li><i className="fas fa-check-circle" style={{ color: "#22c55e" }}></i> 1. Auto-extract & sync data from Google Sheets to Base44.</li>
              <li><i className="fas fa-check-circle" style={{ color: "#22c55e" }}></i> 2. Dynamically separate Installation & Service jobs.</li>
              <li><i className="fas fa-check-circle" style={{ color: "#22c55e" }}></i> 3. Role-specific dashboards with revenue, GST, job analytics.</li>
              <li><i className="fas fa-check-circle" style={{ color: "#22c55e" }}></i> 4. Skeleton loading + pagination for large datasets.</li>
              <li><i className="fas fa-check-circle" style={{ color: "#22c55e" }}></i> 5. Manage service contracts via keyword-based conditions.</li>
            </ul>
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-cubes"></i> 3. Technology Stack</div>
            <div><span className="chip"><i className="fas fa-server"></i> Base44</span> <span className="chip"><i className="fab fa-google"></i> Sheets API</span> <span className="chip"><i className="fas fa-code"></i> HTML5/CSS3/JS</span> <span className="chip"><i className="fas fa-cog"></i> Workflow Automation</span> <span className="chip"><i className="fas fa-table"></i> Pagination (10/page)</span> <span className="chip"><i className="fas fa-skeleton"></i> Skeleton Loading</span></div>
            <div style={{ marginTop: "0.8rem", background: "#f1f5f9", padding: "0.5rem 1rem", borderRadius: "60px", display: "inline-block" }}><i className="fas fa-arrow-right" style={{ color: "#2563eb" }}></i> One‑command deploy: <code style={{ background: "#e2e8f0", padding: "0.1rem 0.6rem", borderRadius: "20px" }}>base44 deploy</code></div>
          </div>
        </div>

        {/* 4. SYSTEM ARCHITECTURE & WORKFLOW */}
        <div className="card" style={{ background: "#f8fcff" }}>
          <div className="card-header"><i className="fas fa-sitemap"></i> 4. System Architecture & Workflow</div>
          <p><strong>High-Level:</strong> [Google Sheets] ↔ [Base44 Integration Layer] ↔ [Base44 Database] ↔ [Frontend UI]</p>
          <div className="flow-row">
            <span className="flow-step"><i className="fas fa-table"></i> Google Sheets</span>
            <span className="flow-arrow"><i className="fas fa-arrow-right"></i></span>
            <span className="flow-step"><i className="fas fa-cloud-upload-alt"></i> Base44 Integration</span>
            <span className="flow-arrow"><i className="fas fa-arrow-right"></i></span>
            <span className="flow-step"><i className="fas fa-database"></i> Base44 DB</span>
            <span className="flow-arrow"><i className="fas fa-arrow-right"></i></span>
            <span className="flow-step"><i className="fas fa-desktop"></i> Frontend UI</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.8rem 1.5rem", marginTop: "0.8rem", fontSize: "0.9rem" }}>
            <div><i className="fas fa-arrow-right" style={{ color: "#2563eb" }}></i> <strong>Data Extraction:</strong> connects to Google Sheets, extracts fields: Name, Appointed Date, Service Dates, Phone, Job Type, Description, Payment Type, Payment Confirmation, Invoice Number, etc.</div>
            <div><i className="fas fa-code-branch" style={{ color: "#2563eb" }}></i> <strong>Data Segregation:</strong> Installation jobs (keywords), Service jobs (filters), Service Contracts (AMC, Annual, Quarterly).</div>
            <div><i className="fas fa-user-plus" style={{ color: "#2563eb" }}></i> <strong>Auto-Creation:</strong> Customers, Jobs (Install/Service tables), Team members extracted from Sheets.</div>
            <div><i className="fas fa-hand-paper" style={{ color: "#2563eb" }}></i> <strong>Manual Operations:</strong> Users can manually add team members, jobs, customers directly in Base44.</div>
            <div><i className="fas fa-sign-in-alt" style={{ color: "#2563eb" }}></i> <strong>User Interaction:</strong> Login → role-specific dashboard. Skeleton loading + pagination (10 rows per page with scrollbar).</div>
          </div>
        </div>

        {/* 5. AUTH & ROLE + 6. INTEGRATION */}
        <div className="grid-2">
          <div className="card">
            <div className="card-header"><i className="fas fa-lock"></i> 5. Authentication & Login</div>
            <p><i className="fas fa-check-circle" style={{ color: "#22c55e" }}></i> Base44 provides a fixed, uneditable login/signup system.</p>
            <p><strong>How Login Works:</strong> Users access Base44 login → credentials validated against internal DB → session management handles auth.</p>
            <p><strong>Role Assignment (Inbuilt):</strong> Admin creates/manages roles via Base44 Dashboard. System checks role and grants page access.</p>
            <div style={{ marginTop: "0.6rem", background: "#f1f5f9", padding: "0.4rem 1rem", borderRadius: "40px", display: "inline-block" }}><i className="fas fa-user-tag"></i> Roles: Admin, Accounting, Service, Install</div>
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-plug"></i> 6. Integration</div>
            <ul style={{ listStyle: "none", paddingLeft: 0 }}>
              <li><i className="fab fa-google" style={{ color: "#ea4335" }}></i> <strong>Google Sheets Integration:</strong> Real-time extraction, syncing of customer/job/financial data, auto-update on modification.</li>
              <li><i className="fas fa-database" style={{ color: "#2563eb" }}></i> <strong>Base44 Platform:</strong> Built-in DB, native UI components (tables, charts, dashboards), workflow automation for conditional extraction.</li>
              <li><i className="fas fa-clock" style={{ color: "#94a3b8" }}></i> <strong>Future scope:</strong> Payment/Email gateways (not integrated at this stage).</li>
            </ul>
          </div>
        </div>

        {/* 7. FEATURES & MODULES */}
        <h2 style={{ marginTop: "2rem", display: "flex", gap: "0.5rem" }}><i className="fas fa-th-large" style={{ color: "#2563eb" }}></i> 7. Features & Modules</h2>
        <div className="grid-2">
          <div className="card">
            <div className="card-header"><i className="fas fa-tachometer-alt"></i> Dashboard</div>
            <strong>First page after login.</strong> Shows: Total Revenue (Install + Service), Total GST, Total Customers, Upcoming Service/Install Dates, Overdue Invoices. <span className="chip">overview stats</span>
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-tools"></i> Installation</div>
            Displays all installation jobs extracted from Google Sheets. View, filter, manually add jobs.
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-briefcase"></i> Jobs</div>
            Shows three tables: Installation Jobs, Services, and Follow-ups. Manage job statuses.
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-coins"></i> Commission</div>
            Shows commission details — how much paid and how much pending. <span className="chip">Track payments</span>
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-file-invoice"></i> Invoices</div>
            Displays extracted invoice data from Google Sheets. Allows generation of new invoices. <span className="chip">Generate & manage</span>
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-file-signature"></i> Service Contracts</div>
            Extracts and manages maintenance services using keyword-based conditions (AMC, Annual, Quarterly). <span className="chip">View & manage contracts</span>
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-user-friends"></i> Teams</div>
            Lists all team members (extracted + manually added). Add/edit team members.
          </div>
          <div className="card">
            <div className="card-header"><i className="fas fa-chart-bar"></i> Charts/Analytics</div>
            Visual representation of job completion status (Completed, Pending, In-Progress). <span className="chip">Analyze performance</span>
          </div>
        </div>

        {/* 8. DATA HANDLING & PERFORMANCE */}
        <div className="card" style={{ background: "#f2f8ff" }}>
          <div className="card-header"><i className="fas fa-database"></i> 8. Data Handling & Performance Optimizations</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: "1.5rem" }}>
            <div><i className="fas fa-circle-notch fa-spin" style={{ color: "#2563eb" }}></i> <strong>Skeleton Loading:</strong> Placeholder UI while data is fetched from Google Sheets/Base44. Prevents blank screens, improves perceived performance.</div>
            <div><i className="fas fa-chevron-circle-down"></i> <strong>Pagination & Scrolling:</strong> 10 rows per page + scrollbar. Smooth rendering even with large datasets.</div>
            <div><i className="fas fa-filter"></i> <strong>Conditional Extraction Logic:</strong> Keywords separate Installation vs. Service jobs, and regular jobs vs. Maintenance contracts. Reduces manual sorting and errors.</div>
          </div>
          {/* mini table demo */}
          <div style={{ marginTop: "1rem", background: "white", borderRadius: "1.5rem", padding: "0.8rem 1rem" }}>
            <div style={{ display: "flex", gap: "2rem", fontWeight: 600, color: "#1e293b", borderBottom: "1px solid #dde5ed", paddingBottom: "0.5rem" }}>
              <span style={{ width: "30%" }}>Customer</span><span style={{ width: "25%" }}>Job</span><span style={{ width: "20%" }}>Date</span><span style={{ width: "15%" }}>Status</span>
            </div>
            <div style={{ display: "flex", gap: "2rem", padding: "0.5rem 0", borderBottom: "1px solid #edf2f7" }}><span style={{ width: "30%" }}>A. Sharma</span><span style={{ width: "25%" }}>Install 5kW</span><span style={{ width: "20%" }}>12/06/2026</span><span style={{ width: "15%" }}><span className="chip" style={{ background: "#dcfce7" }}>Complete</span></span></div>
            <div style={{ display: "flex", gap: "2rem", padding: "0.5rem 0", borderBottom: "1px solid #edf2f7" }}><span style={{ width: "30%" }}>B. Kaur</span><span style={{ width: "25%" }}>AMC Service</span><span style={{ width: "20%" }}>18/06/2026</span><span style={{ width: "15%" }}><span className="chip" style={{ background: "#fef9c3" }}>Pending</span></span></div>
            <div style={{ display: "flex", gap: "2rem", padding: "0.5rem 0" }}><span style={{ width: "30%" }}>C. Mehta</span><span style={{ width: "25%" }}>Quarterly</span><span style={{ width: "20%" }}>24/06/2026</span><span style={{ width: "15%" }}><span className="chip" style={{ background: "#dbeafe" }}>In-progress</span></span></div>
            <div className="pagination-demo"><span><i className="fas fa-chevron-left"></i> 1-10 of 42</span><span><i className="fas fa-chevron-right"></i></span><span><i className="fas fa-arrow-up"></i> scroll</span></div>
          </div>
        </div>

        {/* 9. RBAC + 10. PRODUCTION READY */}
        <div className="grid-2">
          <div className="card">
            <div className="card-header"><i className="fas fa-users-cog"></i> 9. Role-Based Access Control</div>
            <table>
              <tbody>
                <tr><th>Role</th><th>Pages / Modules Accessible</th></tr>
                <tr><td><span className="role-tag role-admin">Admin</span></td><td>All pages (Dashboard, Installation, Jobs, Commission, Invoices, Service Contracts, Teams, Charts)</td></tr>
                <tr><td><span className="role-tag role-accounting">Accounting</span></td><td>Total Revenue Dashboard, Commission Page, Job Tables (Install & Service)</td></tr>
                <tr><td><span className="role-tag role-service">Service</span></td><td>Dashboard (service details), Service Job Sheets, Service Contract Page, Teams Page</td></tr>
                <tr><td><span className="role-tag role-install">Install</span></td><td>Installation Page, Install Analytics, Teams Page</td></tr>
              </tbody>
            </table>
          </div>
          <div className="card" style={{ background: "#f1f9f0", borderLeft: "4px solid #16a34a" }}>
            <div className="card-header"><i className="fas fa-rocket"></i> 10. Production Ready in Base44</div>
            <p><i className="fas fa-check-circle" style={{ color: "#16a34a" }}></i> <strong>One-Command Deployment:</strong> <code style={{ background: "#e2e8f0", padding: "0.2rem 0.8rem", borderRadius: "40px" }}>base44 deploy</code> — deploy backend (entities, functions, auth) and frontend site.</p>
            <p><i className="fas fa-cloud" style={{ color: "#2563eb" }}></i> <strong>Integrated Service:</strong> Base44 handles data management, authentication, server-side logic all in one place. Includes automatic HTTPS and support for custom domains.</p>
          </div>
        </div>

        {/* 11. PROJECT PORTFOLIO – SCREENSHOTS */}
        <h2 style={{ marginTop: "2rem", display: "flex", gap: "0.5rem" }}><i className="fas fa-images" style={{ color: "#2563eb" }}></i> 11. Project Portfolio – Screenshots</h2>
        <div className="gallery-grid">
          {screenshots.map((img, idx) => (
            <div
              key={idx}
              className="gallery-item"
              onClick={() => setLightboxImg(img.src)}
              style={{ cursor: "pointer" }}
            >
              <img src={img.src} alt={img.caption} />
              <div className="gallery-caption">
                <i className="fas fa-image"></i> {img.caption}
              </div>
            </div>
          ))}
        </div>

        {/* Lightbox */}
        {lightboxImg && (
          <>
            <div className="lightbox-overlay" onClick={() => setLightboxImg(null)}>
              <img src={lightboxImg} alt="Screenshot preview" />
            </div>
            <div className="lightbox-close" onClick={() => setLightboxImg(null)}>
              <i className="fas fa-times"></i>
            </div>
          </>
        )}

        {/* FOOTER */}
        <div style={{ marginTop: "2.5rem", textAlign: "center", borderTop: "1px solid #e2e8f0", paddingTop: "1.5rem", color: "#64748b", fontSize: "0.9rem" }}>
          <i className="fas fa-snowflake" style={{ color: "#2563eb" }}></i> Aircon Installation & Services · Built on Base44  <span style={{ margin: "0 8px" }}>|</span>  <i className="fas fa-calendar-alt"></i> June 2026  <span style={{ margin: "0 8px" }}>|</span>  <i className="fas fa-users"></i> 4 roles · 10+ modules
        </div>
      </div>
    </>
  );
}
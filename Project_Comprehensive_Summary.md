# Project BD MAIN 2026: Comprehensive Summary & Scope

This document provides a consolidated overview of the project, its evolution through recent conversations, the technical constraints of the environment, and the current development scope.

---

## 1. Executive Summary
**Project Name:** BD MAIN 2026  
**Objective:** To build a robust, high-performance lead management and business development ecosystem. This includes NPI data processing, a specialized high-density dialer application, and automated synchronization with Google Apps Script, all operating within a highly restricted Windows environment.

---

## 2. Project Scope & Core Components

### **A. NPI Search & Data Hub**
- **Purpose:** Querying and processing National Provider Identifier (NPI) data from the NPPES dataset.
- **Features:** Local search engine powered by `nppes.db`, metadata management, and pipeline processing for large dataset updates (e.g., April 2026 dataset).
- **Interface:** Functional web-based hubs (`ben_npi_search.html`, `npi_tool_hub.html`) for interactive data exploration.

### **B. Tactical Dialer Application**
- **Purpose:** A high-density communication interface for efficient lead outreach.
- **Key Characteristics:** 
    - Optimized for "Docked View" (30% width).
    - "Tactical/Industrial" aesthetic with minimal vertical space usage.
    - Dynamic disposition menus and real-time duplication alerts.
- **Frontend:** Vanilla HTML/CSS/JS with focus on micro-animations and layout resilience.

### **C. Lead Management & External Integration**
- **Purpose:** Preparing leads for outbound campaigns.
- **Formatters:** Custom tools to export data in formats compatible with platforms like Apollo IO (CSV mapping for names, titles, LinkedIn URLs, etc.).

### **D. Automation & Sync Pipeline**
- **Purpose:** Bridging local development with Google Apps Script.
- **Infrastructure:** Using a customized `clasp` implementation to push/pull code despite system restrictions.

---

## 3. Conversation History & Evolution

| Conversation ID | Focus Area | Key Achievements |
| :--- | :--- | :--- |
| `3ef5c3bd` | **NPI API Restoration** | Audited backend routes; restored `/api/npi_search` and `/api/metadata` for NPPES 2026 processing. |
| `bcd7a36d` | **Apps Script Sync** | Reorganized project hierarchy; updated `.clasp.json` and sync scripts to support modular layouts. |
| `45dfe673` | **Apollo IO Formatting** | Implemented lead export formatting for Apollo IO imports. |
| `3c020ea0` / `8c342da6` | **Dialer Optimization** | Redesigned dialer for ultra-dense docked views; fixed layout regressions and CSS overflow issues. |
| `9083a8be` / `e97874e0` | **Dialer UX/UI** | Condensed duplication alerts; tightened spacing for high-density information display. |
| `053b4e16` | **Agent Capabilities** | Installed `planning-with-files` skill for persistent task management across sessions. |

---

## 4. Environment & Constraint Summary
*The following content is integrated from the project's constraint documentation.*

### **System Constraints (The "Sandbox")**
- **OS:** Windows (User: ben.arthur).
- **Shell:** PowerShell via Antigravity.
- **Blocked:** `cmd.exe`, `.cmd`, `.bat`, and `.ps1` execution. Standard `npm`/`npx` triggers are blocked.
- **Privileges:** Non-administrative; portable installations only.

### **Software Architecture**
- **Node.js:** Portable v24.14.1 (`C:\Users\ben.arthur\node-v24.14.1-win-x64\`).
- **Primary Tool:** `@google/clasp`.
- **Working Directory:** `C:\Users\ben.arthur\Desktop\BD MAIN 2026`.

### **Technical Implementation (The "Golden Path")**
To bypass system blocks, the absolute path of `node.exe` is used to call the compiled JavaScript entry point directly:

**Path:** `C:\Users\ben.arthur\node-v24.14.1-win-x64\node_modules\@google\clasp\build\src\index.js`

**Usage Pattern:**
```powershell
# Mandatory structure for clasp commands
& "node.exe" "path/to/clasp/index.js" [command] --no-localhost
```

**Convenience Alias:**
```powershell
function clasp { & "C:\Users\ben.arthur\node-v24.14.1-win-x64\node.exe" "C:\Users\ben.arthur\node-v24.14.1-win-x64\node_modules\@google\clasp\build\src\index.js" @args }
```

---

## 5. Instructions for Agents
When working in this repo, you must:
1. **Never** use global `npm`, `npx`, or `clasp` commands.
2. **Always** use the explicit node execution path for CLI tools.
3. **Respect** the "Tactical/Industrial" design system for UI components (high density, 75-85% scaling friendly).
4. **Maintain** the local `nppes.db` as the source of truth for NPI search logic.

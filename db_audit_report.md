# Database Audit Report: NPPES Dataset

**Date:** 2026-05-12
**Database Path:** `c:\Users\ben.arthur\Downloads\bens html\dialer_app\data\nppes.db`

---

## 📊 High-Level Summary
- **Total Records:** 175,723
- **Database Size:** 155.61 MB
- **Total Columns:** 330
- **Data Integrity:** 100% valid NPI formats (10-digit)

---

## 🌎 Top States Distribution
The dataset covers all 50 states, plus territories and some international locations.

| State | Record Count | % of Total |
| :--- | :--- | :--- |
| **Texas (TX)** | 16,331 | 9.29% |
| **California (CA)** | 15,430 | 8.78% |
| **Florida (FL)** | 15,293 | 8.70% |
| **New York (NY)** | 10,897 | 6.20% |
| **Illinois (IL)** | 6,716 | 3.82% |
| **Pennsylvania (PA)** | 6,375 | 3.63% |
| **North Carolina (NC)** | 6,154 | 3.50% |
| **Georgia (GA)** | 6,014 | 3.42% |
| **Michigan (MI)** | 5,810 | 3.31% |
| **Ohio (OH)** | 5,703 | 3.25% |

---

## 🏷️ Top Taxonomy Codes
The dataset is heavily weighted toward Durable Medical Equipment (DME) and Pharmacies.

| Taxonomy Code | Description (Est.) | Count | % of Total |
| :--- | :--- | :--- | :--- |
| **332B00000X** | Durable Medical Equipment & Supplies | 88,755 | 50.51% |
| **333600000X** | Pharmacy | 25,343 | 14.42% |
| **291U00000X** | Clinical Medical Laboratory | 19,399 | 11.04% |
| **3336C0003X** | Community/Retail Pharmacy | 10,294 | 5.86% |
| **335E00000X** | Ethically Managed Pharmacy | 7,983 | 4.54% |
| **3336L0003X** | Long Term Care Pharmacy | 7,191 | 4.09% |

---

## 🔍 Data Quality Audit
- **NPI Validation:** All records have a valid 10-digit NPI.
- **Phone Numbers:** Most records (approx. 98%) have a Practice Location phone number.
- **Authorized Officials:** Approximately 85% of records have an Authorized Official listed.
- **Sparsity:** Many "Other Provider Identifier" columns (1-20) are empty (>99% sparsity), which is normal for this schema.

---

## 🛠️ Schema Insights
The database follows the standard NPPES Data Dissemination schema with 330 columns. Key fields for dialing and lead generation (Names, Addresses, Phones, Taxonomy) are well-populated.

---

> [!NOTE]
> This audit was performed on a sample of the local SQLite database. For real-time updates or specific state subsets, please use the search interface in the Dialer App.

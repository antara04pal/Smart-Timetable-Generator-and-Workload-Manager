# Smart Timetable & Workload Manager — Academic Scheduling Intelligence with Macro Automation
**Enterprise-grade Excel workbook for managing university timetables, faculty workloads, and room utilisation — with VBA macro automation, smart data entry, automated data cleaning, and a unified analytics dashboard.**

Smart Timetable & Workload Manager enables academic administrators to capture, clean, and analyse institutional scheduling data end-to-end. It combines a guided Timetable Entry form, a six-technique data cleaning pipeline, faculty workload computation, room utilisation tracking, and a three-pivot analytics layer — all feeding into a single visual Dashboard.

Smart data entry &nbsp;·&nbsp; Six-step cleaning pipeline &nbsp;·&nbsp; Faculty workload computation &nbsp;·&nbsp; Room utilisation tracking &nbsp;·&nbsp; Conflict & overload detection &nbsp;·&nbsp; VBA macro automation

---

## Features

### Guided Timetable Entry Form
- Dedicated data entry sheet with field-by-field guidance for every column
- Auto-suggests the next available Slot ID (e.g. TS0109) to prevent ID gaps
- Helper columns for Subject Code and Faculty ID lookup to reduce manual errors
- Dropdown-ready fields for Day, Time Slot, Room Type, Department, and Semester

### Six-Technique Data Cleaning Pipeline
- Documented cleaning process with technique, problem, and Excel function for each step
- Handles duplicates, leading/trailing spaces, numbers stored as text, missing values, blank rows, and column sizing
- Produces two fully clean output sheets — one for timetable, one for faculty data

### Timetable Schedule Management
- Tracks 109+ class slots across departments, semesters, sections, and batches
- Records Room ID, Room Type, Batch year, and Is Lab flag alongside standard schedule fields
- Auto-computes **Conflict Flag** in the Clean Timetable sheet for each slot

### Faculty Workload Computation
- Stores raw faculty profiles for 100 faculty members across 10 departments
- Clean Faculty sheet auto-calculates Subjects Assigned, Theory hrs/wk, Lab hrs/wk, Total Contact Hours, and Workload Status
- Classifies each faculty as **Normal**, **Underloaded**, or **Overloaded** against their Max Allowed Hours

### Room Utilisation Tracking
- Tracks 40 rooms across five room types: Lab, Lecture Hall, Smart Classroom, Seminar Room
- Computes Utilised slots/wk and Utilisation % against Available slots/wk
- Flags each room as **Properly Utilised**, **Underutilised**, or **Overutilised**

### Triple Pivot Analytics
- **Timetable Pivot** — Subject-wise class count across 35 subjects (93 total classes)
- **Faculty Pivot** — Total Contact Hours per faculty member (236 hours aggregate)
- **Room Utilisation Pivot** — Max Utilisation % per room (peak: 150% overutilisation)

### Unified Dashboard
- Single-screen overview of timetable, faculty, and room data
- Visual KPIs: conflict slots, overload alerts, room utilisation breakdown (50 Properly Utilised / 45 Underutilised / 5 Overutilised)
- Designed for department heads and academic administrators

### Navigable Index with Back-Links
- Clickable links to all 10 sheets from the Index
- Every sheet includes a **Back to Index** link for fast two-way navigation

---

## Workbook Preview

### Dashboard Overview
> <img width="1869" height="599" alt="Dashboard_main" src="https://github.com/user-attachments/assets/5c5293ba-b68c-4909-bcdc-b4cf7e624483" />



### Timetable Entry Form
> <img width="1867" height="595" alt="Timetable_entry" src="https://github.com/user-attachments/assets/2981e4bc-356d-4f4c-a813-f2e8d84b5e8e" />


### Faculty Workload — Clean Data & Workload Status
> <img width="1860" height="576" alt="Clean_Faculty_Workload" src="https://github.com/user-attachments/assets/8f7d172d-7f28-43a0-a59f-0f6ffc0d6cb0" />


### Room Utilisation Analysis
> <img width="1147" height="585" alt="Room_Util_Analysis" src="https://github.com/user-attachments/assets/424b89d5-648e-468c-99bd-89a05f3c4404" />

### Duplicate Value Removals
> <img width="1382" height="607" alt="Timetable_duplicate_removal" src="https://github.com/user-attachments/assets/d1daf331-0010-42b4-939e-9f2eaa88fc58" />
> <img width="1345" height="606" alt="Faculty_duplicate_removal" src="https://github.com/user-attachments/assets/fc2d75be-730a-4d1f-9418-f2df0c8a0e39" />

### Taking helper column for blank row removal
> <img width="1428" height="612" alt="Helper" src="https://github.com/user-attachments/assets/93bffb44-1ae4-4117-b420-c29e7704d63f" />





---

## How It Works

### Data Flow
```
Timetable Entry Form
        ↓
  Raw Timetable Sheet          Raw Faculty Data Sheet       Raw Room Utilisation Sheet
        ↓                              ↓                              ↓
        └──────────── Data Cleaning Pipeline (6 Techniques) ─────────┘
        ↓                              ↓                              ↓
Clean Timetable Sheet     Clean Faculty Data & Workload    Modified Room Utilisation
(+ Conflict Flag)         (+ Workload Status Computed)     (+ Utilisation % Computed)
        ↓                              ↓                              ↓
        └───────────────── Pivot Tables (3 Pivot Views) ─────────────┘
                                       ↓
                              Unified Dashboard
```

### Step-by-Step Process
1. Administrator enters new slot data via the guided **Timetable Entry** form.
2. Raw data accumulates in **Raw Timetable**, **Raw Faculty Data**, and **Raw Room Utilisation** sheets.
3. The **Data Cleaning** sheet applies six documented techniques to produce clean datasets.
4. The **Clean Timetable** sheet resolves conflicts and flags each slot accordingly.
5. The **Clean Faculty Data & Workload** sheet auto-computes hours and workload status per faculty.
6. The **Modified Room Utilisation** sheet calculates utilisation % and status for each room.
7. Three **Pivot Table** views aggregate the clean data for subject, faculty, and room analysis.
8. The **Dashboard** surfaces all key metrics in one unified view.

---

## Workbook Architecture

```
Smart_Timetable_Workload_Macro.xlsm
│
├── Index                          → Navigation hub with links to all 10 sheets
├── Timetable Entry                → Guided data entry form with helper columns & auto Slot ID
├── Raw Timetable                  → Raw schedule data (109 records, 14 columns)
├── Clean Timetable                → Cleaned schedule data with Conflict Flag (15 columns)
├── Raw Faculty Data               → Raw faculty profiles (100 faculty, 9 columns)
├── Clean Faculty Data & Workload  → Cleaned profiles with computed workload metrics (14 columns)
├── Raw Room Utilisation           → Raw room data (40 rooms, allocated slots/wk)
├── Modified Room Util             → Cleaned room data with utilisation % and status
├── Pivot Tables                   → Three pivot views: Timetable, Faculty, Room Utilisation
└── Dashboard                      → Visual summary of all KPIs and metrics
```

---

## Data Schema

### Timetable Schedule (Raw & Clean)
| Column | Description |
|---|---|
| Slot ID | Unique identifier for each class slot (e.g. TS0001) |
| Day | Day of the week (Monday–Saturday) |
| Time Slot | Class timing (e.g. 9:00–10:00) |
| Department | CSE, ECE, ME, CE, EEE, IT, MCA, MBA |
| Subject Code | Department-prefixed course code |
| Subject Name | Course name (35 unique subjects) |
| Faculty ID | Assigned faculty identifier |
| Room ID | Assigned room (R101–R140) |
| Room Type | Lab / Lecture Hall / Smart Classroom / Seminar Room |
| Batch | Student intake year (e.g. 2021–2024) |
| Semester | Semester number (1–8) |
| Section | Section label (A–D) |
| Duration (hrs) | Class duration in hours |
| Is Lab | Whether the slot is a lab session (Yes/No) |
| Conflict Flag *(Clean only)* | Scheduling conflict status (Conflict / No Conflict) |

### Faculty Data (Raw & Clean)
| Column | Description |
|---|---|
| Faculty ID / Name | Unique identifier and full name |
| Department | Home department |
| Designation | Lecturer / Senior Lecturer / Assistant Professor / Associate Professor / Professor |
| Qualification | Highest qualification (e.g. M.Tech, Ph.D) |
| Experience (yrs) | Years of teaching experience |
| Semester / Academic Year | Applicable semester and year |
| Max Allowed Hours | Maximum permitted contact hours per week |
| Subjects Assigned *(Clean only)* | Number of subjects currently assigned |
| Theory hrs/wk *(Clean only)* | Computed weekly theory hours |
| Lab hrs/wk *(Clean only)* | Computed weekly lab hours |
| Total Contact Hours *(Clean only)* | Theory + Lab hours combined |
| Workload Status *(Clean only)* | Normal / Underloaded / Overloaded |

### Room Utilisation (Raw & Modified)
| Column | Description |
|---|---|
| Room ID | Room identifier (R101–R140) |
| Room Type | Lab / Lecture Hall / Smart Classroom / Seminar Room |
| Available / Allocated slots/wk | Total slots available per week |
| Utilised slots/wk *(Modified only)* | Actual slots used per week |
| Utilisation % *(Modified only)* | (Utilised / Available) × 100 |
| Utilisation Status *(Modified only)* | Properly Utilised / Underutilised / Overutilised |

---

## Data Cleaning Pipeline

| # | Technique | Problem Addressed | Excel Method |
|---|---|---|---|
| 1 | Remove Duplicates | Repeated rows inflate PivotTable totals | `COUNTIF`, Remove Duplicates |
| 2 | Trim Extra Spaces | Spaces cause LOOKUP mismatches | `TRIM`, `CLEAN` |
| 3 | Convert Numbers Stored as Text | Type mismatch errors in formulas | `VALUE` |
| 4 | Handle Missing Values | Blanks cause formula errors | `IF`, `ISBLANK`, `IFERROR` |
| 5 | Remove Blank Rows | Blank rows cause spill and formula errors | `ISBLANK`, delete rows |
| 6 | Adjust Column & Row Sizes | Improper sizes appear unprofessional | Format → AutoFit |

---

## Key Analytics Snapshot

| Metric | Value |
|---|---|
| Total class slots tracked | 109 |
| Total unique subjects | 35 |
| Total faculty members | 100 |
| Total rooms tracked | 40 |
| Aggregate faculty contact hours | 236 hrs/wk |
| Rooms properly utilised | 50% |
| Rooms underutilised | 45% |
| Rooms overutilised | 5% |
| Peak room utilisation | 150% (R126) |

---

## Use Cases
- University timetable administration and slot management
- Faculty workload compliance monitoring across semesters
- Room allocation optimisation and overutilisation detection
- Academic scheduling conflict identification
- Department-wise resource and faculty planning
- Semester-wise batch and section management
- Administrative reporting and executive dashboards

---

## Getting Started

### Requirements
- Microsoft Excel 2016 or later (`.xlsm` macro-enabled format)
- Macros must be **enabled** when prompted for full VBA automation

### Setup
1. **Download** `Smart_Timetable_Workload_Macro.xlsm`
2. **Open** in Microsoft Excel — click **Enable Macros** when prompted
3. **Navigate** using the **Index** sheet; use **Back to Index** links on each sheet to return
4. **Enter new slots** via the **Timetable Entry** sheet — the helper columns guide Subject Code and Faculty ID selection
5. **Apply cleaning** techniques documented in the **Data Cleaning** sheet to raw input sheets
6. **Refresh all pivot tables**: *Data → Refresh All*
7. **View** the Dashboard for an instant summary of conflicts, workload, and room status

---

## Macro Automation (VBA)
The `.xlsm` format enables VBA macros for workflow automation. Macros in this workbook support:
- Auto-incrementing Slot ID generation on the Timetable Entry form
- Navigation shortcut buttons (Back to Index links on every sheet)
- Data validation enforcement across entry fields

> Ensure macros are enabled in Excel's Trust Center settings for full functionality.

---

## Future Improvements
- Automated conflict detection with colour-highlighted cell alerts
- Dynamic department and semester filter slicers on the Dashboard
- Faculty availability calendar with leave tracking
- Semester-on-semester workload comparison pivot
- Export-to-PDF timetable report via VBA macro
- Room booking request workflow with approval tracking
- Email alert generation for overloaded faculty via Outlook VBA integration
- Student strength and section capacity tracking

---

## Developer
**Sayar Sekhar Ghosh**  
**Antara Pal**  

---

## Contributing
Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a pull request

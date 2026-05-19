#   Smart-Timetable-Generator-and-Workload-Manager(Academic Schedule & Faculty Workload Intelligence)
**Enterprise-grade Excel workbook for managing university timetables and faculty workloads — with automated data cleaning, pivot analytics, and an interactive dashboard.**

Timetable Manager enables academic administrators to track, clean, and analyse institutional scheduling data using structured Excel workflows. It covers end-to-end timetable management: from raw data ingestion to conflict detection, faculty overload flagging, and a unified visual dashboard.

Data-driven scheduling &nbsp;&nbsp; • Automated data cleaning &nbsp;&nbsp;  • Faculty workload tracking &nbsp;&nbsp;  • Conflict & overload detection

---

## Features

### Timetable Schedule Management
- Tracks 100+ class slots across departments, semesters, and sections
- Records Subject Code, Subject Name, Faculty ID, Day, Time Slot, and Duration
- Flags lab sessions and scheduling conflicts automatically

### Faculty Workload Tracking
- Monitors theory hours, lab hours, and total contact hours per faculty
- Compares against maximum allowed hours
- Flags overloaded faculty and logs extra duties (Exam Cell, Research Cell, etc.)

###  Data Cleaning Pipeline
- Dedicated sheet documenting the full data cleaning process
- Handles duplicate rows (e.g. repeated Slot IDs) that would inflate pivot totals
- Resolves missing values in Department and Day fields using `IF` and `COUNTBLANK`
- Produces clean, analysis-ready datasets for both timetable and faculty data

###  Pivot Table Analytics
- **Timetable Pivot** — Slice schedule data by Department, Day, Time Slot, Semester, and Section
- **Faculty Pivot** — Aggregate workload metrics by Department, Designation, and Workload Status

###  Interactive Dashboard
- Unified visual overview of timetable and faculty data
- Key KPIs at a glance: conflict flags, overload alerts, department-wise distribution
- Designed for administrators and department heads

###  Navigable Index Sheet
- Clickable links to all 8 sheets for fast navigation across the workbook

---

##  Workbook Preview

### 🔹 Dashboard Overview
> <img width="723" height="578" alt="Dashboard jpg" src="https://github.com/user-attachments/assets/7321afe0-46c1-4626-9a8c-bbaef7412f18" />


### 🔹 Timetable Schedule — Raw & Cleaned
> <img width="1857" height="577" alt="Raw Timetable jpg" src="https://github.com/user-attachments/assets/66054797-56c4-44bf-a5fd-a24fdb826f62" />

> <img width="1862" height="575" alt="Cleaned Timetable jpg" src="https://github.com/user-attachments/assets/704a1351-776a-4a47-8c5e-68686d8e68e4" />



### 🔹 Faculty Workload Analysis
> <img width="1642" height="537" alt="Faculty Workload jpg" src="https://github.com/user-attachments/assets/244b58ee-016d-4320-8086-2bdbc003b6ee" />




### 🔹 Removal of Duplicate Data
> <img width="1511" height="629" alt="image" src="https://github.com/user-attachments/assets/783af0e3-0e66-48dc-96b6-d8617e5379a9" />



---

##  How It Works

### Data Flow
```
Raw Timetable Data          Raw Faculty Workload Data
        ↓                              ↓
  Data Cleaning Sheet   ←  (Remove Duplicates, Handle Blanks)
        ↓                              ↓
Clean Timetable Sheet       Clean Faculty Workload Sheet
        ↓                              ↓
  Timetable Pivot              Faculty Pivot Table
        ↓                              ↓
              Unified Dashboard
```

### Step-by-Step Process
1. Raw schedule and faculty data are entered into their respective input sheets.
2. The **Data Cleaning** sheet documents and applies cleaning techniques.
3. Duplicate records are removed; missing Department/Day fields are resolved.
4. Clean data feeds into **Pivot Tables** for dynamic summarisation.
5. The **Dashboard** aggregates key metrics into a single decision-making view.

---

##  Workbook Architecture

```
Timetable_Manager.xlsm
│
├── Index                     → Navigation hub (links to all sheets)
├── Timetable_Schedule        → Raw timetable data (100+ records)
├── Faculty_Workload          → Raw faculty workload data
├── Data Cleaning             → Cleaning techniques & documentation
├── Clean_Timetable_Schedule  → Cleaned timetable data
├── Clean_Faculty_Workload    → Cleaned faculty workload data
├── Timetable_Pivot           → Pivot analysis of schedule data 
├── Faculty_Pivot             → Pivot analysis of faculty workload
└── Dashboard                 → Visual summary & KPIs
```

---

##  Data Schema

### Timetable Schedule
| Column | Description |
|---|---|
| Slot ID | Unique identifier for each class slot |
| Day | Day of the week |
| Time Slot | Class timing (e.g. 9:00–10:00) |
| Department | CSE, ECE, ME, CE, EEE, IT, MCA, MBA |
| Subject Code | Course code |
| Subject Name | Course name |
| Faculty ID | Assigned faculty identifier |
| Semester | Semester number (1–8) |
| Section | Section label (A–D) |
| Duration (hrs) | Class duration in hours |
| Is Lab | Whether the slot is a lab session (Yes/No) |
| Conflict Flag | Scheduling conflict indicator |

### Faculty Workload
| Column | Description |
|---|---|
| Faculty ID / Name | Faculty identifier and full name |
| Department | Home department |
| Designation | Assistant Professor / Senior Lecturer / Professor |
| Experience (yrs) | Years of experience |
| Subjects Assigned | Number of subjects assigned |
| Theory Hours/Week | Weekly theory teaching hours |
| Lab Hours/Week | Weekly lab hours |
| Total Contact Hours | Theory + Lab hours |
| Max Allowed Hours | Maximum permitted contact hours |
| Workload Status | Current load classification |
| Extra Duties | Additional responsibilities (Exam Cell, Research Cell, etc.) |
| Overload Flag | Flags faculty exceeding max hours |
| Semester / Academic Year | Applicable semester and year |

---

##  Data Cleaning Techniques

| Technique | Problem Addressed | Excel Method |
|---|---|---|
| Remove Duplicates | Repeated Slot IDs inflate pivot totals | *Remove Duplicates* (Data tab) |
| Handle Missing Values | Blank Department/Day fields cause formula errors | `IF`, `COUNTBLANK` |

---

##  Use Cases
- University timetable administration
- Faculty workload compliance monitoring
- Academic scheduling conflict detection
- Department-wise resource planning
- Semester planning and section management
- Administrative reporting and dashboards

---

##  Getting Started

### Requirements
- Microsoft Excel 2016 or later (`.xlsm` macro-enabled format)
- Macros must be enabled for full functionality

### Setup
1. **Download** `Timetable_Manager.xlsm`
2. **Open** in Microsoft Excel and **Enable Macros** when prompted
3. **Navigate** using the **Index** sheet to jump to any section
4. **Enter or update** raw data in `Timetable_Schedule` and `Faculty_Workload`
5. **Refresh** pivot tables: *Data → Refresh All*
6. **View** the Dashboard for an instant summary

---

##  Future Improvements
- Automated conflict detection with highlighted cell alerts
- Dynamic department filter on the Dashboard
- Semester-wise comparison pivot views
- Faculty availability calendar integration
- Export-to-PDF report generation via VBA macro
- Workload trend analysis across academic years
- Section-wise student strength tracking

---

##  Developers
**Antara Pal** 

**Sayar Sekhar Ghosh**  

---

##  Contributing
Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a pull request

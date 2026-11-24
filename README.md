# Python Programming (CSA08) - Assignment 4 (CO4)
## Robust Data Processing with Files, Exceptions, Modules and Packages

**Student Name:** V. Karthik Harish  
**Register Number:** 192511328  
**Department:** Computer Science and Engineering  
**Institution:** SIMATS Engineering, Saveetha Institute of Medical and Technical Sciences  
**Course Code:** CSA08 / CSA0807  
**Course Name:** Python Programming  
**Course Outcome:** CO4 – Design and implement Python programs using file handling, modules and packages to process and analyse data effectively (reading and writing files, error handling, modules and packages)  
**Submission Date:** 15/11/2026  
**Total Marks:** 50 Marks  
**GitHub Repository:** [https://github.com/KarthikHarish170260/CSA08-192511328-Assignment4](https://github.com/KarthikHarish170260/CSA08-192511328-Assignment4)

---

## 1. Environment and Technology Stack

- **Python Version:** `Python 3.14.3` (fully compatible with Python 3.10+)
- **Operating System:** Windows 11 (64-bit)
- **Primary Standard Libraries:** `os`, `sys`, `time`, `decimal`, `collections`, `importlib`, `subprocess`, `shutil`
- **Visualization & Screenshot Libraries:** `matplotlib >= 3.8.0`, `Pillow >= 10.0.0`, `python-docx >= 1.1.0`, `numpy >= 1.26.0`

Install required dependencies using:
```bash
pip install -r requirements.txt
```

---

## 2. Progressive Git Commit Log

Per assignment guidelines, work was committed progressively task-by-task:

| Commit Hash | Task / Milestone | Description |
| :--- | :--- | :--- |
| `d55e7da` | Setup | Initial project setup: directory layout, requirements.txt, .gitignore, README.md, and screenshot helper |
| `ac5bae6` | Task 01 | Mess bill processor diagnosis, with-statement line-streaming repair, and 50k benchmark |
| `b2e21c2` | Task 02 | Biometric attendance file processing design, SIMATS pseudocode, trace table & reports |
| `1f4f978` | Task 03 | Exception taxonomy, InvalidRecord, complete try statement, and misuse experiment |
| `c65274f` | Task 04 | Library returns forensics, defect reproduction, append mode repair, and cardinal rules |
| `7cd7c68` | Task 05 | Four code organisations comparison, change impact & import semantics |
| `c0520ab` | Task 06 | Cultural fest registration package, O(1) set deduplication & resilience |

---

## 3. Directory Structure

```text
CSA08-192511328-Assignment4/
│
├── Task01_File_Processing_Review/
│   ├── task1_mess_processor.py       # Diagnosed & repaired mess processor with Decimal precision
│   ├── output.txt                    # Execution log & 50k-line streaming vs buffer benchmark
│   └── output_screenshot.png         # High-resolution terminal execution screenshot
│
├── Task02_Requirement_To_Design/
│   ├── task2_attendance_design.py    # Biometric attendance engine (Single-pass accumulator)
│   ├── data/                         # 5 daily classroom biometric text files (600 lines)
│   ├── attendance_summary.csv        # Master CSV summary report
│   ├── defaulters_below_75.txt       # Defaulters report (< 75% threshold)
│   ├── damaged_records.log           # Rejected corrupted line log
│   ├── output.txt                    # Execution log & single-pass O(F*L) complexity proof
│   └── output_screenshot.png         # High-resolution terminal execution screenshot
│
├── Task03_Design_Error_Handling/
│   ├── task3_exceptions.py           # Exception hierarchy, InvalidRecord, parse_line & summarise
│   ├── output.txt                    # Verified output matching assignment specification
│   └── output_screenshot.png         # High-resolution terminal execution screenshot
│
├── Task04_Forensics_And_Repair/
│   ├── task4_library_forensics.py    # 3-day automated defect reproduction & append mode repair
│   ├── library_returns.txt           # Repaired multi-day append audit file
│   ├── library_errors.log            # Timestamped failure audit log
│   ├── output.txt                    # Execution log & leak forensics
│   └── output_screenshot.png         # High-resolution terminal execution screenshot
│
├── Task05_Four_Organisations/
│   ├── sample_data/                  # 5 days of multi-student attendance records
│   ├── org_a_single_script/          # Org A: Monolithic single script (attendance_monolithic.py)
│   ├── org_b_module_import/          # Org B: Single module (attendance_processor.py) + driver
│   ├── org_c_multi_module/           # Org C: Three flat modules (reader.py, stats.py, report.py)
│   ├── org_d_python_package/         # Org D: Python package (attendance/ with __init__.py)
│   ├── task5_compare_runner.py       # Master runner comparing all 4 & sys.modules caching
│   ├── output.txt                    # Execution log, change impact table & import benchmark
│   └── output_screenshot.png         # High-resolution terminal execution screenshot
│
├── Task06_Design_Challenge/
│   ├── department_inputs/            # 5 department text files + missing file test case
│   ├── fest_registration/            # Complete robust package
│   │   ├── __init__.py               # Package metadata and public exports
│   │   ├── validator.py              # Email validation using string methods only (no regex)
│   │   ├── parser.py                 # Fault-tolerant streaming line generator
│   │   ├── aggregator.py             # O(1) set-based deduplication & event routing engine
│   │   └── reporter.py               # Master CSV, individual event files & error log writer
│   ├── outputs/                      # Generated fest artifacts:
│   │   ├── master_registered_students.csv
│   │   ├── fest_summary_report.txt
│   │   ├── fest_processing_errors.log
│   │   └── event_rosters/            # Distinct text rosters for each scheduled event
│   ├── task6_fest_desk.py            # Master design challenge pipeline & O(N) vs O(N^2) benchmark
│   ├── output.txt                    # Execution log & 20k-record empirical speedup proof
│   └── output_screenshot.png         # High-resolution terminal execution screenshot
│
├── report_assets/                    # 6 publication-quality monochrome academic figures
│   ├── fig1_q1_memory_scaling.png
│   ├── fig2_q2_pass_complexity.png
│   ├── fig3_q3_eafp_vs_lbyl.png
│   ├── fig4_q4_descriptor_leakage.png
│   ├── fig5_q5_organisations_evaluation.png
│   └── fig6_q6_deduplication_scaling.png
│
├── run_all_tasks.py                  # Master suite verifying all 6 tasks sequentially
├── render_screenshot_helper.py       # Terminal rendering utility
├── push_task_by_task.ps1             # PowerShell progressive push script
├── push_task_by_task.bat             # Batch progressive push script
├── ASSIGNMENT_REPORT.md              # Complete analytical technical report in Markdown
├── CSA0807_Assignment4_192511328_Report.docx # Formal institutional Word report
├── requirements.txt                  # Python dependencies
└── README.md                         # Project documentation
```

---

## 4. Execution Commands and Summary of Results

### Master Test Runner
To execute and verify all 6 tasks in one unified run:
```bash
python run_all_tasks.py
```

### Individual Task Execution Commands

#### Task 1: File-Processing Design Review
```bash
python Task01_File_Processing_Review/task1_mess_processor.py
```
- **Result:** Successfully diagnosed 8 architectural deficiencies in the hostel mess script. Repaired using `with open(...)` line-by-line streaming, `decimal.Decimal` currency handling, and granular exception trapping. In the 50,000-line benchmark, streaming reduced peak memory consumption from **6.260 MB down to 0.150 MB ($41.8\times$ RAM savings)**.

#### Task 2: Requirement to File-Processing Design
```bash
python Task02_Requirement_To_Design/task2_attendance_design.py
```
- **Result:** Designed biometric attendance file processing from first principles. Authored SIMATS Algorithmic Pseudocode and trace tables. Generated `attendance_summary.csv`, `defaulters_below_75.txt`, and `damaged_records.log`. Single-pass streaming $O(F \cdot L)$ reduced total I/O comparisons by **$600\times$** compared to per-student multi-pass scans.

#### Task 3: Design the Error Handling
```bash
python Task03_Design_Error_Handling/task3_exceptions.py
```
- **Result:** Implemented the complete exception hierarchy and custom `InvalidRecord` class. Completed `parse_line()` and `summarise()`, yielding exact expected output: `(1, 1, [(3, 'expected 3 fields, found 2'), (4, "invalid status 'L'")])`. Proved that bare `except: pass` swallowed internal `NameError`, misreporting zero present students.

#### Task 4: File and Exception Forensics
```bash
python Task04_Forensics_And_Repair/task4_library_forensics.py
```
- **Result:** Identified that write mode `"w"` overwrote previous return records, missing `\n` merged titles, and write errors leaked file descriptors. Created 3-day automated test harness proving data loss. Repaired with append mode `"a"`, `with open(...)`, newline delimiters, and timestamped error logging.

#### Task 5: The Same Utility Organised Four Ways
```bash
python Task05_Four_Organisations/task5_compare_runner.py
```
- **Result:** Implemented 4 distinct codebase architectures: (a) monolithic script, (b) single module + main, (c) three flat modules, and (d) package `attendance/`. Added new requirement (3 consecutive absent days). Evaluated change impact and import syntax differences. Benchmarked `sys.modules` caching: cold load $3.03\text{ ms}$ vs cached warm import $0.0009\text{ ms}$ (**$3,362\times$ speedup**).

#### Task 6: Design Challenge – College Fest Registration Desk
```bash
python Task06_Design_Challenge/task6_fest_desk.py
```
- **Result:** Implemented modular package `fest_registration/` covering all 5 stages. Validated emails using string methods only (no regex). Achieved zero-crash resilience against missing files, damaged records, corrupted characters, and duplicates. Demonstrated empirical benchmark: Set $O(N)$ vs List $O(N^2)$ deduplication: at $N=20,000$, Set was **$396.9\times$ faster** ($1.94\text{ ms}$ vs $768.2\text{ ms}$).

---

## 5. Verification Viva Notes

1. **Why `with` context manager is mandatory:** Ensures `__enter__` and `__exit__` are invoked, guaranteeing OS file descriptor release even if `MemoryError`, `OSError`, or unexpected syntax exceptions occur.
2. **Why string methods over regex for email:** Eliminates polynomial backtracking vulnerabilities (Regular Expression Denial of Service - ReDoS) and removes heavy regex engine invocation overhead.
3. **Why $O(1)$ set deduplication:** In a high-volume registration desk, list membership checks require scanning all previously registered items ($N(N-1)/2$ comparisons), yielding quadratic $O(N^2)$ time explosion.
4. **Why packages over monoliths:** Limits blast radius. When the 3-day consecutive absence rule was added, only `stats.py` and `report.py` were modified; `reader.py` remained completely untouched, preventing regression.

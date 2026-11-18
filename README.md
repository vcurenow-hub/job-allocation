# 🎓 National Open University Job Allocation & Examination Marking Portal

A robust, intelligent Flask web application designed for managing examination paper marking jobs, automated AI data parsing/OCR correction, multi-session time routing (8 AM, 11 AM, 2 PM), live allocation tracking, 1-hour countdown timers, and admin user management.

---

## ✨ Key Features

1. **AI Intelligent Record Parser & Arranger:**
   - Supports importing `.txt`, `.docx`, `.csv`, and `.xlsx` files or pasting unorganized text directly into the Job Paste Chat Box.
   - Automatically heals OCR typos (e.g. letter `O` &rarr; digit `0`).
   - Auto-routes records into canonical **8:00 AM**, **11:00 AM**, and **2:00 PM** session tables.
   - Optional Gemini AI integration (`GEMINI_API_KEY`) for advanced extraction.
2. **Session Tables & Real-Time Allocation Control:**
   - Dedicated session tables for 8 AM, 11 AM, and 2 PM with live Open/Close allocation controls.
   - Individual table search bars, row editing, deleting, section clearing, and bulk record additions.
   - Ability to delete/clear all uploaded records across all sessions in one click.
3. **Priority Jobs & Random Allocation Algorithm:**
   - Admins can mark specific jobs as **⭐ Priority** for immediate allocation on demand.
   - Non-priority jobs are dynamically randomized and shuffled (`RANDOM()`) on every request to ensure fair distribution.
4. **Student Sibling Grouping ("Same Computer"):**
   - Automatically groups different courses for the same student matriculation number at the same session time and allocates them to the same user with a **"Same Computer"** station warning badge.
5. **User Workflow & 1-Hour Countdown Timers:**
   - Allocated job details are initially hidden. Users enter their **Computer IP Address** to open and view details.
   - Sequential start enforcement: users must click **START TIMER** to begin the live 1-hour countdown.
   - The **Submit & Score** form unlocks automatically during the final 10 minutes of the session.
   - Users can hold up to 8 active unsubmitted jobs at a time before new requests are restricted.
6. **Comprehensive Admin Controls:**
   - Real-time **All Jobs & Tracking** table with status color coding (Red = Unattended, Yellow = In Progress, Green = Finished).
   - User Management with active workloads, start times, and live duration tickers ("Active for X hrs Y mins").
   - Actionable notifications with direct action buttons linking straight to resolution pages.

---

## 🚀 Installation & Quick Start

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/your-username/job-allocation-portal.git
   cd job-allocation-portal
   ```

2. **Create and Activate a Virtual Environment:**
   ```bash
   python -m venv venv
   # On Windows:
   venv\Scripts\activate
   # On macOS/Linux:
   source venv/bin/activate
   ```

3. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Application:**
   ```python
   python app.py
   ```

5. **Access the Portal:**
   - Open your browser and go to `http://127.0.0.1:5000`
   - **Default Administrator Credentials:**
     - **Username / Email:** `admin` / `vcure.now@gmail.com`
     - **Password:** `AdminSecure2026!`

---

## 🛠️ Tech Stack

- **Backend:** Python, Flask, SQLite
- **Frontend:** Bootstrap 5, Bootstrap Icons, HTML5, JavaScript (Real-time search & live timers)
- **Document Processing:** `python-docx`, `openpyxl`

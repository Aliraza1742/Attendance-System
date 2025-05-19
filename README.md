# Campus Attendance System

This is a **Python-based Attendance Management System** using Excel as the backend storage (`Attendance.xlsx`) and the `openpyxl` library for Excel manipulation. It simulates a simple attendance system suitable for classroom or campus use.

---

## 📦 Features

- **Sign-up System**: Users must sign up before marking attendance.
- **Attendance Entry**: Records entry time and handles late entries.
- **Ban List**: Admin (Rector) can ban/unban students using a password-protected area.
- **Exit Logging**: Tracks when users leave campus and distinguishes early and late leavers.
- **Data Persistence**: All records are stored in `Attendance.xlsx`.

---

## 📁 File Structure

- **`Attendance.xlsx`**: Stores all attendance data.
- **`Sheet1`** (inside the workbook) uses the following columns:
  - **A**: Name
  - **B**: Roll No.
  - **C**: Entry Time
  - **E**: Sign-up Names
  - **F**: Sign-up Roll Numbers
  - **G**: Passwords
  - **I**: Ban List (Roll Nos.)
  - **J**: Late Students
  - **K**: Leave Before 5 PM
  - **L**: Leave After 5 PM
  - **M**: Still in Campus

---

## ⚙️ Setup

1. **Install Required Library**
   ```bash
   pip install openpyxl
🔍 Functionality Overview
-------------------------

### AppendData(list, startrow, columnnumber)

Appends data from the list into the specified Excel column starting from startrow, avoiding overwriting existing cells.

### ToRead(list, columnnumber)

Reads data from a specific column (starting from row 2) and appends it into the provided list.

### removedata(columnname)

Clears all data in a specified Excel column starting from row 2 (leaves headers intact).

👤 Class: Attendance
--------------------

This class encapsulates all logic related to sign-ups, attendance marking, campus exit, and administration.

### 📘 ReadAllData()

Populates all global class lists with current data from the Excel file.

### 🟢 Entry()

Allows an already signed-up student to mark their attendance. It:

*   Records entry time
    
*   Marks student as "in campus"
    
*   Flags them as "late" if they arrived after 9:00 AM
    
*   Rejects banned users
    

### 🟨 Signup()

Allows new users to sign up by providing name, roll number, and password. Stores the data in the respective columns.

### 🔎 Information(name)

Displays a user’s:

*   Name
    
*   Roll number
    
*   Entry time (if available)
    

### 🔴 Exit()

Logs when a student leaves the campus:

*   Before 5 PM: Adds to "Leave Before 5" list
    
*   After 5 PM: Adds to "Leave After 5" list
    
*   Removes them from "Still in Campus" list
    

### 🔐 banlist()

Password-protected access (Namal123) for Rector to:

*   View banned roll numbers
    
*   Add/remove roll numbers from ban list
    

🚀 Running the System
---------------------

Run the program:

```bash
   python attendance.py
```

You will see this menu:

``` bash
Campus Attendance System
1. Entry (Mark Attendance)
2. Signup
3. Get Information
4. Exit
5. Ban List (Rector Access)
0. Quit
```

Choose the options to interact with the system.

📝 Notes
--------

*   The header row is initialized only once. After first execution, you can comment out the code that writes headers (A1, B1, etc.).
    
*   The system is **console-based**, designed for local use or small-scale management.
    
*   Make sure Attendance.xlsx is not open in Excel when running the script to avoid file access issues.
    


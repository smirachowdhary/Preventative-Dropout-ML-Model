
# Preventative Dropout Early Warning System
Project Link: https://preventative-dropout-ml-model.vercel.app/

## Overview

The Preventative Dropout Early Warning System is a web-based application designed to help educators identify at-risk students through structured data analysis. By enabling early detection of potential risk factors such as attendance, academic performance, and behavioral indicators, the system supports timely intervention and improved student outcomes.

---

## Purpose

The primary objectives of this system are:

- To identify students at risk of dropping out at an early stage  
- To provide educators with data-driven insights  
- To support proactive intervention strategies  
- To organize and track student-specific observations and actions  

---

## Features

### Data Upload and Management
- Upload Excel files (`.xlsx`) containing student data  
- Support for multiple file uploads  
- Automatic parsing and structuring of student records  
- Duplicate file detection using file fingerprinting  
- Ability to delete uploaded datasets  

### Student Risk Analysis
- Evaluation of key indicators such as:
  - Attendance  
  - Academic performance  
  - Behavioral data  
- Classification of students into risk categories (e.g., low, moderate, high)  
- Structured display of student-level insights  

### Notes and Tracking
- Ability to add and manage notes for individual students  
- Organized tracking of interventions and observations  

### User Interface
- Dashboard-based layout for clarity and usability  
- Tab-based navigation for managing datasets and views  
- Structured presentation of data for ease of interpretation  

### Authentication
- Secure user authentication using Supabase  
- Email and password-based login system  
- Password reset functionality via email  

---

## Technology Stack

| Layer        | Technology                |
|--------------|--------------------------|
| Frontend     | Next.js (App Router)     |
| Backend      | Supabase (Auth, Database)|
| Deployment   | Vercel                   |
| Data Parsing | SheetJS (xlsx)           |
| Styling      | Tailwind CSS             |

---

## System Workflow

1. **Data Upload**  
   Users upload Excel files containing student data.

2. **Data Processing**  
   The system parses the uploaded files and extracts relevant fields.  
   Risk indicators are evaluated and used to classify students.

3. **Data Storage**  
   Data is maintained per user session, with optional persistence via Supabase.

4. **Visualization**  
   The dashboard displays student information, categorized risk levels, and notes.

5. **Intervention**  
   Educators can review insights, add notes, and take appropriate action.

---

## Authentication Flow (Supabase)

### Password Reset Example

```ts
await supabase.auth.resetPasswordForEmail(email, {
  redirectTo: `${window.location.origin}/login?mode=reset`,
});

# 📚 Library System for Junior High School

A web-based digital library system built for a junior high school (**Library Spentigs** — SMPN 39 Surabaya) using **Google Apps Script** and **Google Sheets** as the backend/database. This was my very first time building a project with Google Apps Script and Spreadsheets, developed iteratively with **Gemini** as my AI learning partner. While building it, I focused on practical end-to-end workflows—catalog browsing, barcode scanning, borrower accounts, book procurement requests, an school library chatbot, admin analytics, and Excel/PDF reporting.

## 📦 Technologies

* `Google Apps Script`
* `Google Sheets`
* `HTML / CSS / JavaScript`
* `Chart.js`
* `html5-qrcode & JsBarcode`
* `SheetJS (xlsx) & jsPDF`
* `Gemini (AI Pair-Programming & Learning Assistant)`

## ✨ Features

Here's what you can do with **Library System for Junior High School**:

* **Browse & Filter the Book Catalog:** Search books in real time and filter by category with clear, high-contrast UI labels so titles and stock counts are easy to read.
* **Scan & Generate Barcodes:** Use the device camera (`html5-qrcode`) to scan book barcodes during lookup or borrowing, and generate/print barcodes (`JsBarcode`) for books and student cards.
* **Unified Login (Admin & Students/Visitors):** Log in through a single entry point where role access is handled by account credentials—admins manage the system, while students/visitors access their personal borrowing history.
* **Personal Borrowing History:** Logged-in students can view their profile, digital library card with barcode, and active/past borrowed books tied to their own account.
* **Book Procurement Requests:** Students and visitors can submit requests for new books they want the library to acquire, while admins can review and update request statuses.
* **Library Chatbot Assistant:** Ask questions about book availability, opening hours, library rules, and information about **SMPN 39 Surabaya**'s library.
* **Admin Dashboard & Analytics:** View visual charts (`Chart.js`) summarizing circulation and categories, manage books, process borrowing and returns, and track procurement requests.
* **Bulk Account Import & Excel/PDF Exports:** Download Excel templates for students and teachers, bulk-import new accounts, and export data to **Excel (`.xlsx`)** or **PDF** across admin sections.

## ☕ The Process

I started from scratch with zero prior experience in Google Apps Script or using Google Sheets as a database backend, using **Gemini** to break down concepts, troubleshoot errors, and iterate quickly. First, I set up the core structure so the web app could read and write book and borrowing data in Google Sheets.

Early on, I ran into real bugs—like book stock not decreasing right after borrowing and the loading screen spinning endlessly when `Date` objects from Sheets weren't serialized properly for `google.script.run`. Debugging those issues with Gemini taught me how client-server communication works in Apps Script and how to make state updates reliable.

Next, I expanded the app based on real library needs: replacing a public loan lookup with a **Book Procurement Request** feature, adding a **personal borrower view** per account, integrating **barcode scanning and generation**, and building a **library chatbot** tailored to **SMPN 39 Surabaya**.

Finally, I built out the **Admin Dashboard** with charts, student/teacher Excel templates and bulk import, and **Excel/PDF export** buttons across admin features. Along the way, I documented what broke and why—because explaining how something works in your own words is one of the best ways to lock in what you learned.

## 📚 What I Learned

During this project, I picked up practical full-stack habits and a much clearer mental model of how data moves between a frontend UI, server scripts, and a spreadsheet database.

* **Working with Gemini as a First-Timer:**
  * **Faster Learning Loop:** Using Gemini helped me translate ideas into working Google Apps Script and UI code even when I didn't know the platform conventions beforehand, while still pushing me to read and test every change.
* **Google Apps Script & Google Sheets Backend:**
  * **Data Serialization & State:** I learned why raw spreadsheet `Date` values can break `google.script.run` responses and how normalizing types and forcing spreadsheet flushes prevents stale stock counts or stuck loading states.
  * **Spreadsheet Hygiene:** I learned to structure sheets (`Books`, `Borrowings`, `Users`, `Requests`, `Config`) cleanly and guard against formula injection or long ID numbers turning into scientific notation.
* **Role-Based Flows & Practical Utilities:**
  * **User & Admin Separation:** Separating student borrowing history, procurement requests, and admin circulation/account tools showed how role boundaries keep interfaces cleaner and safer.
  * **Reporting & Barcodes:** Integrating camera barcode scanning, chart visualizations, Excel imports/exports, and PDF generation turned a basic CRUD app into something much closer to a real school tool.
* **Overall Growth:**
  * Even as a beginner to Google Apps Script and Spreadsheets, building this project showed me how much you can ship by pairing curiosity, structured debugging, and AI assistance with clear documentation. Every bug I hit forced me to understand the system a little deeper rather than just copying code blindly.

## 💬 How Can It Be Improved?

* Add stronger authentication/session handling and password hashing instead of basic credential checks.
* Add due-date reminders, late return tracking, and fine calculation rules for junior high school circulation.
* Add book cover uploads and edit modals for existing book details and student profiles.
* Add automated tests or sample seed scripts so anyone testing the repo can spin up demo data effortlessly.
* Refine mobile responsiveness for barcode scanning and admin tables on smaller screens.

## 🚦 How to Build & Run the Project

Because this project runs on **Google Sheets + Google Apps Script**, you don't need a local Node server to run it—follow these steps to set it up from scratch:

1. **Create the Database Spreadsheet:** Go to Google Sheets, create a new spreadsheet (for example, `Database Library Spentigs`), and open **Extensions > Apps Script**.
2. **Add Backend & Frontend Code:** Paste your server code into `Code.gs` and create an HTML file named `Index` (`Index.html`) for the web interface, then save (`Ctrl + S` / `Cmd + S`).
3. **Initialize the Sheets:** In the Apps Script editor, select `setupDatabase` from the function dropdown and click **Run**, granting the requested Google permissions so the required tabs and starter records are created.
4. **Deploy as a Web App:** Click **Deploy > New deployment**, select **Web app**, set **Execute as** to **Me** and **Who has access** to **Anyone** (or your school/organization scope as needed), and click **Deploy**.
5. **Open & Test:** Open the generated web app URL in your browser to browse the catalog, test the chatbot and procurement requests, and log in as Admin or a sample student account. *(Whenever you update the code later, create a **New deployment** or update an existing deployment to a **New version** so your changes go live.)*

![Preview Demo Aplikasi](LibrarySystem.mp4)

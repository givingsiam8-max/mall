# Email Inventory & Campaign Manager

A high-performance email inventory and bulk outreach campaign management application built with **React 19**, **Vite**, **TypeScript**, **Express**, **Tailwind CSS**, and **SQLite (sql.js)**.

Features real-time Excel/CSV ingestion, automated contact deduplication, CAN-SPAM compliant suppression management, and background queue dispatching with live progress monitoring.

---

## ✨ Features

- **Multi-File Excel & CSV Ingestion**: Drag-and-drop `.xlsx`, `.xls`, and `.csv` files. Automatically scans all columns, normalizes addresses, and eliminates duplicates across historical imports.
- **Direct Gmail & SMTP Delivery**: Built-in support for Gmail (using 16-character Google App Passwords), custom SMTP, Resend, SendGrid, Amazon SES, and Mailgun.
- **Background Batch Processing**: Asynchronous delivery engine with customizable batch sizing (e.g. 50/min), delay throttling, and automatic retry on transient socket errors.
- **Suppression & CAN-SPAM Compliance**: Single-click unsubscribe tokens, manual suppression blacklist, and automatic exclusion of suppressed emails before any queue dispatch.
- **Live Campaign Analytics**: Real-time progress bars, delivery breakdown (Sent / Pending / Failed / Suppressed), and per-recipient message IDs from mail servers.
- **Dark & Light Mode**: Accessible, polished responsive user interface with system theme sync.

---

## 🚀 Quick Start (Local Run)

### 1. Prerequisites
- **Node.js** 18+ or 20+
- **npm** or **pnpm** / **yarn**

### 2. Installation
```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
npm install
```

### 3. Development Server
Start the full-stack server (Express + Vite frontend):
```bash
npm run dev
```
Open [http://localhost:3000](http://localhost:3000) in your browser.

### 4. Production Build
```bash
npm run build
npm start
```

---

## 📤 How to Publish to GitHub

You can publish this project directly to your GitHub account:

### Step 1: Create a new repository on GitHub
1. Go to [https://github.com/new](https://github.com/new).
2. Name your repository (e.g. `email-campaign-manager`).
3. Choose **Private** (recommended if you intend to store confidential lists) or **Public**.
4. Leave "Initialize with README" **unchecked** (we already have one).
5. Click **Create repository**.

### Step 2: Push your code
In your terminal inside this project folder:

```bash
# Initialize git repository (if not already initialized)
git init

# Stage all files (.gitignore automatically protects SQLite database and local secrets)
git add .

# Create your first commit
git commit -m "Initial commit: Email Inventory & Campaign Manager"

# Rename default branch to main
git branch -M main

# Link to your GitHub repository (replace with your URL from Step 1)
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/<YOUR_REPOSITORY_NAME>.git

# Push to GitHub
git push -u origin main
```

---

## 🔒 Security & Privacy Notice

- **Your credentials are protected**: `.gitignore` is pre-configured to exclude `data/`, `*.sqlite`, `.env`, and local credentials so passwords and email recipient databases are **never committed to your public GitHub repo**.
- When deploying to production servers (Render, Railway, DigitalOcean, VPS), enter your Gmail credentials or API keys directly in the **Settings** UI or as environment variables.

---

## ✉️ Email Provider Setup

### Using Gmail (Fastest Setup)
1. Turn ON **2-Step Verification** on your Google Account.
2. Go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
3. Type `Campaign Manager` as the App name and click **Generate**.
4. Copy the 16-character code (e.g. `twgb ocxi cnsb xwze`).
5. Open the app's **Settings & SMTP** tab:
   - Select **Gmail**.
   - Enter your Gmail address as Sender.
   - Paste the 16-character App Password.
   - Click **Save Configuration**.
6. Use the **Send Verification Test** tool at the bottom to verify connectivity!

---

## 📜 License
MIT License. Free to use, modify, and distribute for personal or commercial projects.

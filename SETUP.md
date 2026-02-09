# IJTA Roll Call App - Setup Guide

## What You Have

- `index.html` — The roll call app (PWA)
- `manifest.json` — PWA manifest for "Add to Home Screen"
- `google-apps-script.js` — Script that writes attendance to Google Sheets

## Step 1: Set Up the Google Apps Script

This is the "middleman" that lets the app write to your attendance sheet while keeping it view-only.

1. Open your **Attendance Google Sheet**: https://docs.google.com/spreadsheets/d/1ipQEh5KCRywBOin8GM4xjzvGh9iK1YWp8VD9BXGH_YA/edit
2. Go to **Extensions > Apps Script**
3. Delete any code already there
4. Open the file `google-apps-script.js` from this folder, copy ALL the code
5. Paste it into the Apps Script editor
6. Click the **floppy disk icon** (or Ctrl+S) to save
7. Click **Deploy > New deployment**
8. Click the **gear icon** next to "Select type" and choose **Web app**
9. Set **"Execute as"** to **Me** (your Google account)
10. Set **"Who has access"** to **Anyone**
11. Click **Deploy**
12. Google will ask you to authorize — click through and allow it
13. **Copy the Web App URL** — it looks like: `https://script.google.com/macros/s/XXXXX/exec`

## Step 2: Make Your Roster Sheet Viewable

1. Open your **Roster Google Sheet**: https://docs.google.com/spreadsheets/d/1DszkseXqMekH_erFHVEVELcKRI6UgPLexSZ9YtLYqBc/edit
2. Click **Share** (top right)
3. Under "General access", change to **Anyone with the link**
4. Set the role to **Viewer**
5. Click **Done**

## Step 3: Host the App

The app needs to be hosted somewhere so instructors can access it via a URL. The simplest free options:

### Option A: GitHub Pages (Recommended)
1. Create a free GitHub account if you don't have one
2. Create a new repository (e.g., "ijta-roll-call")
3. Upload `index.html` and `manifest.json`
4. Go to Settings > Pages > Source: "main" branch
5. Your app will be live at: `https://yourusername.github.io/ijta-roll-call/`

### Option B: Netlify Drop (Easiest, no account needed)
1. Go to https://app.netlify.com/drop
2. Drag and drop the folder containing `index.html` and `manifest.json`
3. Your app will be live at a random URL (you can customize it with a free account)

## Step 4: First Use

1. Open the app URL on your phone
2. Select a clinic
3. The first time you submit, it will ask for the **Google Apps Script URL** (from Step 1, #13)
4. Paste the URL and submit — it saves this so you only enter it once

## Step 5: Add to Home Screen

### iPhone/iPad:
1. Open the app in Safari
2. Tap the **Share** button (square with arrow)
3. Tap **Add to Home Screen**
4. Tap **Add**

### Android:
1. Open the app in Chrome
2. Tap the **three-dot menu**
3. Tap **Add to Home screen** or **Install app**

## Updating the Coach List or Clinics

These are hardcoded in `index.html` for simplicity. To update:

1. Open `index.html` in a text editor
2. Find the `COACHES` array and edit the names
3. Find the `CLINICS` array and edit the clinic names
4. Re-upload to your hosting provider

## How the Attendance Sheet Works

Each submission creates rows like:

| Date | Clinic | Coaches | Player Name |
|------|--------|---------|-------------|
| 02/07/2026 | Red Ball (Ages 8 and Under) | Joey Francis, J.C. Freeman | McMillian, William |
| 02/07/2026 | Red Ball (Ages 8 and Under) | | Wilson, Elle |
| 02/07/2026 | Red Ball (Ages 8 and Under) | | McMillian, Caroline |

- Coaches appear only on the first row of each session
- Only present players are recorded (absent = no row)
- Players are listed as "Last, First"

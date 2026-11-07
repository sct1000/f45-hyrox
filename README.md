# F45 Morgan Hill HYROX Hub - Training Sign-Up & Strava Sync

This Google Apps Script powers both:
1. **Training Session Sign-Up Web App** (`TRAINING SESSIONS` tab, GID `728305493`) — Mobile-friendly UI for registering/deregistering for group and time-slotted sessions (solo or doubles), non-member waitlist, and Coach session creation tools.
2. **Strava Run Accountability Auto-Sync** (`🏃🏻‍♀️Run Accountability!` tab, GID `1418914961`) — Automates syncing runs and weekend Hyrox workouts for **F45 Morgan Hill** members.

- **Live Sheet**: [HYROX INFO & TRAININGS new](https://docs.google.com/spreadsheets/d/1-45Ncf8zFqMCE_z5YPJ_1P48ct-HQdYEH_NUUr_-ZYA/edit)
- **Test Copy Sheet**: [TEST COPY - HYROX INFO & TRAININGS](https://docs.google.com/spreadsheets/d/1KZtKRFMQXrjFlcfSRrViLlp-48SOJIyzenV5ru58Ez4/edit)
- **Script Location**: Managed directly under your own Google account (`script.google.com`).

---

## 1. Strava API Application Setup

1. Go to [Strava API Settings](https://www.strava.com/settings/api).
2. Create an application:
   - **Application Name**: `F45 Morgan Hill Run Tracker`
   - **Category**: Club / Event / Social
   - **Authorization Callback Domain**: `script.google.com`
3. Copy your **Client ID** and **Client Secret**.

---

## 2. Google Apps Script Setup (Under Your Account)

1. Open [script.google.com](https://script.google.com) and open your `F45 Strava Run Accountability` project.
2. Replace all text in `Code.gs` with the entire content of [`Code.js`](file:///Users/steviet/Projects/Strava/Code.js).
3. Go to **Project Settings (⚙️)** > **Script Properties** and verify:
   - `SPREADSHEET_ID`: `1-45Ncf8zFqMCE_z5YPJ_1P48ct-HQdYEH_NUUr_-ZYA`
   - `TEST_SPREADSHEET_ID` *(optional)*: `1KZtKRFMQXrjFlcfSRrViLlp-48SOJIyzenV5ru58Ez4`
   - `STRAVA_CLIENT_ID`: Your Strava Client ID
   - `STRAVA_CLIENT_SECRET`: Your Strava Client Secret
   - `STRAVA_VERIFY_TOKEN`: Any secret passphrase (e.g. `f45_morgan_hill_strava_2026`)
4. Click **Save script properties**.

---

## 3. Deploy & Access Control (Making the Sheet Read-Only)

Because the Web App runs with **Execute as: `Me (your email)`**:
- You and Coach Jade can make the Google Sheet **View Only (Read-Only)** for everyone else (`Anyone with the link: Viewer`, with only you and Coach Jade as `Editors`).
- All registrations and deregistrations submitted through the Web App will still write cleanly to the sheet through your script's permissions!
- **Requiring a Google Account**:
  - In Google Apps Script **Deploy > Manage deployments** (or **New deployment**), set **Who has access** to **`Anyone with Google Account`**.
  - *(Note: Since Strava runs are synced via the 10-minute background trigger `setupAutoSyncTrigger`, setting the Web App to `Anyone with Google Account` works with both Training Sign-Up and Strava OAuth!)*

### URL Parameters Supported:
- `?page=training` *(default)*: Opens the **🏋️ Training Sign-Up** tab directly.
- `?page=strava`: Opens the **🏃 Strava Run Sync** opt-in tab directly.
- `?test=1`: Connects the Web App to the **Test Copy** spreadsheet (`1KZtKRFMQXrjFlcfSRrViLlp-48SOJIyzenV5ru58Ez4`) so you can test registering, deregistering, and Coach tools without touching the live sheet.

---

## 4. How the Training Sign-Up Works

- **Source of Truth**: Reads and writes directly to the `TRAINING SESSIONS` tab (`MEMBER SIGN UP` in Column G and `NON-MEMBER WAITLIST` in Column H).
- **Two Session Modes**:
  1. **Time-Slotted Sessions** (e.g. Simulation Days):
     - Automatically detected when `MEMBER SIGN UP` has time headers (`10:30 AM`, `10:38 AM`, etc.) and numbered spots (`1.`, `2.`).
     - Athletes can claim any open spot as **👤 Solo** or **👥 Doubles / Pair** (both take 1 slot).
     - Anyone can remove a signup by clicking `✕` next to the name and confirming.
  2. **Group Sessions** (no time slots):
     - Displays a numbered list of member sign-ups (`1. Name`, `2. Athlete A & Athlete B`).
- **Non-Member / Waitlist**:
  - Available on all sessions; maintains a clean numbered list in Column H (`NON-MEMBER WAITLIST`).
- **Coach Tools (`⚙️ Coach: Add Session` & `⚙️ Manage`)**:
  - Coach Jade can still add rows directly in the Google Sheet (and paste time-slot templates just like before), **or** click **⚙️ Coach: Add Session** in the Web App to insert a new row at the top of `TRAINING SESSIONS` and auto-generate time slots (e.g. starting at `10:30 AM`, every `8` minutes, `12` slots, `2` groups per slot) or a capped Group session (`MAX X SPOTS`).
  - On any upcoming session card, clicking **⚙️ Manage** lets a coach extend the number of slots (`+1` / `+2` / `+4` time slots, or `+2` / `+5` group spots) and reschedule a timed session's start time and minute gap while keeping existing sign-ups in their slot order.
- **Concurrency Safe**: Uses `LockService.getScriptLock()` so simultaneous mobile sign-ups never overwrite each other or overfill a time slot.

---

## 4. Register the Strava Webhook (Instant Sync)

1. In the Apps Script editor, open the function dropdown at the top (next to "Run" and "Debug").
2. Select `registerStravaWebhook` and click **Run**.
3. View the **Execution log**. You should see:
   ```json
   { "id": 123456, "resource_state": 2, "application_id": ..., "callback_url": "..." }
   ```
   *(Your webhook is now live! Whenever an opted-in athlete completes a run, Strava will instantly notify your script within seconds).*

---

## 5. Adding the Opt-In Link into the Sheet

1. Open the [Run Accountability! sheet](https://docs.google.com/spreadsheets/d/1oKlUXQ7qOEDI_kx-gkjPc8cZN6QTbx0A4AGZTUyEYFw/edit?gid=1418914961).
2. In Cell **A1** (which is currently empty above the names), add a formula or hyperlink:
   ```excel
   =HYPERLINK("YOUR_WEB_APP_URL", "👉 Opt-in with Strava")
   ```
3. When members click it:
   - It opens a clean webpage.
   - It pulls the roster directly from Column A (`Jade`, `Sandy`, `Steve T`, `Shane`, etc.).
   - They choose their name from the dropdown and click **Connect with Strava**.
   - After authorization, their account is paired, and their runs are backfilled immediately!

---

## 6. How it Handles Logic & Manual Edits

- **Sports Filter**: Only activities categorized as `Run` with distance ≥ 500m or duration ≥ 3 mins are counted (avoids accidental taps).
- **Date Matching**: Automatically matches the date of the run with the weekly headers (`Aug 31 - Sep 6`, `Sep 7-13`, `Sep 14 - 20`, `Sep 21-27`, and any newly added week columns).
- **Preserves Manual Entries**: Uses `Math.max(existingCellValue, stravaCount)`. If someone manually logged `4` (e.g. treadmill without GPS) and Strava has `3`, the `4` is preserved.

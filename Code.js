/**
 * F45 Morgan Hill Run Accountability - Robust Strava Auto-Sync Engine
 * 
 * Target Spreadsheet: https://docs.google.com/spreadsheets/d/1oKlUXQ7qOEDI_kx-gkjPc8cZN6QTbx0A4AGZTUyEYFw/edit
 * Sheet Tab: "Run Accountability!"
 * 
 * Handled Edge Cases:
 * 1. Table relocated anywhere on the sheet (auto-locates table header & name column).
 * 2. Inconsistent human week header formats ("Sep 28 - Oct 4", "9/28-10/4", "Sep 28-4", "September 28 to October 4", etc.).
 * 3. Automatic addition of new week columns when a new week arrives.
 * 4. Automatic addition of new member rows when a new runner opts in.
 * 5. Preservation of manual numbers (uses Math.max).
 */

const SETTINGS = {
  SHEET_TAB: 'Run Accountability!',
  TRAINING_TAB: 'TRAINING SESSIONS',
  TRAINING_TAB_GID: '728305493',
  RACE_WAVES_TAB: 'Hyrox Anaheim 2026 times',
  RACE_TRAVEL_TAB: 'Travel ✈️ Hyrox Anaheim 2026!! ',
  TEST_SPREADSHEET_ID: '1KZtKRFMQXrjFlcfSRrViLlp-48SOJIyzenV5ru58Ez4',
  CONFIG_TAB: '_Strava_Athletes_Config',
  STRAVA_AUTH_URL: 'https://www.strava.com/oauth/token',
  STRAVA_API_BASE: 'https://www.strava.com/api/v3',
  MIN_RUN_DISTANCE_METERS: 500, // 0.5 km minimum
  MIN_RUN_DURATION_SECONDS: 180, // 3 minutes minimum
  DEFAULT_COST_TEXT: 'For unlimited and 9x per month members, the cost is $10. For 4x per month members and all non-members, the cost is $30. Class pack purchasers are able to use their pre-purchased classes towards the Hyrox sessions. Your card on file with F45 will be charged when we close the registration window (the day prior to the training session)'
};

// ==========================================
// 1. WEB APP (doGet: Training Sign-Up & Strava Onboarding)
// ==========================================

function getWebAppUrl() {
  const props = PropertiesService.getScriptProperties();
  const customUrl = props.getProperty('WEB_APP_URL');
  if (customUrl) return customUrl.trim();
  try {
    return ScriptApp.getServiceUrl();
  } catch (err) {
    return 'https://script.google.com/macros/s/AKfycbztqv4-3b2LVJ8DHuDgK07gkhIAnxacD4F9u7FqOU3wVN5kx3LqItg966GslZgiOgBqnA/exec';
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiAction(action, payload) {
  try {
    const p = payload || {};
    const useTestSheet = Boolean(p.useTestSheet);
    let state;
    if (action === 'getState') {
      state = getTrainingAppState(useTestSheet, p.coachEmail);
    } else if (action === 'registerSignup') {
      state = registerTrainingSignup(p);
    } else if (action === 'deregisterSignup') {
      state = deregisterTrainingSignup(p);
    } else if (action === 'coachCreateSession') {
      state = coachCreateTrainingSession(p);
    } else if (action === 'coachUpdateSession') {
      state = coachUpdateTrainingSession(p);
    } else if (action === 'coachDeleteSession') {
      state = coachDeleteTrainingSession(p);
    } else if (action === 'saveRaceWaveEntry') {
      state = saveRaceWaveEntry(p);
    } else if (action === 'deleteRaceWaveEntry') {
      state = deleteRaceWaveEntry(p);
    } else if (action === 'saveRaceTravelEntry') {
      state = saveRaceTravelEntry(p);
    } else if (action === 'deleteRaceTravelEntry') {
      state = deleteRaceTravelEntry(p);
    } else if (action === 'verifyCoachEmail') {
      const res = verifyCoachEmail(p.email, useTestSheet);
      return jsonResponse({ ok: true, isCoach: res.isCoach, state: res.state });
    } else {
      throw new Error('Unknown API action: ' + action);
    }
    return jsonResponse({ ok: true, state: state });
  } catch (err) {
    return jsonResponse({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}

/**
 * Handles Web App GET requests:
 * 1. Strava Webhook subscription verification (?hub.mode=subscribe&...)
 * 2. Strava OAuth callback (?code=...&state=...)
 * 3. JSON API (?api=1&action=getState&...)
 * 4. Unified Web App with tabs for Training Sign-Up (?page=training) & Strava Sync (?page=strava)
 */
function doGet(e) {
  const params = (e && e.parameter) || {};

  // Strava Webhook verification challenge
  if (params['hub.mode'] === 'subscribe' && params['hub.challenge']) {
    const props = PropertiesService.getScriptProperties();
    const verifyToken = props.getProperty('STRAVA_VERIFY_TOKEN');
    if (!verifyToken || params['hub.verify_token'] === verifyToken) {
      return ContentService.createTextOutput(
        JSON.stringify({ 'hub.challenge': params['hub.challenge'] })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput('Invalid verify token').setMimeType(ContentService.MimeType.TEXT);
  }

  if (params.code) {
    return handleOAuthCallback(params.code, params.state, params.scope);
  }

  if (params.api === '1') {
    const action = params.action || 'getState';
    let payload = {};
    if (params.payload) {
      try { payload = JSON.parse(params.payload); } catch (err) {}
    } else {
      payload = {
        useTestSheet: params.test === '1' || params.test === 'true',
        coachEmail: params.coachEmail || ''
      };
    }
    return handleApiAction(action, payload);
  }

  const pageLower = (params.page || params.tab || '').toLowerCase();
  const initialTab = pageLower === 'strava' ? 'strava' : (
    (pageLower === 'race' || pageLower === 'waves' || pageLower === 'travel' || pageLower === 'anaheim') ? 'race' : 'training'
  );
  const useTestSheet = params.test === '1' || params.test === 'true';

  return renderUnifiedWebApp(initialTab, useTestSheet);
}

/**
 * Handles Web App POST requests:
 * 1. Standalone GitHub Pages JSON API (Content-Type: text/plain;charset=utf-8 with { api: '1', action, payload })
 * 2. Strava Webhook activity events (object_type === 'activity')
 */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    if (data && (data.api === '1' || data.action)) {
      return handleApiAction(data.action || 'getState', data.payload || {});
    }
    if (data && data.object_type === 'activity' && data.owner_id) {
      const athletes = getAllAthletes();
      const match = athletes.find(a => String(a.athleteId) === String(data.owner_id));
      if (match) {
        recalculateAthleteRuns(match);
      }
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}

/**
 * Renders the unified Web App with:
 * - Tab 1: Training Sign-Up (Register / Deregister for Group & Time-Slotted Sessions + Coach Tools)
 * - Tab 2: Strava Run Accountability Opt-In
 */
function renderUnifiedWebApp(initialTab, useTestSheet) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
  const spreadsheetId = useTestSheet
    ? (props.getProperty('TEST_SPREADSHEET_ID') || SETTINGS.TEST_SPREADSHEET_ID)
    : props.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    return HtmlService.createHtmlOutput(
      '<h3>Configuration Error</h3><p>SPREADSHEET_ID must be set in Script Properties.</p>'
    ).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

      let initialState;
  try {
    initialState = getTrainingAppState(useTestSheet);
  } catch (err) {
    initialState = {
      error: err.message,
      sessions: [],
      roster: [],
      isCoach: false,
      detectedEmail: '',
      useTestSheet: Boolean(useTestSheet),
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`
    };
  }

  const scriptUrl = getWebAppUrl();
  const roster = initialState.roster || [];
  const optionsHtml = roster.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
  const stateJson = JSON.stringify(initialState).replace(/</g, '\\u003c');

  const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
    <title>F45 Morgan Hill HYROX Hub</title>
    <style>
      :root {
        --brand: #FC4C02;
        --brand-dark: #e04402;
        --navy: #111827;
        --bg: #f3f4f6;
        --card: #ffffff;
        --border: #e5e7eb;
        --text: #1f2937;
        --muted: #6b7280;
        --green: #059669;
        --green-bg: #ecfdf5;
        --amber: #d97706;
        --amber-bg: #fffbeb;
        --danger: #dc2626;
      }
      * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
      html, body {
        overflow-x: hidden;
        -webkit-text-size-adjust: 100%;
      }
      body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        background: var(--bg);
        color: var(--text);
        margin: 0;
        padding: 0 0 calc(60px + env(safe-area-inset-bottom, 0px)) 0;
      }
      .topbar {
        background: var(--navy);
        color: #fff;
        padding: 14px 16px 0 16px;
        position: sticky;
        top: 0;
        z-index: 50;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      }
      .topbar-inner {
        max-width: 760px;
        margin: 0 auto;
      }
      .brand-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px 10px;
        margin-bottom: 10px;
      }
      .brand-title {
        font-size: 18px;
        font-weight: 800;
        letter-spacing: -0.01em;
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
      }
      .test-badge {
        background: #f59e0b;
        color: #111827;
        font-size: 11px;
        font-weight: 800;
        padding: 3px 8px;
        border-radius: 999px;
        text-transform: uppercase;
      }
      .sheet-link {
        color: #9ca3af;
        text-decoration: none;
        font-size: 13px;
        font-weight: 500;
        display: inline-flex;
        align-items: center;
        min-height: 32px;
      }
      .sheet-link:hover { color: #fff; }
      .nav-tabs {
        display: flex;
        gap: 6px;
      }
      .nav-tab {
        background: transparent;
        color: #9ca3af;
        border: none;
        border-bottom: 3px solid transparent;
        padding: 10px 14px;
        min-height: 44px;
        font-size: 14px;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.15s;
        display: inline-flex;
        align-items: center;
        justify-content: center;
      }
      .nav-tab.active {
        color: #fff;
        border-bottom-color: var(--brand);
      }
      .container {
        max-width: 760px;
        margin: 18px auto;
        padding: 0 14px;
      }
      .banner-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 10px;
        margin-bottom: 14px;
      }
      .banner-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }
      .btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        border: none;
        border-radius: 8px;
        padding: 10px 14px;
        min-height: 40px;
        font-size: 14px;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.15s;
        touch-action: manipulation;
      }
      .btn:disabled { opacity: 0.55; cursor: not-allowed; }
      .btn-primary { background: var(--brand); color: #fff; }
      .btn-primary:hover:not(:disabled) { background: var(--brand-dark); }
      .btn-danger { background: var(--danger); color: #fff; }
      .btn-danger:hover:not(:disabled) { background: #b91c1c; }
      .btn-outline {
        background: #fff;
        color: var(--text);
        border: 1px solid var(--border);
      }
      .btn-outline:hover:not(:disabled) { background: #f9fafb; }
      .btn-sm { padding: 7px 12px; min-height: 36px; font-size: 12px; border-radius: 6px; }
      .btn-claim { padding: 5px 10px; min-height: 32px; font-size: 12px; border-radius: 6px; flex-shrink: 0; }
      .session-card {
        background: var(--card);
        border-radius: 14px;
        border: 1px solid var(--border);
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        margin-bottom: 18px;
        overflow: hidden;
      }
      .session-header {
        padding: 16px 18px;
        border-bottom: 1px solid var(--border);
        background: linear-gradient(to right, #ffffff, #fff7ed);
      }
      .session-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-bottom: 8px;
      }
      .badge {
        font-size: 11px;
        font-weight: 700;
        padding: 3px 8px;
        border-radius: 999px;
      }
      .badge-slotted { background: #dbeafe; color: #1d4ed8; }
      .badge-group { background: #e0e7ff; color: #4338ca; }
      .badge-open { background: var(--green-bg); color: var(--green); }
      .badge-full { background: #fee2e2; color: var(--danger); }
      .session-title {
        font-size: 18px;
        font-weight: 800;
        margin: 0 0 4px 0;
        color: var(--navy);
        line-height: 1.3;
        word-break: break-word;
      }
      .session-subtitle {
        font-size: 14px;
        font-weight: 600;
        color: var(--brand);
        margin: 0 0 10px 0;
      }
      .meta-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 8px;
        font-size: 13px;
        color: var(--muted);
      }
      .meta-item {
        word-break: break-word;
        line-height: 1.4;
      }
      .meta-item strong { color: var(--text); }
      .details-box {
        margin-top: 10px;
        font-size: 13px;
        color: #4b5563;
        background: #f9fafb;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 10px 12px;
        white-space: pre-wrap;
        word-break: break-word;
      }
      .session-body { padding: 16px 18px; }
      .section-label {
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: var(--muted);
        margin: 0 0 10px 0;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
      }
      .slots-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(215px, 1fr));
        gap: 10px;
        margin-bottom: 16px;
      }
      .slot-card {
        border: 1px solid var(--border);
        border-radius: 10px;
        padding: 10px 12px;
        background: #fff;
        min-width: 0;
      }
      .slot-card.slot-open { border-color: #cbd5e1; }
      .slot-card.slot-full { background: #f9fafb; }
      .slot-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
      }
      .slot-time {
        font-size: 14px;
        font-weight: 800;
        color: var(--navy);
      }
      .slot-cap {
        font-size: 11px;
        font-weight: 600;
        color: var(--muted);
      }
      .spot-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 8px;
        background: #f3f4f6;
        border-radius: 6px;
        padding: 6px 8px;
        margin-bottom: 5px;
        font-size: 13px;
        min-height: 38px;
      }
      .spot-row.empty {
        background: #fff;
        border: 1px dashed #cbd5e1;
        color: var(--muted);
      }
      .spot-name {
        font-weight: 600;
        color: var(--text);
        word-break: break-word;
        min-width: 0;
        flex: 1;
      }
      .remove-btn {
        background: transparent;
        border: none;
        color: #9ca3af;
        font-size: 15px;
        font-weight: 700;
        cursor: pointer;
        padding: 4px 8px;
        min-width: 34px;
        min-height: 34px;
        border-radius: 6px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        touch-action: manipulation;
      }
      .remove-btn:hover { color: var(--danger); background: #fee2e2; }
      .signup-chip-list {
        display: flex;
        flex-direction: column;
        gap: 6px;
        margin-bottom: 14px;
      }
      .signup-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 8px;
        background: #f9fafb;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 8px 12px;
        font-size: 14px;
        min-height: 42px;
      }
      .waitlist-section {
        border-top: 1px dashed var(--border);
        padding-top: 14px;
        margin-top: 12px;
      }
      /* Modal Styles */
      .modal-backdrop {
        position: fixed;
        inset: 0;
        background: rgba(17,24,39,0.6);
        display: none;
        justify-content: center;
        align-items: center;
        z-index: 100;
        padding: 16px;
      }
      .modal-backdrop.open { display: flex; }
      .modal {
        background: #fff;
        border-radius: 14px;
        max-width: 440px;
        width: 100%;
        padding: 22px;
        box-shadow: 0 12px 32px rgba(0,0,0,0.2);
        max-height: 92vh;
        overflow-y: auto;
        -webkit-overflow-scrolling: touch;
      }
      .modal h3 { margin: 0 0 6px 0; font-size: 18px; color: var(--navy); }
      .modal p.sub { margin: 0 0 16px 0; font-size: 13px; color: var(--muted); }
      .field { margin-bottom: 14px; }
      .field label {
        display: block;
        font-size: 13px;
        font-weight: 700;
        margin-bottom: 5px;
        color: #374151;
      }
      .field select, .field input[type="text"], .field input[type="email"], .field input[type="number"], .field textarea {
        width: 100%;
        padding: 10px 12px;
        min-height: 44px;
        border: 1px solid #d1d5db;
        border-radius: 8px;
        font-size: 16px;
        background: #fff;
      }
      .form-grid-2 {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
      }
      .form-grid-3 {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 8px;
      }
      .mode-toggle {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px;
        margin-bottom: 14px;
      }
      .mode-btn {
        padding: 10px 9px;
        min-height: 44px;
        border: 1px solid #d1d5db;
        background: #f9fafb;
        border-radius: 8px;
        font-weight: 700;
        font-size: 14px;
        cursor: pointer;
        touch-action: manipulation;
      }
      .mode-btn.active {
        border-color: var(--brand);
        background: #fff7ed;
        color: var(--brand);
      }
      .modal-actions {
        display: flex;
        justify-content: flex-end;
        gap: 8px;
        margin-top: 18px;
      }
      .toast {
        position: fixed;
        bottom: calc(20px + env(safe-area-inset-bottom, 0px));
        left: 50%;
        transform: translateX(-50%);
        background: var(--navy);
        color: #fff;
        padding: 10px 18px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 600;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        display: none;
        z-index: 200;
        max-width: 92vw;
        text-align: center;
      }
      .strava-card {
        background: white;
        padding: 28px;
        border-radius: 14px;
        border: 1px solid var(--border);
        max-width: 460px;
        margin: 24px auto;
        text-align: center;
      }
      .race-hero {
        background: linear-gradient(135deg, #111827 0%, #1f2937 100%);
        color: #fff;
        border-radius: 14px;
        padding: 16px 18px;
        margin-bottom: 14px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 12px;
      }
      .race-hero-title { font-size: 18px; font-weight: 800; margin: 0 0 4px 0; }
      .race-hero-sub { font-size: 13px; color: #d1d5db; margin: 0; }
      .race-subtabs {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 6px;
        background: #e5e7eb;
        padding: 4px;
        border-radius: 10px;
        margin-bottom: 14px;
      }
      .race-subtab {
        border: none;
        background: transparent;
        color: #4b5563;
        font-size: 14px;
        font-weight: 700;
        padding: 10px 12px;
        border-radius: 8px;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        min-height: 42px;
      }
      .race-subtab.active { background: #fff; color: var(--navy); box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
      .day-pills { display: flex; gap: 6px; overflow-x: auto; padding-bottom: 4px; margin-bottom: 12px; }
      .day-pill {
        border: 1px solid var(--border);
        background: #fff;
        color: var(--text);
        padding: 7px 12px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 700;
        cursor: pointer;
        white-space: nowrap;
        min-height: 36px;
      }
      .day-pill.active { background: var(--navy); color: #fff; border-color: var(--navy); }
      .race-filter-bar { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; align-items: center; }
      .race-search-input { flex: 1 1 180px; padding: 9px 12px; min-height: 40px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; background: #fff; }
      .race-div-select { flex: 0 1 180px; padding: 9px 10px; min-height: 40px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; background: #fff; }
      .race-day-card { background: #fff; border-radius: 14px; border: 1px solid var(--border); margin-bottom: 16px; overflow: hidden; }
      .race-day-header { padding: 13px 16px; background: linear-gradient(to right, #f9fafb, #fff7ed); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; gap: 8px; }
      .race-day-title { font-size: 16px; font-weight: 800; color: var(--navy); margin: 0; }
      .wave-row { display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 11px 16px; border-bottom: 1px solid #f3f4f6; }
      .wave-row:last-child { border-bottom: none; }
      .wave-left { display: flex; align-items: center; gap: 12px; min-width: 0; flex: 1; }
      .wave-time-badge { min-width: 88px; text-align: center; font-size: 13px; font-weight: 800; padding: 6px 10px; border-radius: 8px; background: var(--navy); color: #fff; flex-shrink: 0; cursor: pointer; border: none; }
      .wave-time-badge.unset { background: #fff7ed; color: var(--brand); border: 1px dashed #fdba74; font-weight: 700; font-size: 12px; }
      .wave-info { min-width: 0; flex: 1; }
      .wave-athletes { font-size: 15px; font-weight: 700; color: var(--text); word-break: break-word; }
      .wave-meta { margin-top: 3px; display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }
      .div-pill { font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 999px; background: #f3f4f6; color: #374151; }
      .div-pill.doubles { background: #dbeafe; color: #1d4ed8; }
      .div-pill.relay { background: #ede9fe; color: #6d28d9; }
      .div-pill.pro { background: #fef3c7; color: #b45309; }
      .travel-summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 12px; margin-bottom: 16px; }
      .summary-card { background: #fff; border: 1px solid var(--border); border-radius: 14px; padding: 14px 16px; }
      .summary-card h4 { margin: 0 0 10px 0; font-size: 14px; font-weight: 800; color: var(--navy); display: flex; align-items: center; justify-content: space-between; }
      .dinner-row { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; padding: 7px 0; border-bottom: 1px dashed #e5e7eb; font-size: 13px; }
      .dinner-row:last-child { border-bottom: none; padding-bottom: 0; }
      .dinner-day-label { font-weight: 700; color: var(--navy); white-space: nowrap; min-width: 92px; }
      .dinner-names { color: var(--text); flex: 1; text-align: right; word-break: break-word; }
      .hotel-group-item { padding: 7px 0; border-bottom: 1px dashed #e5e7eb; font-size: 13px; }
      .hotel-group-item:last-child { border-bottom: none; padding-bottom: 0; }
      .travel-card { background: #fff; border: 1px solid var(--border); border-radius: 14px; padding: 15px 16px; margin-bottom: 12px; }
      .travel-card-top { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 10px; }
      .travel-name { font-size: 16px; font-weight: 800; color: var(--navy); }
      .travel-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 8px; margin-bottom: 10px; }
      .travel-leg-box { background: #f9fafb; border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 13px; }
      .travel-leg-label { font-size: 11px; font-weight: 800; text-transform: uppercase; color: var(--muted); margin-bottom: 3px; display: flex; justify-content: space-between; align-items: center; }
      .travel-day-tag { background: #ffedd5; color: #c2410c; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: 700; }
      .dinner-toggle-grid { display: flex; flex-wrap: wrap; gap: 6px; }
      .dinner-toggle-btn { border: 1px solid #d1d5db; background: #f9fafb; color: var(--text); border-radius: 8px; padding: 8px 10px; font-size: 13px; font-weight: 700; cursor: pointer; min-height: 38px; }
      .dinner-toggle-btn.active { background: #ecfdf5; border-color: var(--green); color: var(--green); }
      .hotel-chip-btn { border: 1px solid #d1d5db; background: #f9fafb; color: #374151; border-radius: 999px; padding: 4px 10px; font-size: 12px; font-weight: 600; cursor: pointer; }
      @media (max-width: 600px) {
        .topbar {
          padding: 10px 12px 0 12px;
        }
        .brand-row {
          margin-bottom: 8px;
        }
        .brand-title {
          font-size: 16px;
        }
        .sheet-link {
          font-size: 12px;
        }
        .nav-tabs {
          gap: 0;
        }
        .nav-tab {
          flex: 1;
          padding: 10px 6px;
          font-size: 13px;
          min-height: 44px;
        }
        .container {
          margin: 12px auto;
          padding: 0 10px;
        }
        .banner-bar {
          flex-direction: column;
          align-items: stretch;
          gap: 8px;
        }
        .banner-actions {
          width: 100%;
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
        }
        .banner-actions .btn {
          flex: 1 1 auto;
          min-height: 42px;
          font-size: 13px;
        }
        .session-header {
          padding: 14px 14px;
        }
        .session-body {
          padding: 14px 14px;
        }
        .session-title {
          font-size: 16px;
        }
        .meta-grid {
          grid-template-columns: 1fr;
          gap: 6px;
        }
        .section-label {
          flex-direction: column;
          align-items: stretch;
          gap: 8px;
        }
        .section-label .btn {
          width: 100%;
          min-height: 42px;
          font-size: 13px;
        }
        .slots-grid {
          grid-template-columns: 1fr;
          gap: 8px;
        }
        .spot-row {
          min-height: 44px;
          padding: 6px 10px;
        }
        .btn-claim {
          min-height: 40px;
          padding: 8px 14px;
          font-size: 13px;
        }
        .remove-btn {
          min-width: 42px;
          min-height: 42px;
          font-size: 16px;
        }
        .signup-item {
          min-height: 46px;
          padding: 6px 10px 6px 12px;
        }
        .modal-backdrop {
          padding: 0;
          align-items: flex-end;
        }
        .modal {
          max-width: 100%;
          border-radius: 18px 18px 0 0;
          padding: 20px 16px calc(20px + env(safe-area-inset-bottom, 0px)) 16px;
          max-height: 90vh;
        }
        .modal-actions {
          position: sticky;
          bottom: calc(-1 * (20px + env(safe-area-inset-bottom, 0px)));
          background: #fff;
          padding: 12px 0 calc(8px + env(safe-area-inset-bottom, 0px)) 0;
          margin-top: 14px;
          border-top: 1px solid var(--border);
          display: grid;
          grid-template-columns: 1fr 1.3fr;
          gap: 10px;
        }
        .modal-actions .btn {
          width: 100%;
          min-height: 44px;
          font-size: 15px;
        }
        .form-grid-2 {
          grid-template-columns: 1fr;
          gap: 0;
        }
        .strava-card {
          padding: 20px 16px;
          margin: 12px auto;
        }
      }
    </style>
  </head>
  <body>
    <div class="topbar">
      <div class="topbar-inner">
        <div class="brand-row">
          <div class="brand-title">
            <span>🔥 F45 Morgan Hill HYROX</span>
            <span id="testBadge" class="test-badge" style="display:none;">Test Sheet</span>
          </div>
          <a id="sheetLink" class="sheet-link" href="#" target="_blank" rel="noopener">📄 View Google Sheet ↗</a>
        </div>
        <div class="nav-tabs">
          <button id="tabBtnTraining" class="nav-tab" onclick="switchTab('training')">🏋️ Training</button>
          <button id="tabBtnRace" class="nav-tab" onclick="switchTab('race')">🏁 Anaheim 2026</button>
          <button id="tabBtnStrava" class="nav-tab" onclick="switchTab('strava')">⚡ Strava</button>
        </div>
      </div>
    </div>

    <!-- TAB 1: TRAINING SIGN-UP -->
    <div id="tabTraining" class="container" style="display:none;">
      <div class="banner-bar">
        <div style="font-size: 13px; color: var(--muted);">
          Tap a session or time slot below to sign up (solo or doubles) or remove your signup.
        </div>
        <div class="banner-actions">
          <button class="btn btn-outline btn-sm" onclick="refreshSessions()">🔄 Refresh</button>
          <button id="btnAddSession" class="btn btn-primary btn-sm" onclick="openCoachModal()">⚙️ Coach: Add Session</button>
        </div>
      </div>

      <div id="upcomingContainer"></div>

      <div id="pastWrapper" style="margin-top: 24px; display: none;">
        <button id="togglePastBtn" class="btn btn-outline" style="width:100%;" onclick="togglePastSessions()">
          🕓 Show Past Training Sessions
        </button>
        <div id="pastContainer" style="display:none; margin-top: 14px;"></div>
      </div>
    </div>

    <!-- TAB 2: HYROX ANAHEIM 2026 RACE HUB -->
    <div id="tabRace" class="container" style="display:none;">
      <div class="race-hero">
        <div>
          <h2 class="race-hero-title">🏁 HYROX Anaheim 2026</h2>
          <p class="race-hero-sub">📅 Thu Dec 3 – Sun Dec 6, 2026 • 📍 Anaheim Convention Center</p>
        </div>
        <div style="display:flex; gap:8px; flex-wrap:wrap;">
          <button class="btn btn-outline btn-sm" style="background:rgba(255,255,255,0.1); color:#fff; border-color:rgba(255,255,255,0.25);" onclick="refreshSessions()">🔄 Refresh</button>
        </div>
      </div>

      <div class="race-subtabs">
        <button type="button" id="subtabBtnWaves" class="race-subtab active" onclick="switchRaceSubtab('waves')">
          ⏱️ Start Waves <span id="wavesCountBadge" class="badge badge-open" style="margin-left:2px;">0</span>
        </button>
        <button type="button" id="subtabBtnTravel" class="race-subtab" onclick="switchRaceSubtab('travel')">
          ✈️ Travel &amp; Dinners <span id="travelCountBadge" class="badge badge-slotted" style="margin-left:2px;">0</span>
        </button>
      </div>

      <div id="raceWavesView">
        <div id="raceDayPills" class="day-pills"></div>
        <div class="race-filter-bar">
          <input type="text" id="waveSearchInput" class="race-search-input" placeholder="🔍 Search athlete or team..." oninput="renderRaceWaves()">
          <select id="waveDivFilter" class="race-div-select" onchange="renderRaceWaves()">
            <option value="ALL">All Divisions</option>
            <option value="SINGLES">Singles (Open / Pro)</option>
            <option value="DOUBLES">Doubles</option>
            <option value="RELAY">Relay Teams</option>
          </select>
          <button type="button" class="btn btn-primary btn-sm" onclick="openWaveModal(null, false)">+ Add / Update My Wave</button>
        </div>
        <div id="raceWavesContainer"></div>
      </div>

      <div id="raceTravelView" style="display:none;">
        <div class="banner-bar">
          <div style="font-size: 13px; color: var(--muted);">
            Coordinate flights, carpools, hotels, and group dinners in Anaheim!
          </div>
          <div class="banner-actions">
            <button type="button" class="btn btn-primary btn-sm" onclick="openTravelModal(null)">+ Add / Edit My Travel</button>
          </div>
        </div>
        <div id="raceTravelSummaryContainer" class="travel-summary-grid"></div>
        <div id="raceTravelCardsContainer"></div>
      </div>
    </div>

    <datalist id="clubRosterDatalist"></datalist>

    <!-- TAB 2: STRAVA RUN ACCOUNTABILITY -->
    <div id="tabStrava" class="container" style="display:none;">
      <div class="strava-card">
        <h2 style="color: var(--brand); margin-top: 0;">🏃 F45 Run Accountability</h2>
        <p style="color: #555; line-height: 1.5; font-size: 15px;">
          Connect your Strava account to automatically sync your runs &amp; weekend Hyrox workouts to the club sheet!
        </p>
        <form id="authForm" action="https://www.strava.com/oauth/authorize" method="GET" target="_top" onsubmit="return validateStravaForm()">
          <input type="hidden" name="client_id" value="${escapeHtml(clientId)}">
          <input type="hidden" name="response_type" value="code">
          <input type="hidden" name="redirect_uri" value="${escapeHtml(scriptUrl)}">
          <input type="hidden" name="approval_prompt" value="auto">
          <input type="hidden" name="scope" value="read,activity:read_all">
          <input type="hidden" id="stateInput" name="state" value="">

          <div class="field" style="text-align:left;">
            <label>Select your name in the sheet:</label>
            <select id="nameSelect" onchange="toggleCustomStravaName(this.value)">
              <option value="" disabled selected>-- Choose your name --</option>
              ${optionsHtml}
              <option value="__NEW__">➕ Not listed? Add my name...</option>
            </select>
          </div>

          <div id="customNameDiv" class="field" style="display: none; text-align:left;">
            <label>Enter your name as it should appear in the table:</label>
            <input type="text" id="customNameInput" placeholder="e.g. First Name / Nickname">
          </div>

          <button type="submit" class="btn btn-primary" style="width:100%; padding:13px; font-size:16px;">
            Connect with Strava
          </button>
        </form>
        <div style="font-size: 12px; color: #888; margin-top: 15px;">
          Counts running activities + weekend "Hyrox" workouts. Manual numbers are preserved if higher!
        </div>
      </div>
    </div>

    <!-- SIGN-UP MODAL -->
    <div id="signupModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeSignupModal()">
      <div class="modal">
        <h3 id="modalTitle">Sign Up for Training</h3>
        <p id="modalSubtitle" class="sub"></p>

        <div id="slotPickerField" class="field" style="display:none;">
          <label>Time Slot</label>
          <select id="modalSlotSelect"></select>
        </div>

        <div class="field">
          <label>Training Format (Solo and Doubles both use 1 slot)</label>
          <div class="mode-toggle">
            <button type="button" id="btnModeSolo" class="mode-btn active" onclick="setPairMode('solo')">👤 Solo</button>
            <button type="button" id="btnModePair" class="mode-btn" onclick="setPairMode('pair')">👥 Doubles / Pair</button>
          </div>
        </div>

        <div class="field">
          <label id="athlete1Label">Athlete Name</label>
          <select id="athlete1Select" onchange="toggleCustomInput('athlete1Select', 'athlete1Custom')"></select>
          <input type="text" id="athlete1Custom" placeholder="Type name..." style="display:none; margin-top:6px;">
        </div>

        <div id="athlete2Field" class="field" style="display:none;">
          <label>Partner Name</label>
          <select id="athlete2Select" onchange="toggleCustomInput('athlete2Select', 'athlete2Custom')"></select>
          <input type="text" id="athlete2Custom" placeholder="Type partner's name..." style="display:none; margin-top:6px;">
        </div>

        <div class=" field" id="noteField" style="display:none;">
          <label>Optional Note (e.g. "Robin's friend")</label>
          <input type="text" id="signupNoteInput" placeholder="Optional note...">
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeSignupModal()">Cancel</button>
          <button type="button" id="confirmSignupBtn" class="btn btn-primary" onclick="submitSignup()">Confirm Sign-Up</button>
        </div>
      </div>
    </div>

    <!-- COACH ADD SESSION MODAL -->
    <div id="coachModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeCoachModal()">
      <div class="modal" style="max-width: 500px;">
        <h3>⚙️ Coach: Add Training Session</h3>
        <p class="sub">Inserts a new session row at the top of the TRAINING SESSIONS sheet.</p>

        <div class="field">
          <label>Date Header (Column A)</label>
          <input type="text" id="coachDate" placeholder="e.g. TRAINING ON: Sunday 11/22/26">
        </div>
        <div class="field">
          <label>Event Name (Column D)</label>
          <input type="text" id="coachEvent" value="HYROX training!">
        </div>
        <div class="form-grid-2">
          <div class="field">
            <label>Time (Column B)</label>
            <input type="text" id="coachTime" placeholder="e.g. 10:45am-12:15pm">
          </div>
          <div class="field">
            <label>RSVP By (Column C)</label>
            <input type="text" id="coachRsvp" placeholder="e.g. Sat 11/21 at noon">
          </div>
        </div>
        <div class="form-grid-2">
          <div class="field">
            <label>Location (Column E)</label>
            <input type="text" id="coachLocation" value="F45">
          </div>
          <div class="field">
            <label>Coach / Coordinated By (Column I)</label>
            <input type="text" id="coachCoord" value="Jade">
          </div>
        </div>
        <div class="field">
          <label>Sign-Up Mode</label>
          <div class="mode-toggle">
            <button type="button" id="coachModeGroup" class="mode-btn active" onclick="setCoachSessionMode('group')">👥 Group Session</button>
            <button type="button" id="coachModeSlotted" class="mode-btn" onclick="setCoachSessionMode('slotted')">⏱️ Time Slots</button>
          </div>
        </div>

        <div id="coachGroupConfig" style="background:#f9fafb; border:1px solid var(--border); padding:12px; border-radius:8px; margin-bottom:14px;">
          <div class="field" style="margin:0;">
            <label># of Member Slots (0 = unlimited)</label>
            <input type="number" id="coachGroupCap" value="20" min="0" max="100">
          </div>
        </div>

        <div id="coachSlotsConfig" style="display:none; background:#f9fafb; border:1px solid var(--border); padding:12px; border-radius:8px; margin-bottom:14px;">
          <div class="form-grid-3">
            <div class="field" style="margin:0;">
              <label>First Slot</label>
              <input type="text" id="coachFirstSlot" value="10:30 AM">
            </div>
            <div class="field" style="margin:0;">
              <label>Every (mins)</label>
              <input type="number" id="coachIntervalMins" value="8" min="1" max="120">
            </div>
            <div class="field" style="margin:0;">
              <label># of Slots</label>
              <input type="number" id="coachNumSlots" value="12" min="1" max="40">
            </div>
          </div>
          <div class="field" style="margin:10px 0 0 0;">
            <label>Groups / Athletes per Time Slot</label>
            <input type="number" id="coachSlotCap" value="2" min="1" max="10">
          </div>
        </div>

        <div class="field">
          <label>Workout Notes (Optional, Column J)</label>
          <textarea id="coachNotes" rows="3" placeholder="Workout format, station notes..."></textarea>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeCoachModal()">Cancel</button>
          <button type="button" id="coachCreateBtn" class="btn btn-primary" onclick="submitCoachNewSession()">Create Session</button>
        </div>
      </div>
    </div>

    <!-- COACH MANAGE EXISTING SESSION MODAL -->
    <div id="manageModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeManageModal()">
      <div class="modal" style="max-width: 500px;">
        <h3>⚙️ Coach: Manage Session</h3>
        <p id="manageModalSubtitle" class="sub"></p>

        <div class="field">
          <label>Date Header (Column A)</label>
          <input type="text" id="manageDate" placeholder="e.g. TRAINING ON: Sunday 11/22/26">
        </div>
        <div class="field">
          <label>Event Name (Column D)</label>
          <input type="text" id="manageEvent" placeholder="e.g. HYROX training!">
        </div>
        <div class="form-grid-2">
          <div class="field">
            <label>Time (Column B)</label>
            <input type="text" id="manageTime" placeholder="e.g. 10:45am-12:15pm">
          </div>
          <div class="field">
            <label>RSVP By (Column C)</label>
            <input type="text" id="manageRsvp" placeholder="e.g. Sat 11/21 at noon">
          </div>
        </div>
        <div class="form-grid-2">
          <div class="field">
            <label>Location (Column E)</label>
            <input type="text" id="manageLocation" placeholder="e.g. F45">
          </div>
          <div class="field">
            <label>Coach / Coordinated By (Column I)</label>
            <input type="text" id="manageCoord" placeholder="e.g. Jade">
          </div>
        </div>

        <div class="field">
          <label>Sign-Up Mode</label>
          <div class="mode-toggle">
            <button type="button" id="manageModeGroup" class="mode-btn" onclick="setManageSessionMode('group')">👥 Group Session</button>
            <button type="button" id="manageModeSlotted" class="mode-btn" onclick="setManageSessionMode('slotted')">⏱️ Time Slots</button>
          </div>
        </div>

        <div id="manageGroupConfig" style="display:none; background:#f9fafb; border:1px solid var(--border); padding:12px; border-radius:8px; margin-bottom:14px;">
          <div class="field" style="margin:0;">
            <label># of Member Slots (0 = unlimited)</label>
            <input type="number" id="manageGroupCap" value="20" min="0" max="100" oninput="updateManagePreview()">
          </div>
          <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
            <button type="button" class="btn btn-outline btn-sm" onclick="adjustManageSlots('group', 2)">+2 Slots</button>
            <button type="button" class="btn btn-outline btn-sm" onclick="adjustManageSlots('group', 5)">+5 Slots</button>
          </div>
          <div id="manageGroupHint" style="font-size:12px; color:var(--muted); margin-top:8px;"></div>
        </div>

        <div id="manageSlotsConfig" style="display:none; background:#f9fafb; border:1px solid var(--border); padding:12px; border-radius:8px; margin-bottom:14px;">
          <div class="form-grid-3">
            <div class="field" style="margin:0;">
              <label>First Slot</label>
              <input type="text" id="manageFirstSlot" value="10:30 AM" oninput="updateManagePreview()">
            </div>
            <div class="field" style="margin:0;">
              <label>Every (mins)</label>
              <input type="number" id="manageIntervalMins" value="8" min="1" max="120" oninput="updateManagePreview()">
            </div>
            <div class="field" style="margin:0;">
              <label># of Slots</label>
              <input type="number" id="manageNumSlots" value="12" min="1" max="40" oninput="updateManagePreview()">
            </div>
          </div>
          <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
            <button type="button" class="btn btn-outline btn-sm" onclick="adjustManageSlots('slotted', 1)">+1 Time Slot</button>
            <button type="button" class="btn btn-outline btn-sm" onclick="adjustManageSlots('slotted', 2)">+2 Time Slots</button>
            <button type="button" class="btn btn-outline btn-sm" onclick="adjustManageSlots('slotted', 4)">+4 Time Slots</button>
          </div>
          <div class="field" style="margin:12px 0 0 0;">
            <label>Groups / Athletes per Time Slot</label>
            <input type="number" id="manageSlotCap" value="2" min="1" max="10" oninput="updateManagePreview()">
          </div>
          <div id="manageSchedulePreview" style="font-size:12px; color:var(--muted); margin-top:10px; line-height:1.4;"></div>
        </div>

        <div class="field">
          <label>Workout Notes (Column J)</label>
          <textarea id="manageNotes" rows="3" placeholder="Workout format, station notes..."></textarea>
        </div>

        <details style="margin-bottom:14px; font-size:13px; color:#374151;">
          <summary style="cursor:pointer; font-weight:700; padding:4px 0;">💲 Edit Cost &amp; Registration Info (Column F)</summary>
          <div class="field" style="margin-top:8px; margin-bottom:0;">
            <textarea id="manageCost" rows="3" placeholder="Cost and registration policy..."></textarea>
          </div>
        </details>

        <div style="border-top:1px dashed var(--border); padding-top:12px; margin-top:6px;">
          <button type="button" class="btn btn-outline btn-sm" style="color:var(--danger); border-color:#fecaca; width:100%;" onclick="openDeleteSessionModal()">
            🗑️ Delete / Cancel Session
          </button>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeManageModal()">Cancel</button>
          <button type="button" id="manageSaveBtn" class="btn btn-primary" onclick="submitManageSession()">Save Changes</button>
        </div>
      </div>
    </div>

    <!-- DELETE SESSION CONFIRMATION MODAL -->
    <div id="deleteSessionModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeDeleteSessionModal()">
      <div class="modal" style="max-width: 400px;">
        <h3>🗑️ Delete / Cancel Session?</h3>
        <p id="deleteSessionModalText" class="sub" style="font-size: 14px; color: var(--text); margin: 10px 0 18px 0; line-height: 1.5;"></p>
        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeDeleteSessionModal()">Keep Session</button>
          <button type="button" id="confirmDeleteSessionBtn" class="btn btn-danger" onclick="executeDeleteSession()">Yes, Delete Session</button>
        </div>
      </div>
    </div>

    <!-- REMOVE CONFIRMATION MODAL -->
    <div id="removeModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeRemoveModal()">
      <div class="modal" style="max-width: 380px;">
        <h3>🗑️ Remove Sign-Up?</h3>
        <p id="removeModalText" class="sub" style="font-size: 15px; color: var(--text); margin: 10px 0 18px 0;"></p>
        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeRemoveModal()">Cancel</button>
          <button type="button" id="confirmRemoveBtn" class="btn btn-danger" onclick="executeDeregister()">Remove</button>
        </div>
      </div>
    </div>

    <!-- RACE WAVE ADD / EDIT MODAL -->
    <div id="waveModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeWaveModal()">
      <div class="modal" style="max-width: 480px;">
        <h3 id="waveModalTitle">⏱️ Add / Update Start Wave</h3>
        <p class="sub">Updates the <b>Hyrox Anaheim 2026 times</b> sheet tab.</p>
        <div class="field">
          <label>Select Athlete / Team Entry</label>
          <select id="waveExistingSelect" onchange="onWaveExistingChange(this.value)"></select>
        </div>
        <div class="form-grid-2">
          <div class="field">
            <label>Race Day</label>
            <select id="waveDaySelect">
              <option value="Thursday">Thursday (12/3)</option>
              <option value="Friday">Friday (12/4)</option>
              <option value="Saturday">Saturday (12/5)</option>
              <option value="Sunday">Sunday (12/6)</option>
            </select>
          </div>
          <div class="field">
            <label>Division</label>
            <select id="waveDivisionSelect" onchange="onWaveDivisionChange()">
              <option value="Womens Open">Womens Open</option>
              <option value="Mens Open">Mens Open</option>
              <option value="Womens Pro">Womens Pro</option>
              <option value="Mens Pro">Mens Pro</option>
              <option value="Womens Doubles Open">Womens Doubles Open</option>
              <option value="Mens Doubles Open">Mens Doubles Open</option>
              <option value="Mixed Doubles Open">Mixed Doubles Open</option>
              <option value="Womens Doubles Pro">Womens Doubles Pro</option>
              <option value="Mens Doubles Pro">Mens Doubles Pro</option>
              <option value="Mixed Doubles Pro">Mixed Doubles Pro</option>
              <option value="Womens Relay">Womens Relay</option>
              <option value="Mens Relay">Mens Relay</option>
              <option value="Mixed Relay">Mixed Relay</option>
            </select>
          </div>
        </div>
        <div class="field">
          <label>Start Wave Time</label>
          <div class="mode-toggle" style="margin-bottom:8px;">
            <button type="button" id="waveTimeModeSet" class="mode-btn active" onclick="setWaveTimeMode('set')">⏱️ Set Wave Time</button>
            <button type="button" id="waveTimeModeTbd" class="mode-btn" onclick="setWaveTimeMode('tbd')">⏳ Not Released Yet</button>
          </div>
          <div id="waveTimePickerRow" class="form-grid-3">
            <div>
              <select id="waveHourSelect" onchange="onWaveHourChange(this.value)">
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10" selected>10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="1">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
              </select>
            </div>
            <div>
              <select id="waveMinSelect">
                <option value="00">:00</option>
                <option value="05">:05</option>
                <option value="10">:10</option>
                <option value="15">:15</option>
                <option value="20">:20</option>
                <option value="25">:25</option>
                <option value="30">:30</option>
                <option value="35">:35</option>
                <option value="40">:40</option>
                <option value="45">:45</option>
                <option value="50">:50</option>
                <option value="55">:55</option>
              </select>
            </div>
            <div>
              <select id="waveAmPmSelect">
                <option value="AM" selected>AM</option>
                <option value="PM">PM</option>
              </select>
            </div>
          </div>
        </div>
        <div class="field">
          <label id="wavePrimaryLabel">Athlete Name</label>
          <input type="text" id="waveNameInput" list="clubRosterDatalist" placeholder="e.g. Steve T">
        </div>
        <div id="wavePartner1Field" class="field" style="display:none;">
          <label id="wavePartner1Label">Doubles / Relay Partner 1</label>
          <input type="text" id="wavePartner1Input" list="clubRosterDatalist" placeholder="Partner's name...">
        </div>
        <div id="waveRelayPartnersField" class="form-grid-2" style="display:none;">
          <div class="field">
            <label>Relay Partner 2</label>
            <input type="text" id="wavePartner2Input" list="clubRosterDatalist" placeholder="Partner 3...">
          </div>
          <div class="field">
            <label>Relay Partner 3</label>
            <input type="text" id="wavePartner3Input" list="clubRosterDatalist" placeholder="Partner 4...">
          </div>
        </div>
        <div id="waveDeleteWrap" style="display:none; border-top:1px dashed var(--border); padding-top:12px; margin-top:6px;">
          <button type="button" class="btn btn-outline btn-sm" style="color:var(--danger); border-color:#fecaca; width:100%;" onclick="openDeleteWaveModal()">
            🗑️ Delete This Race Entry
          </button>
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeWaveModal()">Cancel</button>
          <button type="button" id="waveSaveBtn" class="btn btn-primary" onclick="submitWaveModal()">Save Wave</button>
        </div>
      </div>
    </div>

    <div id="deleteWaveModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeDeleteWaveModal()">
      <div class="modal" style="max-width: 390px;">
        <h3>🗑️ Delete Race Entry?</h3>
        <p id="deleteWaveModalText" class="sub" style="font-size: 14px; color: var(--text); margin: 10px 0 18px 0;"></p>
        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeDeleteWaveModal()">Keep Entry</button>
          <button type="button" id="confirmDeleteWaveBtn" class="btn btn-danger" onclick="executeDeleteWave()">Yes, Delete</button>
        </div>
      </div>
    </div>

    <!-- TRAVEL, HOTEL & DINNERS ADD / EDIT MODAL -->
    <div id="travelModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeTravelModal()">
      <div class="modal" style="max-width: 500px;">
        <h3 id="travelModalTitle">✈️ Add / Edit Travel &amp; Dinners</h3>
        <p class="sub">Updates the <b>Travel ✈️ Hyrox Anaheim 2026!!</b> sheet tab.</p>
        <div class="field">
          <label>Select Entry</label>
          <select id="travelExistingSelect" onchange="onTravelExistingChange(this.value)"></select>
        </div>
        <div class="field">
          <label>Name(s) (e.g. "Steve T" or "Jade &amp; Aaron")</label>
          <input type="text" id="travelNameInput" list="clubRosterDatalist" placeholder="Your name(s)...">
        </div>
        <div style="background:#f9fafb; border:1px solid var(--border); padding:12px; border-radius:10px; margin-bottom:14px;">
          <div class="field" style="margin-bottom:8px;">
            <label>🛫 Outbound to Anaheim — Day</label>
            <select id="travelOutboundDaySelect"></select>
          </div>
          <div class="field" style="margin:0;">
            <label>Outbound Flight / Driving / Carpool Info</label>
            <textarea id="travelOutboundDetails" rows="2" placeholder="e.g. Southwest SJC to LGB @ 12:55pm, or Driving w/ room for 2"></textarea>
          </div>
        </div>
        <div style="background:#f9fafb; border:1px solid var(--border); padding:12px; border-radius:10px; margin-bottom:14px;">
          <div class="field" style="margin-bottom:8px;">
            <label>🛬 Return Home — Day</label>
            <select id="travelReturnDaySelect"></select>
          </div>
          <div class="field" style="margin:0;">
            <label>Return Flight / Driving Info</label>
            <textarea id="travelReturnDetails" rows="2" placeholder="e.g. Southwest LGB to SJC @ 5:45pm, or Driving Sunday morning"></textarea>
          </div>
        </div>
        <div class="field">
          <label>🏨 Hotel / Lodging</label>
          <input type="text" id="travelLodgingInput" placeholder="e.g. BW Lulu Anaheim, Hilton Anaheim, Sheraton...">
          <div id="travelHotelChips" style="display:flex; flex-wrap:wrap; gap:6px; margin-top:6px;"></div>
        </div>
        <div class="field">
          <label>🍽️ Interested in Group Dinner (tap all that apply)</label>
          <div id="travelDinnerToggles" class="dinner-toggle-grid"></div>
        </div>
        <div class="field">
          <label>🎡 Other Plans (Disneyland, Rooftop Drinks, etc.)</label>
          <input type="text" id="travelOtherPlansInput" placeholder="Optional plans to share with the group...">
        </div>
        <div id="travelDeleteWrap" style="display:none; border-top:1px dashed var(--border); padding-top:12px; margin-top:6px;">
          <button type="button" class="btn btn-outline btn-sm" style="color:var(--danger); border-color:#fecaca; width:100%;" onclick="openDeleteTravelModal()">
            🗑️ Delete Travel Entry
          </button>
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeTravelModal()">Cancel</button>
          <button type="button" id="travelSaveBtn" class="btn btn-primary" onclick="submitTravelModal()">Save Travel Info</button>
        </div>
      </div>
    </div>

    <div id="deleteTravelModalBackdrop" class="modal-backdrop" onclick="if(event.target===this) closeDeleteTravelModal()">
      <div class="modal" style="max-width: 390px;">
        <h3>🗑️ Delete Travel Entry?</h3>
        <p id="deleteTravelModalText" class="sub" style="font-size: 14px; color: var(--text); margin: 10px 0 18px 0;"></p>
        <div class="modal-actions">
          <button type="button" class="btn btn-outline" onclick="closeDeleteTravelModal()">Keep Entry</button>
          <button type="button" id="confirmDeleteTravelBtn" class="btn btn-danger" onclick="executeDeleteTravel()">Yes, Delete</button>
        </div>
      </div>
    </div>

    <div id="toast" class="toast"></div>

    <script>
      let APP_STATE = ${stateJson};
      let ACTIVE_TAB = "${escapeHtml(initialTab)}";
      let PAST_OPEN = false;
      let CLIENT_COACH_EMAIL = '';
      try {
        if (typeof localStorage !== 'undefined' && localStorage) {
          CLIENT_COACH_EMAIL = String(localStorage.getItem('hyrox_coach_email') || '').trim().toLowerCase();
        }
      } catch (e) {}

      let MODAL_CTX = {
        rowNumber: null,
        sessionKey: '',
        listType: 'member',
        mode: 'group',
        slotTime: null,
        pairMode: 'solo'
      };
      let REMOVE_CTX = null;
      let COACH_MODE = 'group';
      let MANAGE_CTX = {
        rowNumber: null,
        sessionKey: '',
        mode: 'group',
        minSlots: 1,
        minSlotCap: 1,
        minGroupCap: 0
      };

      function getActiveCoachEmail() {
        return String(APP_STATE.detectedEmail || CLIENT_COACH_EMAIL || '').trim().toLowerCase();
      }

      function initApp() {
        switchTab(ACTIVE_TAB);
        switchRaceSubtab(RACE_SUBTAB);
        renderAll();
      }

      function switchTab(tab) {
        ACTIVE_TAB = tab === 'strava' ? 'strava' : (tab === 'race' ? 'race' : 'training');
        document.getElementById('tabTraining').style.display = ACTIVE_TAB === 'training' ? 'block' : 'none';
        document.getElementById('tabRace').style.display = ACTIVE_TAB === 'race' ? 'block' : 'none';
        document.getElementById('tabStrava').style.display = ACTIVE_TAB === 'strava' ? 'block' : 'none';
        document.getElementById('tabBtnTraining').classList.toggle('active', ACTIVE_TAB === 'training');
        document.getElementById('tabBtnRace').classList.toggle('active', ACTIVE_TAB === 'race');
        document.getElementById('tabBtnStrava').classList.toggle('active', ACTIVE_TAB === 'strava');
      }

      function showToast(msg) {
        const el = document.getElementById('toast');
        el.textContent = msg;
        el.style.display = 'block';
        clearTimeout(window.__toastTimer);
        window.__toastTimer = setTimeout(() => { el.style.display = 'none'; }, 3400);
      }

      function escapeHtmlClient(s) {
        return String(s || '')
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#039;');
      }

      function renderAll() {
        const badge = document.getElementById('testBadge');
        badge.style.display = APP_STATE.useTestSheet ? 'inline-block' : 'none';
        const link = document.getElementById('sheetLink');
        if (APP_STATE.spreadsheetUrl) link.href = APP_STATE.spreadsheetUrl;

        renderRaceHub();

        const upcomingEl = document.getElementById('upcomingContainer');
        const pastEl = document.getElementById('pastContainer');
        const pastWrap = document.getElementById('pastWrapper');

        if (APP_STATE.error) {
          upcomingEl.innerHTML = '<div class="session-card"><div class="session-body" style="color:#dc2626;">Error: ' +
            escapeHtmlClient(APP_STATE.error) + '</div></div>';
          return;
        }

        const sessions = APP_STATE.sessions || [];
        const upcoming = sessions.filter(s => s.isUpcoming);
        const past = sessions.filter(s => !s.isUpcoming);

        if (upcoming.length === 0) {
          upcomingEl.innerHTML = '<div class="session-card"><div class="session-body" style="text-align:center; color:#6b7280; padding:30px;">No upcoming training sessions currently scheduled.</div></div>';
        } else {
          upcomingEl.innerHTML = upcoming.map(s => renderSessionCard(s, false)).join('');
        }

        if (past.length > 0) {
          pastWrap.style.display = 'block';
          document.getElementById('togglePastBtn').textContent =
            (PAST_OPEN ? '🔼 Hide Past Training Sessions (' : '🕓 Show Past Training Sessions (') + past.length + ')';
          pastEl.style.display = PAST_OPEN ? 'block' : 'none';
          pastEl.innerHTML = past.map(s => renderSessionCard(s, true)).join('');
        } else {
          pastWrap.style.display = 'none';
        }
      }

      function togglePastSessions() {
        PAST_OPEN = !PAST_OPEN;
        renderAll();
      }

      function renderSessionCard(s, isReadOnly) {
        const isSlotted = s.signup.mode === 'slotted';
        let badgeHtml = isSlotted
          ? '<span class="badge badge-slotted">⏱️ Time-Slotted</span>'
          : '<span class="badge badge-group">👥 Group Session</span>';

        if (isSlotted) {
          const totalSpots = s.signup.slots.reduce((acc, sl) => acc + sl.capacity, 0);
          const filledSpots = s.signup.slots.reduce((acc, sl) => acc + sl.entries.length, 0);
          const openSpots = Math.max(0, totalSpots - filledSpots);
          badgeHtml += openSpots > 0
            ? '<span class="badge badge-open">' + openSpots + ' of ' + totalSpots + ' spots open</span>'
            : '<span class="badge badge-full">All ' + totalSpots + ' slots full</span>';
        } else {
          const count = s.signup.entries.length;
          const cap = Number(s.signup.capacity) || 0;
          if (cap > 0) {
            const openSpots = Math.max(0, cap - count);
            badgeHtml += openSpots > 0
              ? '<span class="badge badge-open">' + openSpots + ' of ' + cap + ' spots open</span>'
              : '<span class="badge badge-full">All ' + cap + ' spots full</span>';
          } else {
            badgeHtml += '<span class="badge badge-open">' + count + ' signed up</span>';
          }
        }
        if (!isReadOnly) {
          badgeHtml += '<button class="btn btn-outline btn-sm" style="margin-left:auto;" onclick="openManageModal(' + s.rowNumber + ')">⚙️ Manage</button>';
        }

        let metaHtml = '';
        if (s.timeText) metaHtml += '<div class="meta-item">🕒 <strong>Time:</strong> ' + escapeHtmlClient(s.timeText) + '</div>';
        if (s.locationText) metaHtml += '<div class="meta-item">📍 <strong>Location:</strong> ' + escapeHtmlClient(s.locationText) + '</div>';
        if (s.rsvpText) metaHtml += '<div class="meta-item">📅 <strong>RSVP By:</strong> ' + escapeHtmlClient(s.rsvpText) + '</div>';
        if (s.coordText) metaHtml += '<div class="meta-item">📣 <strong>Coach:</strong> ' + escapeHtmlClient(s.coordText) + '</div>';

        let notesHtml = '';
        if (s.notesText) {
          notesHtml += '<div class="details-box"><strong>Workout Notes:</strong>\\n' + escapeHtmlClient(s.notesText) + '</div>';
        }
        if (s.costText) {
          notesHtml += '<details style="margin-top:8px; font-size:12px; color:#6b7280;"><summary style="cursor:pointer; font-weight:600;">💲 View Cost &amp; Registration Info</summary><div style="margin-top:6px; padding:8px; background:#f9fafb; border-radius:6px;">' + escapeHtmlClient(s.costText) + '</div></details>';
        }

        // Member Sign-Up Body
        let memberHtml = '';
        if (isSlotted) {
          const openSlotsCount = s.signup.slots.filter(sl => sl.entries.length < sl.capacity).length;
          memberHtml += '<div class="section-label"><span>Member Time Slots (' + escapeHtmlClient(s.signup.banner || '2 per slot') + ')</span>';
          if (!isReadOnly && openSlotsCount > 0) {
            memberHtml += '<button class="btn btn-primary btn-sm" onclick="openSignupModal(' + s.rowNumber + ', \\'member\\', null)">+ Sign Up for a Slot</button>';
          }
          memberHtml += '</div><div class="slots-grid">';

          s.signup.slots.forEach((sl, slotIdx) => {
            const isFull = sl.entries.length >= sl.capacity;
            memberHtml += '<div class="slot-card ' + (isFull ? 'slot-full' : 'slot-open') + '">';
            memberHtml += '<div class="slot-top"><span class="slot-time">' + escapeHtmlClient(sl.time) + '</span>';
            memberHtml += '<span class="slot-cap">' + sl.entries.length + '/' + sl.capacity + '</span></div>';

            for (let i = 0; i < sl.capacity; i++) {
              const entry = sl.entries[i];
              if (entry) {
                memberHtml += '<div class="spot-row"><span class="spot-name">' + (i + 1) + '. ' + escapeHtmlClient(entry) + '</span>';
                if (!isReadOnly) {
                  memberHtml += '<button class="remove-btn" title="Remove signup" onclick="confirmDeregisterByIdx(' +
                    s.rowNumber + ', \\'member\\', ' + slotIdx + ', ' + i + ')">✕</button>';
                }
                memberHtml += '</div>';
              } else {
                memberHtml += '<div class="spot-row empty"><span>' + (i + 1) + '. Open</span>';
                if (!isReadOnly) {
                  memberHtml += '<button class="btn btn-primary btn-claim" onclick="openSignupModalBySlotIdx(' +
                    s.rowNumber + ', ' + slotIdx + ')">+ Claim</button>';
                }
                memberHtml += '</div>';
              }
            }
            memberHtml += '</div>';
          });
          memberHtml += '</div>';
        } else {
          const groupCap = Number(s.signup.capacity) || 0;
          const groupCount = s.signup.entries.length;
          const isGroupFull = groupCap > 0 && groupCount >= groupCap;
          const labelCount = groupCap > 0 ? (groupCount + ' / ' + groupCap) : String(groupCount);
          memberHtml += '<div class="section-label"><span>Member Sign-Up (' + labelCount + ')</span>';
          if (!isReadOnly && !isGroupFull) {
            memberHtml += '<button class="btn btn-primary btn-sm" onclick="openSignupModal(' + s.rowNumber + ', \\'member\\', null)">+ Member Sign Up</button>';
          }
          memberHtml += '</div>';

          if (s.signup.entries.length === 0) {
            memberHtml += '<div style="font-size:13px; color:#9ca3af; margin-bottom:14px; font-style:italic;">No member sign-ups yet. Be the first!</div>';
          } else {
            memberHtml += '<div class="signup-chip-list">';
            s.signup.entries.forEach((entry, idx) => {
              memberHtml += '<div class="signup-item"><span class="spot-name">' + (idx + 1) + '. ' + escapeHtmlClient(entry) + '</span>';
              if (!isReadOnly) {
                memberHtml += '<button class="remove-btn" title="Remove signup" onclick="confirmDeregisterByIdx(' +
                  s.rowNumber + ', \\'member\\', -1, ' + idx + ')">✕</button>';
              }
              memberHtml += '</div>';
            });
            memberHtml += '</div>';
          }
        }

        // Non-Member Waitlist Section
        let waitlistHtml = '<div class="waitlist-section"><div class="section-label"><span>Non-Member / Waitlist (' + s.waitlist.length + ')</span>';
        if (!isReadOnly) {
          waitlistHtml += '<button class="btn btn-outline btn-sm" onclick="openSignupModal(' + s.rowNumber + ', \\'waitlist\\', null)">+ Join Non-Member / Waitlist</button>';
        }
        waitlistHtml += '</div>';

        if (s.waitlist.length === 0) {
          waitlistHtml += '<div style="font-size:13px; color:#9ca3af; font-style:italic;">No non-member / waitlist entries.</div>';
        } else {
          waitlistHtml += '<div class="signup-chip-list">';
          s.waitlist.forEach((entry, idx) => {
            waitlistHtml += '<div class="signup-item"><span class="spot-name">' + (idx + 1) + '. ' + escapeHtmlClient(entry) + '</span>';
            if (!isReadOnly) {
              waitlistHtml += '<button class="remove-btn" title="Remove from waitlist" onclick="confirmDeregisterByIdx(' +
                s.rowNumber + ', \\'waitlist\\', -1, ' + idx + ')">✕</button>';
            }
            waitlistHtml += '</div>';
          });
          waitlistHtml += '</div>';
        }
        waitlistHtml += '</div>';

        return '<div class="session-card">' +
          '<div class="session-header">' +
            '<div class="session-badges">' + badgeHtml + '</div>' +
            '<h3 class="session-title">' + escapeHtmlClient(s.dateText || s.eventText || 'Training Session') + '</h3>' +
            (s.eventText ? '<div class="session-subtitle">' + escapeHtmlClient(s.eventText) + '</div>' : '') +
            '<div class="meta-grid">' + metaHtml + '</div>' +
            notesHtml +
          '</div>' +
          '<div class="session-body">' + memberHtml + waitlistHtml + '</div>' +
        '</div>';
      }

      function populateRosterSelect(selectId) {
        const sel = document.getElementById(selectId);
        const roster = APP_STATE.roster || [];
        let html = '<option value="" disabled selected>-- Select name --</option>';
        roster.forEach(name => {
          html += '<option value="' + escapeHtmlClient(name) + '">' + escapeHtmlClient(name) + '</option>';
        });
        html += '<option value="__CUSTOM__">➕ Other / Type name...</option>';
        sel.innerHTML = html;
      }

      function toggleCustomInput(selectId, inputId) {
        const val = document.getElementById(selectId).value;
        const inp = document.getElementById(inputId);
        if (val === '__CUSTOM__') {
          inp.style.display = 'block';
          inp.focus();
        } else {
          inp.style.display = 'none';
        }
      }

      function setPairMode(mode) {
        MODAL_CTX.pairMode = mode;
        document.getElementById('btnModeSolo').classList.toggle('active', mode === 'solo');
        document.getElementById('btnModePair').classList.toggle('active', mode === 'pair');
        document.getElementById('athlete1Label').textContent = mode === 'pair' ? 'Athlete 1' : 'Athlete Name';
        document.getElementById('athlete2Field').style.display = mode === 'pair' ? 'block' : 'none';
      }

      function openSignupModalBySlotIdx(rowNumber, slotIdx) {
        const session = (APP_STATE.sessions || []).find(s => s.rowNumber === rowNumber);
        const slot = session && session.signup && session.signup.slots && session.signup.slots[slotIdx];
        openSignupModal(rowNumber, 'member', slot ? slot.time : null);
      }

      function confirmDeregisterByIdx(rowNumber, listType, slotIdx, entryIdx) {
        const session = (APP_STATE.sessions || []).find(s => s.rowNumber === rowNumber);
        if (!session) return;
        let slotTime = null;
        let entryText = '';
        if (listType === 'waitlist') {
          entryText = (session.waitlist || [])[entryIdx] || '';
        } else if (session.signup.mode === 'slotted' && slotIdx >= 0) {
          const slot = (session.signup.slots || [])[slotIdx];
          if (!slot) return;
          slotTime = slot.time;
          entryText = (slot.entries || [])[entryIdx] || '';
        } else {
          entryText = (session.signup.entries || [])[entryIdx] || '';
        }
        confirmDeregister(rowNumber, listType, slotTime, entryIdx, entryText);
      }

      function openSignupModal(rowNumber, listType, preselectedSlot) {
        const session = (APP_STATE.sessions || []).find(s => s.rowNumber === rowNumber);
        if (!session) return;

        MODAL_CTX.rowNumber = rowNumber;
        MODAL_CTX.sessionKey = session.dateText;
        MODAL_CTX.listType = listType;
        MODAL_CTX.mode = session.signup.mode;
        MODAL_CTX.slotTime = preselectedSlot;

        document.getElementById('modalTitle').textContent =
          listType === 'waitlist' ? 'Join Non-Member / Waitlist' : 'Member Sign-Up';
        document.getElementById('modalSubtitle').textContent =
          (session.dateText || '') + (session.eventText ? ' — ' + session.eventText : '');

        populateRosterSelect('athlete1Select');
        populateRosterSelect('athlete2Select');
        document.getElementById('athlete1Custom').style.display = 'none';
        document.getElementById('athlete1Custom').value = '';
        document.getElementById('athlete2Custom').style.display = 'none';
        document.getElementById('athlete2Custom').value = '';
        document.getElementById('signupNoteInput').value = '';
        document.getElementById('noteField').style.display = listType === 'waitlist' ? 'block' : 'none';

        // If time-slotted member signup, show slot dropdown
        const slotField = document.getElementById('slotPickerField');
        if (listType === 'member' && session.signup.mode === 'slotted') {
          slotField.style.display = 'block';
          const slotSel = document.getElementById('modalSlotSelect');
          const openSlots = session.signup.slots.filter(sl => sl.entries.length < sl.capacity);
          slotSel.innerHTML = openSlots.map(sl => {
            const rem = sl.capacity - sl.entries.length;
            const selected = (preselectedSlot && sl.time === preselectedSlot) ? ' selected' : '';
            return '<option value="' + escapeHtmlClient(sl.time) + '"' + selected + '>' +
              escapeHtmlClient(sl.time) + ' (' + rem + ' spot' + (rem > 1 ? 's' : '') + ' open)</option>';
          }).join('');
          slotSel.value = preselectedSlot || (openSlots[0] ? openSlots[0].time : '');
        } else {
          slotField.style.display = 'none';
        }

        setPairMode('solo');
        document.getElementById('signupModalBackdrop').classList.add('open');
      }

      function closeSignupModal() {
        document.getElementById('signupModalBackdrop').classList.remove('open');
      }

      function resolveSelectedName(selectId, customId) {
        const selVal = document.getElementById(selectId).value;
        if (!selVal) return '';
        if (selVal === '__CUSTOM__') {
          return document.getElementById(customId).value.trim();
        }
        return selVal.trim();
      }

      function submitSignup() {
        const name1 = resolveSelectedName('athlete1Select', 'athlete1Custom');
        if (!name1) {
          showToast('⚠️ Please select or enter your name.');
          return;
        }
        let finalEntry = name1;
        if (MODAL_CTX.pairMode === 'pair') {
          const name2 = resolveSelectedName('athlete2Select', 'athlete2Custom');
          if (!name2) {
            showToast('⚠️ Please select or enter your partner\\'s name.');
            return;
          }
          finalEntry = name1 + ' & ' + name2;
        }

        if (MODAL_CTX.listType === 'waitlist') {
          const note = document.getElementById('signupNoteInput').value.trim();
          if (note) finalEntry += ' (' + note + ')';
        }

        const slotTime = (MODAL_CTX.listType === 'member' && MODAL_CTX.mode === 'slotted')
          ? document.getElementById('modalSlotSelect').value
          : null;

        const btn = document.getElementById('confirmSignupBtn');
        btn.disabled = true;
        btn.textContent = 'Saving...';

        const payload = {
          rowNumber: MODAL_CTX.rowNumber,
          sessionKey: MODAL_CTX.sessionKey,
          listType: MODAL_CTX.listType,
          slotTime: slotTime,
          entryText: finalEntry,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Confirm Sign-Up';
            closeSignupModal();
            APP_STATE = newState;
            renderAll();
            showToast('✅ Signed up: ' + finalEntry);
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Confirm Sign-Up';
            showToast('❌ Sign-up failed: ' + (err.message || err));
            refreshSessions();
          })
          .registerTrainingSignup(payload);
      }

      function confirmDeregister(rowNumber, listType, slotTime, entryIndex, entryText) {
        const where = slotTime ? (' from ' + slotTime) : (listType === 'waitlist' ? ' from the waitlist' : '');
        REMOVE_CTX = {
          rowNumber: rowNumber,
          listType: listType,
          slotTime: slotTime,
          entryIndex: entryIndex,
          entryText: entryText
        };
        document.getElementById('removeModalText').textContent =
          'Remove "' + entryText + '"' + where + '?';
        const btn = document.getElementById('confirmRemoveBtn');
        btn.disabled = false;
        btn.textContent = 'Remove';
        document.getElementById('removeModalBackdrop').classList.add('open');
      }

      function closeRemoveModal() {
        document.getElementById('removeModalBackdrop').classList.remove('open');
        REMOVE_CTX = null;
      }

      function executeDeregister() {
        if (!REMOVE_CTX) return;
        const ctx = REMOVE_CTX;
        const btn = document.getElementById('confirmRemoveBtn');
        btn.disabled = true;
        btn.textContent = 'Removing...';

        const session = (APP_STATE.sessions || []).find(s => s.rowNumber === ctx.rowNumber);
        const payload = {
          rowNumber: ctx.rowNumber,
          sessionKey: session ? session.dateText : '',
          listType: ctx.listType,
          slotTime: ctx.slotTime,
          entryIndex: ctx.entryIndex,
          entryText: ctx.entryText,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Remove';
            closeRemoveModal();
            APP_STATE = newState;
            renderAll();
            showToast('🗑️ Removed ' + ctx.entryText);
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Remove';
            closeRemoveModal();
            showToast('❌ Could not remove signup: ' + (err.message || err));
            refreshSessions();
          })
          .deregisterTrainingSignup(payload);
      }

      function refreshSessions() {
        showToast('Refreshing from Google Sheet...');
        google.script.run
          .withSuccessHandler(newState => {
            APP_STATE = newState;
            renderAll();
            showToast('Up to date!');
          })
          .withFailureHandler(err => {
            showToast('❌ Error refreshing: ' + (err.message || err));
          })
          .getTrainingAppState(Boolean(APP_STATE.useTestSheet), getActiveCoachEmail());
      }

      // Coach Sign-In & Access Control Logic
      function openCoachAuthModal() {
        const errEl = document.getElementById('coachAuthError');
        errEl.style.display = 'none';
        errEl.textContent = '';
        const inp = document.getElementById('coachEmailInput');
        inp.value = CLIENT_COACH_EMAIL || '';
        document.getElementById('coachAuthModalBackdrop').classList.add('open');
        inp.focus();
      }

      function closeCoachAuthModal() {
        document.getElementById('coachAuthModalBackdrop').classList.remove('open');
      }

      function submitCoachAuth() {
        const email = document.getElementById('coachEmailInput').value.trim().toLowerCase();
        const errEl = document.getElementById('coachAuthError');
        if (!email || !email.includes('@')) {
          errEl.textContent = 'Please enter a valid Coach email address.';
          errEl.style.display = 'block';
          return;
        }
        errEl.style.display = 'none';
        const btn = document.getElementById('confirmCoachAuthBtn');
        btn.disabled = true;
        btn.textContent = 'Verifying...';

        google.script.run
          .withSuccessHandler(res => {
            btn.disabled = false;
            btn.textContent = 'Unlock Coach Mode';
            if (res && res.isCoach && res.state) {
              CLIENT_COACH_EMAIL = email;
              try {
                if (typeof localStorage !== 'undefined' && localStorage) {
                  localStorage.setItem('hyrox_coach_email', email);
                }
              } catch (e) {}
              APP_STATE = res.state;
              closeCoachAuthModal();
              renderAll();
              showToast('🔓 Coach Mode unlocked!');
            } else {
              errEl.textContent = 'That email is not on the authorized Coach list.';
              errEl.style.display = 'block';
              showToast('❌ Unauthorized email for Coach Mode.');
            }
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Unlock Coach Mode';
            errEl.textContent = err.message || String(err);
            errEl.style.display = 'block';
          })
          .verifyCoachEmail(email, Boolean(APP_STATE.useTestSheet));
      }

      function signOutCoachMode() {
        CLIENT_COACH_EMAIL = '';
        try {
          if (typeof localStorage !== 'undefined' && localStorage) {
            localStorage.removeItem('hyrox_coach_email');
          }
        } catch (e) {}
        APP_STATE.isCoach = false;
        APP_STATE.detectedEmail = '';
        renderAll();
        showToast('🔒 Exited Coach Mode');
      }

      // Coach Add Session Modal Logic
      function openCoachModal() {
        document.getElementById('coachModalBackdrop').classList.add('open');
      }
      function closeCoachModal() {
        document.getElementById('coachModalBackdrop').classList.remove('open');
      }
      function setCoachSessionMode(mode) {
        COACH_MODE = mode;
        document.getElementById('coachModeGroup').classList.toggle('active', mode === 'group');
        document.getElementById('coachModeSlotted').classList.toggle('active', mode === 'slotted');
        document.getElementById('coachGroupConfig').style.display = mode === 'group' ? 'block' : 'none';
        document.getElementById('coachSlotsConfig').style.display = mode === 'slotted' ? 'block' : 'none';
      }
      function submitCoachNewSession() {
        const dateText = document.getElementById('coachDate').value.trim();
        if (!dateText) {
          showToast('⚠️ Please enter the Date Header (e.g. TRAINING ON: Sunday 11/22/26)');
          return;
        }
        const btn = document.getElementById('coachCreateBtn');
        btn.disabled = true;
        btn.textContent = 'Creating...';

        const payload = {
          dateText: dateText,
          eventText: document.getElementById('coachEvent').value.trim(),
          timeText: document.getElementById('coachTime').value.trim(),
          rsvpText: document.getElementById('coachRsvp').value.trim(),
          locationText: document.getElementById('coachLocation').value.trim() || 'F45',
          coordText: document.getElementById('coachCoord').value.trim() || 'Jade',
          notesText: document.getElementById('coachNotes').value.trim(),
          mode: COACH_MODE,
          groupCap: Math.max(0, Number(document.getElementById('coachGroupCap').value) || 0),
          firstSlot: document.getElementById('coachFirstSlot').value.trim(),
          intervalMins: Number(document.getElementById('coachIntervalMins').value) || 8,
          numSlots: Number(document.getElementById('coachNumSlots').value) || 12,
          slotCap: Number(document.getElementById('coachSlotCap').value) || 2,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Create Session';
            closeCoachModal();
            APP_STATE = newState;
            renderAll();
            showToast('✅ Created new training session!');
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Create Session';
            showToast('❌ Failed to create session: ' + (err.message || err));
          })
          .coachCreateTrainingSession(payload);
      }

      // Coach Manage Existing Session Modal Logic
      function parseSlotMinsClient(timeStr) {
        const m = String(timeStr || '').trim().match(/^(\\d{1,2}):(\\d{2})\\s*(AM|PM)?$/i);
        if (!m) return null;
        let h = parseInt(m[1], 10);
        const min = parseInt(m[2], 10);
        const mer = (m[3] || 'AM').toUpperCase();
        if (mer === 'PM' && h < 12) h += 12;
        if (mer === 'AM' && h === 12) h = 0;
        return h * 60 + min;
      }

      function formatSlotMinsClient(totalMins) {
        const h24 = Math.floor(totalMins / 60) % 24;
        const minPart = totalMins % 60;
        const ampm = h24 >= 12 ? 'PM' : 'AM';
        const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
        return h12 + ':' + (minPart < 10 ? '0' + minPart : minPart) + ' ' + ampm;
      }

      function openManageModal(rowNumber) {
        const session = (APP_STATE.sessions || []).find(s => s.rowNumber === rowNumber);
        if (!session) return;

        MANAGE_CTX.rowNumber = rowNumber;
        MANAGE_CTX.sessionKey = session.dateText || '';

        document.getElementById('manageModalSubtitle').textContent =
          (session.dateText || '') + (session.eventText ? ' — ' + session.eventText : '');

        document.getElementById('manageDate').value = session.dateText || '';
        document.getElementById('manageEvent').value = session.eventText || '';
        document.getElementById('manageTime').value = session.timeText || '';
        document.getElementById('manageRsvp').value = session.rsvpText || '';
        document.getElementById('manageLocation').value = session.locationText || '';
        document.getElementById('manageCoord').value = session.coordText || '';
        document.getElementById('manageNotes').value = session.notesText || '';
        document.getElementById('manageCost').value = session.costText || '';

        const isSlotted = session.signup.mode === 'slotted';
        let firstSlot = '10:30 AM';
        let intervalMins = 8;
        let numSlots = 12;
        let slotCap = 2;
        let minSlots = 1;
        let minSlotCap = 1;
        let groupCap = 20;
        let minGroupCap = 0;

        if (isSlotted) {
          const slots = session.signup.slots || [];
          if (slots.length > 0) {
            firstSlot = slots[0].time || '10:30 AM';
            numSlots = slots.length;
            slotCap = slots.reduce((m, sl) => Math.max(m, sl.capacity || 2), 2);
          }
          if (slots.length >= 2) {
            const t0 = parseSlotMinsClient(slots[0].time);
            const t1 = parseSlotMinsClient(slots[1].time);
            if (t0 !== null && t1 !== null && t1 > t0 && (t1 - t0) <= 180) {
              intervalMins = t1 - t0;
            }
          }
          slots.forEach((sl, idx) => {
            if ((sl.entries || []).length > 0) {
              minSlots = Math.max(minSlots, idx + 1);
              minSlotCap = Math.max(minSlotCap, sl.entries.length);
            }
          });
          minGroupCap = slots.reduce((acc, sl) => acc + (sl.entries || []).length, 0);
          groupCap = Math.max(20, minGroupCap);
        } else {
          const entries = session.signup.entries || [];
          minGroupCap = entries.length;
          groupCap = Number(session.signup.capacity) || Math.max(20, minGroupCap);
          minSlots = Math.max(1, Math.ceil(minGroupCap / 2));
        }

        MANAGE_CTX.minSlots = minSlots;
        MANAGE_CTX.minSlotCap = minSlotCap;
        MANAGE_CTX.minGroupCap = minGroupCap;

        document.getElementById('manageGroupCap').value = groupCap;
        document.getElementById('manageFirstSlot').value = firstSlot;
        document.getElementById('manageIntervalMins').value = intervalMins;
        document.getElementById('manageNumSlots').value = numSlots;
        document.getElementById('manageSlotCap').value = slotCap;

        setManageSessionMode(isSlotted ? 'slotted' : 'group');
        document.getElementById('manageModalBackdrop').classList.add('open');
      }

      function closeManageModal() {
        document.getElementById('manageModalBackdrop').classList.remove('open');
      }

      function setManageSessionMode(mode) {
        MANAGE_CTX.mode = mode;
        document.getElementById('manageModeGroup').classList.toggle('active', mode === 'group');
        document.getElementById('manageModeSlotted').classList.toggle('active', mode === 'slotted');
        document.getElementById('manageGroupConfig').style.display = mode === 'group' ? 'block' : 'none';
        document.getElementById('manageSlotsConfig').style.display = mode === 'slotted' ? 'block' : 'none';
        updateManagePreview();
      }

      function adjustManageSlots(mode, delta) {
        if (mode === 'group') {
          const inp = document.getElementById('manageGroupCap');
          const cur = Number(inp.value) || 0;
          inp.value = Math.max(MANAGE_CTX.minGroupCap, cur + delta);
        } else {
          const inp = document.getElementById('manageNumSlots');
          const cur = Number(inp.value) || 1;
          inp.value = Math.min(40, Math.max(MANAGE_CTX.minSlots, cur + delta));
        }
        updateManagePreview();
      }

      function updateManagePreview() {
        if (MANAGE_CTX.mode === 'group') {
          const cap = Number(document.getElementById('manageGroupCap').value) || 0;
          const hint = document.getElementById('manageGroupHint');
          hint.textContent = cap > 0
            ? ('Currently ' + MANAGE_CTX.minGroupCap + ' signed up • ' + Math.max(0, cap - MANAGE_CTX.minGroupCap) + ' open spots after saving.')
            : ('Currently ' + MANAGE_CTX.minGroupCap + ' signed up • Unlimited member spots.');
        } else {
          const firstStr = document.getElementById('manageFirstSlot').value.trim() || '10:30 AM';
          const step = Math.max(1, Number(document.getElementById('manageIntervalMins').value) || 8);
          const count = Math.max(1, Number(document.getElementById('manageNumSlots').value) || 12);
          const cap = Math.max(1, Number(document.getElementById('manageSlotCap').value) || 2);
          const startMins = parseSlotMinsClient(firstStr);
          const prevEl = document.getElementById('manageSchedulePreview');
          if (startMins !== null) {
            const firstFmt = formatSlotMinsClient(startMins);
            const lastFmt = formatSlotMinsClient(startMins + (count - 1) * step);
            prevEl.textContent = '📅 ' + count + ' slots from ' + firstFmt + ' to ' + lastFmt +
              ' (every ' + step + 'm, ' + (count * cap) + ' total spots). Existing sign-ups stay in their slot order.';
          } else {
            prevEl.textContent = '⚠️ Enter First Slot as HH:MM AM/PM (e.g. 10:30 AM).';
          }
        }
      }

      function submitManageSession() {
        const dateText = document.getElementById('manageDate').value.trim();
        if (!dateText) {
          showToast('⚠️ Date Header cannot be empty.');
          return;
        }
        const mode = MANAGE_CTX.mode;
        const groupCap = Math.max(0, Number(document.getElementById('manageGroupCap').value) || 0);
        const firstSlot = document.getElementById('manageFirstSlot').value.trim();
        const intervalMins = Math.max(1, Number(document.getElementById('manageIntervalMins').value) || 8);
        const numSlots = Math.max(1, Number(document.getElementById('manageNumSlots').value) || 12);
        const slotCap = Math.max(1, Number(document.getElementById('manageSlotCap').value) || 2);

        if (mode === 'group' && groupCap > 0 && groupCap < MANAGE_CTX.minGroupCap) {
          showToast('⚠️ Cannot set slots below ' + MANAGE_CTX.minGroupCap + ' (already signed up).');
          return;
        }
        if (mode === 'slotted') {
          if (parseSlotMinsClient(firstSlot) === null) {
            showToast('⚠️ Please enter a valid First Slot time (e.g. 10:30 AM).');
            return;
          }
          if (numSlots < MANAGE_CTX.minSlots) {
            showToast('⚠️ Cannot reduce below ' + MANAGE_CTX.minSlots + ' slots (Slot #' + MANAGE_CTX.minSlots + ' has sign-ups).');
            return;
          }
          if (slotCap < MANAGE_CTX.minSlotCap) {
            showToast('⚠️ Cannot reduce per-slot capacity below ' + MANAGE_CTX.minSlotCap + '.');
            return;
          }
        }

        const btn = document.getElementById('manageSaveBtn');
        btn.disabled = true;
        btn.textContent = 'Saving...';

        const payload = {
          rowNumber: MANAGE_CTX.rowNumber,
          sessionKey: MANAGE_CTX.sessionKey,
          dateText: dateText,
          eventText: document.getElementById('manageEvent').value.trim(),
          timeText: document.getElementById('manageTime').value.trim(),
          rsvpText: document.getElementById('manageRsvp').value.trim(),
          locationText: document.getElementById('manageLocation').value.trim(),
          coordText: document.getElementById('manageCoord').value.trim(),
          notesText: document.getElementById('manageNotes').value.trim(),
          costText: document.getElementById('manageCost').value.trim(),
          mode: mode,
          groupCap: groupCap,
          firstSlot: firstSlot,
          intervalMins: intervalMins,
          numSlots: numSlots,
          slotCap: slotCap,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Save Changes';
            closeManageModal();
            APP_STATE = newState;
            renderAll();
            showToast('✅ Updated session details & schedule!');
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Save Changes';
            showToast('❌ Update failed: ' + (err.message || err));
            refreshSessions();
          })
          .coachUpdateTrainingSession(payload);
      }

      // Coach Delete / Cancel Session Modal Logic
      function openDeleteSessionModal() {
        const session = (APP_STATE.sessions || []).find(s => s.rowNumber === MANAGE_CTX.rowNumber);
        const label = session
          ? ((session.dateText || '') + (session.eventText ? ' — ' + session.eventText : ''))
          : (MANAGE_CTX.sessionKey || 'this session');
        document.getElementById('deleteSessionModalText').textContent =
          'Are you sure you want to permanently delete "' + label + '" and all of its sign-ups from the TRAINING SESSIONS sheet?';
        const btn = document.getElementById('confirmDeleteSessionBtn');
        btn.disabled = false;
        btn.textContent = 'Yes, Delete Session';
        document.getElementById('deleteSessionModalBackdrop').classList.add('open');
      }

      function closeDeleteSessionModal() {
        document.getElementById('deleteSessionModalBackdrop').classList.remove('open');
      }

      function executeDeleteSession() {
        const btn = document.getElementById('confirmDeleteSessionBtn');
        btn.disabled = true;
        btn.textContent = 'Deleting...';

        const payload = {
          rowNumber: MANAGE_CTX.rowNumber,
          sessionKey: MANAGE_CTX.sessionKey,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Yes, Delete Session';
            closeDeleteSessionModal();
            closeManageModal();
            APP_STATE = newState;
            renderAll();
            showToast('🗑️ Deleted training session.');
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Yes, Delete Session';
            closeDeleteSessionModal();
            showToast('❌ Could not delete session: ' + (err.message || err));
          })
          .coachDeleteTrainingSession(payload);
      }

      // Strava Tab Helpers
      function toggleCustomStravaName(val) {
        const div = document.getElementById('customNameDiv');
        const customInput = document.getElementById('customNameInput');
        if (val === '__NEW__') {
          div.style.display = 'block';
          customInput.focus();
        } else {
          div.style.display = 'none';
          document.getElementById('stateInput').value = val;
        }
      }

      function validateStravaForm() {
        const selectVal = document.getElementById('nameSelect').value;
        if (!selectVal) {
          showToast('⚠️ Please select your name or choose "Not listed"');
          return false;
        }
        let finalName = selectVal;
        if (selectVal === '__NEW__') {
          finalName = document.getElementById('customNameInput').value.trim();
          if (!finalName) {
            showToast('⚠️ Please type your name');
            return false;
          }
        }
        document.getElementById('stateInput').value = finalName;

        const authUrl = "https://www.strava.com/oauth/authorize?client_id=" + encodeURIComponent("${escapeHtml(clientId)}") +
          "&response_type=code" +
          "&redirect_uri=" + encodeURIComponent("${escapeHtml(scriptUrl)}") +
          "&approval_prompt=auto" +
          "&scope=read,activity:read_all" +
          "&state=" + encodeURIComponent(finalName);

        try {
          window.top.location.href = authUrl;
          return false;
        } catch (e) {
          window.open(authUrl, '_top');
          return false;
        }
      }

      // HYROX ANAHEIM 2026 RACE HUB LOGIC
      let RACE_SUBTAB = 'waves';
      let RACE_DAY_FILTER = 'ALL';

      const DAY_META = {
        'Thursday': { short: 'Thu 12/3', full: 'Thursday, Dec 3' },
        'Friday': { short: 'Fri 12/4', full: 'Friday, Dec 4' },
        'Saturday': { short: 'Sat 12/5', full: 'Saturday, Dec 5' },
        'Sunday': { short: 'Sun 12/6', full: 'Sunday, Dec 6' }
      };

      const DINNER_SHORT_LABELS = {
        'Thurs Dinner': '🍽️ Thu 12/3',
        'Friday Dinner': '🍽️ Fri 12/4',
        'Saturday Dinner': '🍽️ Sat 12/5',
        'Sunday Dinner': '🍽️ Sun 12/6',
        'Monday Dinner': '🍽️ Mon 12/7'
      };

      const DEFAULT_HOTEL_CHIPS = [
        'BW Lulu Anaheim',
        'Hilton Anaheim',
        'Sheraton Anaheim',
        'Residence Inn Anaheim',
        "King's Inn Anaheim",
        'Westin Anaheim'
      ];

      let WAVE_MODAL_CTX = {
        rowNumber: null,
        originalName: '',
        originalDay: '',
        timeMode: 'set'
      };

      let TRAVEL_MODAL_CTX = {
        rowNumber: null,
        originalName: '',
        dinners: []
      };

      function switchRaceSubtab(subtab) {
        RACE_SUBTAB = subtab === 'travel' ? 'travel' : 'waves';
        const btnWaves = document.getElementById('subtabBtnWaves');
        const btnTravel = document.getElementById('subtabBtnTravel');
        const viewWaves = document.getElementById('raceWavesView');
        const viewTravel = document.getElementById('raceTravelView');
        if (btnWaves) btnWaves.classList.toggle('active', RACE_SUBTAB === 'waves');
        if (btnTravel) btnTravel.classList.toggle('active', RACE_SUBTAB === 'travel');
        if (viewWaves) viewWaves.style.display = RACE_SUBTAB === 'waves' ? 'block' : 'none';
        if (viewTravel) viewTravel.style.display = RACE_SUBTAB === 'travel' ? 'block' : 'none';
      }

      function setRaceDayFilter(day) {
        RACE_DAY_FILTER = day || 'ALL';
        renderRaceWaves();
      }

      function populateClubDatalist() {
        const dl = document.getElementById('clubRosterDatalist');
        if (!dl) return;
        const nameSet = new Set(APP_STATE.roster || []);
        const waves = (APP_STATE.raceWaves && APP_STATE.raceWaves.entries) || [];
        waves.forEach(w => {
          if (w.name) nameSet.add(w.name);
          if (w.partner1) nameSet.add(w.partner1);
          if (w.partner2) nameSet.add(w.partner2);
          if (w.partner3) nameSet.add(w.partner3);
        });
        const sorted = Array.from(nameSet).sort((a, b) => a.localeCompare(b));
        dl.innerHTML = sorted.map(n => '<option value="' + escapeHtmlClient(n) + '"></option>').join('');
      }

      function renderRaceHub() {
        populateClubDatalist();
        const waveEntries = (APP_STATE.raceWaves && APP_STATE.raceWaves.entries) || [];
        const travelEntries = (APP_STATE.raceTravel && APP_STATE.raceTravel.entries) || [];

        const wavesBadge = document.getElementById('wavesCountBadge');
        if (wavesBadge) wavesBadge.textContent = String(waveEntries.length);

        const travelBadge = document.getElementById('travelCountBadge');
        if (travelBadge) travelBadge.textContent = String(travelEntries.length);

        renderRaceWaves();
        renderRaceTravel();
      }

      function formatWaveAthletesDisplay(entry) {
        const parts = [entry.name, entry.partner1, entry.partner2, entry.partner3].map(s => String(s || '').trim()).filter(Boolean);
        if (parts.length === 0) return 'TBD';
        if (parts.length === 1) return escapeHtmlClient(parts[0]);
        if (parts.length === 2) return escapeHtmlClient(parts[0]) + ' <span style="color:var(--brand); font-weight:800;">&amp;</span> ' + escapeHtmlClient(parts[1]);
        return parts.map(p => escapeHtmlClient(p)).join(', ');
      }

      function getDivisionPillClass(div) {
        const d = String(div || '').toLowerCase();
        if (d.includes('relay')) return 'div-pill relay';
        if (d.includes('pro')) return 'div-pill pro';
        if (d.includes('doubles')) return 'div-pill doubles';
        return 'div-pill';
      }

      function renderRaceWaves() {
        const pillsEl = document.getElementById('raceDayPills');
        const listEl = document.getElementById('raceWavesContainer');
        if (!pillsEl || !listEl) return;

        const waveData = APP_STATE.raceWaves || { entries: [], days: ['Thursday', 'Friday', 'Saturday', 'Sunday'] };
        const entries = waveData.entries || [];
        const days = waveData.days || ['Thursday', 'Friday', 'Saturday', 'Sunday'];

        const countsByDay = { ALL: entries.length };
        days.forEach(d => { countsByDay[d] = 0; });
        entries.forEach(e => {
          const d = e.day || 'Thursday';
          countsByDay[d] = (countsByDay[d] || 0) + 1;
        });

        let pillsHtml = '<button type="button" class="day-pill ' + (RACE_DAY_FILTER === 'ALL' ? 'active' : '') +
          '" onclick="setRaceDayFilter(\\'ALL\\')">All Days (' + countsByDay.ALL + ')</button>';
        days.forEach(d => {
          const meta = DAY_META[d] || { short: d, full: d };
          const cnt = countsByDay[d] || 0;
          pillsHtml += '<button type="button" class="day-pill ' + (RACE_DAY_FILTER === d ? 'active' : '') +
            '" onclick="setRaceDayFilter(\\'' + escapeHtmlClient(d) + '\\')">' +
            escapeHtmlClient(meta.short) + ' (' + cnt + ')</button>';
        });
        pillsEl.innerHTML = pillsHtml;

        const searchEl = document.getElementById('waveSearchInput');
        const query = searchEl ? String(searchEl.value || '').trim().toLowerCase() : '';
        const divFilterEl = document.getElementById('waveDivFilter');
        const divFilter = divFilterEl ? divFilterEl.value : 'ALL';

        const filtered = entries.filter(e => {
          if (RACE_DAY_FILTER !== 'ALL' && e.day !== RACE_DAY_FILTER) return false;
          if (divFilter === 'DOUBLES' && !/doubles/i.test(e.division) && (!e.partner1 || e.partner2)) return false;
          if (divFilter === 'RELAY' && !/relay/i.test(e.division) && !e.partner2) return false;
          if (divFilter === 'SINGLES' && (/doubles|relay/i.test(e.division) || e.partner1)) return false;
          if (query) {
            const hay = [e.name, e.partner1, e.partner2, e.partner3, e.division, e.startWave, e.day].join(' ').toLowerCase();
            if (!hay.includes(query)) return false;
          }
          return true;
        });

        if (filtered.length === 0) {
          listEl.innerHTML = '<div class="session-card"><div class="session-body" style="text-align:center; color:#6b7280; padding:28px;">No matching race wave entries found.</div></div>';
          return;
        }

        const activeDays = RACE_DAY_FILTER === 'ALL' ? days : [RACE_DAY_FILTER];
        let html = '';

        activeDays.forEach(d => {
          const dayEntries = filtered.filter(e => e.day === d);
          if (dayEntries.length === 0) return;

          const sorted = dayEntries.slice().sort((a, b) => {
            const ma = typeof a.startMins === 'number' ? a.startMins : 9999;
            const mb = typeof b.startMins === 'number' ? b.startMins : 9999;
            if (ma !== mb) return ma - mb;
            return (a.rowNumber || 0) - (b.rowNumber || 0);
          });

          const meta = DAY_META[d] || { short: d, full: d };
          const withTimeCount = sorted.filter(e => Boolean(e.startWave)).length;

          html += '<div class="race-day-card">';
          html += '<div class="race-day-header">';
          html += '<h3 class="race-day-title">📅 ' + escapeHtmlClient(meta.full) + '</h3>';
          html += '<span class="badge badge-slotted">' + sorted.length + (sorted.length === 1 ? ' entry' : ' entries') +
            (withTimeCount > 0 ? ' • ' + withTimeCount + ' w/ wave time' : '') + '</span>';
          html += '</div>';
          html += '<div class="wave-list">';

          sorted.forEach(e => {
            const hasWave = Boolean(e.startWave);
            const timeBtnHtml = hasWave
              ? '<button type="button" class="wave-time-badge" title="Update wave time" onclick="openWaveModal(' + e.rowNumber + ', true)">⏱️ ' + escapeHtmlClient(e.startWave) + '</button>'
              : '<button type="button" class="wave-time-badge unset" title="Add start wave time" onclick="openWaveModal(' + e.rowNumber + ', true)">+ Wave Time</button>';

            html += '<div class="wave-row">';
            html += '<div class="wave-left">';
            html += timeBtnHtml;
            html += '<div class="wave-info">';
            html += '<div class="wave-athletes">' + formatWaveAthletesDisplay(e) + '</div>';
            html += '<div class="wave-meta"><span class="' + getDivisionPillClass(e.division) + '">' + escapeHtmlClient(e.division || 'Open') + '</span></div>';
            html += '</div>';
            html += '</div>';
            html += '<button type="button" class="btn btn-outline btn-sm" onclick="openWaveModal(' + e.rowNumber + ', false)">✏️ Edit</button>';
            html += '</div>';
          });

          html += '</div></div>';
        });

        listEl.innerHTML = html;
      }

      function parseHotelDisplayAndUrl(rawLodging) {
        const str = String(rawLodging || '').trim();
        if (!str) return { name: '', url: '' };
        const urlMatch = str.match(/https?:\\/\\/\\S+/i);
        const url = urlMatch ? urlMatch[0] : '';
        const cleanName = str.replace(/https?:\\/\\/\\S+/gi, '').replace(/\\s+/g, ' ').trim();
        return {
          name: cleanName || (url ? 'Hotel Link' : ''),
          url: url
        };
      }

      function renderRaceTravel() {
        const summaryEl = document.getElementById('raceTravelSummaryContainer');
        const cardsEl = document.getElementById('raceTravelCardsContainer');
        if (!summaryEl || !cardsEl) return;

        const travelData = APP_STATE.raceTravel || {
          entries: [],
          outboundDays: ['Wednesday 12/2', 'Thursday 12/3', 'Friday 12/4', 'Saturday 12/5', 'Sunday 12/6'],
          returnDays: ['Thursday 12/3', 'Friday 12/4', 'Saturday 12/5', 'Sunday 12/6', 'Monday 12/7'],
          dinnerLabels: ['Thurs Dinner', 'Friday Dinner', 'Saturday Dinner', 'Sunday Dinner', 'Monday Dinner']
        };
        const entries = travelData.entries || [];
        const dinnerLabels = travelData.dinnerLabels || ['Thurs Dinner', 'Friday Dinner', 'Saturday Dinner', 'Sunday Dinner', 'Monday Dinner'];

        let dinnerHtml = '<div class="summary-card"><h4><span>🍽️ Group Dinner Interest</span><span class="badge badge-open">Dec 3–7</span></h4>';
        dinnerLabels.forEach(dl => {
          const interested = entries.filter(e => (e.dinners || []).some(d => d.toLowerCase() === dl.toLowerCase())).map(e => e.name);
          const shortLabel = DINNER_SHORT_LABELS[dl] || dl;
          dinnerHtml += '<div class="dinner-row">';
          dinnerHtml += '<div class="dinner-day-label">' + escapeHtmlClient(shortLabel) + ' <span class="badge ' + (interested.length > 0 ? 'badge-open' : '') + '" style="margin-left:4px;">' + interested.length + '</span></div>';
          dinnerHtml += '<div class="dinner-names">' + (interested.length > 0 ? interested.map(n => escapeHtmlClient(n)).join(', ') : '<span style="color:#9ca3af; font-style:italic;">No RSVPs yet</span>') + '</div>';
          dinnerHtml += '</div>';
        });
        dinnerHtml += '</div>';

        const hotelMap = {};
        entries.forEach(e => {
          const parsed = parseHotelDisplayAndUrl(e.lodging);
          if (!parsed.name || parsed.name.toLowerCase() === 'n/a') return;
          const key = parsed.name.toLowerCase();
          if (!hotelMap[key]) {
            hotelMap[key] = { name: parsed.name, url: parsed.url, guests: [] };
          } else if (!hotelMap[key].url && parsed.url) {
            hotelMap[key].url = parsed.url;
          }
          hotelMap[key].guests.push(e.name);
        });
        const hotelGroups = Object.values(hotelMap);

        let hotelHtml = '<div class="summary-card"><h4><span>🏨 Where People Are Staying</span><span class="badge badge-slotted">' + hotelGroups.length + (hotelGroups.length === 1 ? ' hotel' : ' hotels') + '</span></h4>';
        if (hotelGroups.length === 0) {
          hotelHtml += '<div style="font-size:13px; color:#9ca3af; font-style:italic; padding:8px 0;">No hotels listed yet. Add yours so others can book nearby!</div>';
        } else {
          hotelGroups.forEach(hg => {
            hotelHtml += '<div class="hotel-group-item">';
            hotelHtml += '<div style="display:flex; justify-content:space-between; align-items:center; gap:8px;">';
            hotelHtml += '<strong style="color:var(--navy);">' + escapeHtmlClient(hg.name) + '</strong>';
            if (hg.url) {
              hotelHtml += '<a href="' + escapeHtmlClient(hg.url) + '" target="_blank" rel="noopener" style="font-size:12px; color:var(--brand); font-weight:700; text-decoration:none;">📍 Map ↗</a>';
            }
            hotelHtml += '</div>';
            hotelHtml += '<div style="color:var(--muted); margin-top:2px;">👥 ' + hg.guests.map(g => escapeHtmlClient(g)).join(', ') + '</div>';
            hotelHtml += '</div>';
          });
        }
        hotelHtml += '</div>';

        summaryEl.innerHTML = dinnerHtml + hotelHtml;

        if (entries.length === 0) {
          cardsEl.innerHTML = '<div class="session-card"><div class="session-body" style="text-align:center; color:#6b7280; padding:28px;">No travel plans added yet. Tap <b>+ Add / Edit My Travel</b> to share yours!</div></div>';
          return;
        }

        let cardsHtml = '';
        entries.forEach(e => {
          const outList = e.outbound || [];
          const retList = e.returnTrip || [];
          const hotelParsed = parseHotelDisplayAndUrl(e.lodging);
          const dinners = e.dinners || [];

          cardsHtml += '<div class="travel-card">';
          cardsHtml += '<div class="travel-card-top">';
          cardsHtml += '<div class="travel-name">🧳 ' + escapeHtmlClient(e.name) + '</div>';
          cardsHtml += '<button type="button" class="btn btn-outline btn-sm" onclick="openTravelModal(' + e.rowNumber + ')">✏️ Edit</button>';
          cardsHtml += '</div>';

          cardsHtml += '<div class="travel-grid">';
          cardsHtml += '<div class="travel-leg-box">';
          cardsHtml += '<div class="travel-leg-label"><span>🛫 Outbound to Anaheim</span>' +
            (e.outboundDay ? '<span class="travel-day-tag">' + escapeHtmlClient(e.outboundDay) + '</span>' : '') + '</div>';
          if (outList.length > 0) {
            cardsHtml += '<div style="white-space:pre-wrap; word-break:break-word; color:var(--text);">' +
              outList.map(o => (outList.length > 1 ? '<b>' + escapeHtmlClient(o.day) + ':</b> ' : '') + escapeHtmlClient(o.details)).join('\\n') + '</div>';
          } else {
            cardsHtml += '<div style="color:#9ca3af; font-style:italic;">Not set yet</div>';
          }
          cardsHtml += '</div>';

          cardsHtml += '<div class="travel-leg-box">';
          cardsHtml += '<div class="travel-leg-label"><span>🛬 Return Home</span>' +
            (e.returnDay ? '<span class="travel-day-tag">' + escapeHtmlClient(e.returnDay) + '</span>' : '') + '</div>';
          if (retList.length > 0) {
            cardsHtml += '<div style="white-space:pre-wrap; word-break:break-word; color:var(--text);">' +
              retList.map(r => (retList.length > 1 ? '<b>' + escapeHtmlClient(r.day) + ':</b> ' : '') + escapeHtmlClient(r.details)).join('\\n') + '</div>';
          } else {
            cardsHtml += '<div style="color:#9ca3af; font-style:italic;">Not set yet</div>';
          }
          cardsHtml += '</div>';
          cardsHtml += '</div>';

          cardsHtml += '<div style="display:flex; flex-direction:column; gap:6px; font-size:13px;">';
          if (hotelParsed.name) {
            cardsHtml += '<div>🏨 <strong>Hotel:</strong> ' + escapeHtmlClient(hotelParsed.name) +
              (hotelParsed.url ? ' <a href="' + escapeHtmlClient(hotelParsed.url) + '" target="_blank" rel="noopener" style="color:var(--brand); font-weight:700; text-decoration:none; margin-left:4px;">📍 Map ↗</a>' : '') +
              '</div>';
          }
          if (dinners.length > 0) {
            cardsHtml += '<div style="display:flex; flex-wrap:wrap; gap:5px; align-items:center;"><span>🍽️ <strong>Dinners:</strong></span>' +
              dinners.map(d => '<span class="badge badge-open">' + escapeHtmlClient(DINNER_SHORT_LABELS[d] || d) + '</span>').join('') +
              '</div>';
          }
          if (e.otherPlans) {
            cardsHtml += '<div>🎡 <strong>Other Plans:</strong> ' + escapeHtmlClient(e.otherPlans) + '</div>';
          }
          cardsHtml += '</div>';
          cardsHtml += '</div>';
        });

        cardsEl.innerHTML = cardsHtml;
      }

      function populateWaveExistingDropdown(selectedRowNum) {
        const sel = document.getElementById('waveExistingSelect');
        if (!sel) return;
        const entries = (APP_STATE.raceWaves && APP_STATE.raceWaves.entries) || [];
        let html = '<option value="NEW"' + (!selectedRowNum ? ' selected' : '') + '>➕ Add a New Race Entry</option>';
        entries.forEach(e => {
          const teamStr = [e.name, e.partner1, e.partner2, e.partner3].filter(Boolean).join(' & ');
          const label = teamStr + ' — ' + (e.day || '') + ' (' + (e.division || 'Open') + (e.startWave ? ', ' + e.startWave : '') + ')';
          const isSel = Number(selectedRowNum) === Number(e.rowNumber) ? ' selected' : '';
          html += '<option value="' + e.rowNumber + '"' + isSel + '>' + escapeHtmlClient(label) + '</option>';
        });
        sel.innerHTML = html;
      }

      function setWaveTimeMode(mode) {
        WAVE_MODAL_CTX.timeMode = mode === 'tbd' ? 'tbd' : 'set';
        document.getElementById('waveTimeModeSet').classList.toggle('active', WAVE_MODAL_CTX.timeMode === 'set');
        document.getElementById('waveTimeModeTbd').classList.toggle('active', WAVE_MODAL_CTX.timeMode === 'tbd');
        document.getElementById('waveTimePickerRow').style.display = WAVE_MODAL_CTX.timeMode === 'set' ? 'grid' : 'none';
      }

      function onWaveHourChange(hourVal) {
        const h = parseInt(hourVal, 10);
        const ampmSel = document.getElementById('waveAmPmSelect');
        if (!ampmSel || isNaN(h)) return;
        if (h >= 7 && h <= 11) {
          ampmSel.value = 'AM';
        } else {
          ampmSel.value = 'PM';
        }
      }

      function applyTimeToWavePickers(startWaveStr) {
        const str = String(startWaveStr || '').trim();
        if (!str) {
          setWaveTimeMode('tbd');
          document.getElementById('waveHourSelect').value = '10';
          document.getElementById('waveMinSelect').value = '00';
          document.getElementById('waveAmPmSelect').value = 'AM';
          return;
        }
        setWaveTimeMode('set');
        const m = str.match(/^(\\d{1,2}):(\\d{2})\\s*(AM|PM)?$/i);
        if (m) {
          const h = String(parseInt(m[1], 10));
          const min = m[2];
          const mer = (m[3] || 'AM').toUpperCase();
          document.getElementById('waveHourSelect').value = h;
          document.getElementById('waveMinSelect').value = min;
          document.getElementById('waveAmPmSelect').value = mer;
        }
      }

      function onWaveDivisionChange() {
        const div = String(document.getElementById('waveDivisionSelect').value || '');
        const isRelay = /relay/i.test(div);
        const isDoubles = /doubles/i.test(div);
        const hasP1 = Boolean(document.getElementById('wavePartner1Input').value.trim());
        const hasP2P3 = Boolean(document.getElementById('wavePartner2Input').value.trim() || document.getElementById('wavePartner3Input').value.trim());

        document.getElementById('wavePrimaryLabel').textContent = isRelay ? 'Relay Athlete 1' : (isDoubles ? 'Athlete 1' : 'Athlete Name');
        document.getElementById('wavePartner1Label').textContent = isRelay ? 'Relay Partner 1 (Athlete 2)' : 'Doubles Partner';
        document.getElementById('wavePartner1Field').style.display = (isDoubles || isRelay || hasP1) ? 'block' : 'none';
        document.getElementById('waveRelayPartnersField').style.display = (isRelay || hasP2P3) ? 'grid' : 'none';
      }

      function loadWaveEntryIntoModal(rowNumber, forceSetTimeMode) {
        const entries = (APP_STATE.raceWaves && APP_STATE.raceWaves.entries) || [];
        const entry = entries.find(e => Number(e.rowNumber) === Number(rowNumber));

        if (!entry) {
          WAVE_MODAL_CTX.rowNumber = null;
          WAVE_MODAL_CTX.originalName = '';
          WAVE_MODAL_CTX.originalDay = '';
          document.getElementById('waveModalTitle').textContent = '⏱️ Add New Race Entry';
          document.getElementById('waveDaySelect').value = (RACE_DAY_FILTER !== 'ALL' ? RACE_DAY_FILTER : 'Thursday');
          document.getElementById('waveDivisionSelect').value = 'Womens Open';
          document.getElementById('waveNameInput').value = '';
          document.getElementById('wavePartner1Input').value = '';
          document.getElementById('wavePartner2Input').value = '';
          document.getElementById('wavePartner3Input').value = '';
          setWaveTimeMode('tbd');
          onWaveDivisionChange();
          document.getElementById('waveDeleteWrap').style.display = 'none';
          return;
        }

        WAVE_MODAL_CTX.rowNumber = entry.rowNumber;
        WAVE_MODAL_CTX.originalName = entry.name || '';
        WAVE_MODAL_CTX.originalDay = entry.day || 'Thursday';

        document.getElementById('waveModalTitle').textContent = '⏱️ Edit Start Wave — ' + (entry.name || '');
        document.getElementById('waveDaySelect').value = entry.day || 'Thursday';
        document.getElementById('waveDivisionSelect').value = entry.division || 'Womens Open';
        document.getElementById('waveNameInput').value = entry.name || '';
        document.getElementById('wavePartner1Input').value = entry.partner1 || '';
        document.getElementById('wavePartner2Input').value = entry.partner2 || '';
        document.getElementById('wavePartner3Input').value = entry.partner3 || '';

        applyTimeToWavePickers(entry.startWave);
        if (forceSetTimeMode && !entry.startWave) {
          setWaveTimeMode('set');
        }
        onWaveDivisionChange();
        document.getElementById('waveDeleteWrap').style.display = 'block';
      }

      function onWaveExistingChange(val) {
        if (val === 'NEW' || !val) {
          loadWaveEntryIntoModal(null, false);
        } else {
          loadWaveEntryIntoModal(Number(val), false);
        }
      }

      function openWaveModal(rowNumber, focusTimePicker) {
        populateClubDatalist();
        populateWaveExistingDropdown(rowNumber);
        loadWaveEntryIntoModal(rowNumber, Boolean(focusTimePicker));
        document.getElementById('waveModalBackdrop').classList.add('open');
      }

      function closeWaveModal() {
        document.getElementById('waveModalBackdrop').classList.remove('open');
      }

      function submitWaveModal() {
        const name = document.getElementById('waveNameInput').value.trim();
        if (!name) {
          showToast('⚠️ Please enter the primary athlete name.');
          return;
        }
        const day = document.getElementById('waveDaySelect').value;
        const division = document.getElementById('waveDivisionSelect').value;
        const isRelay = /relay/i.test(division);
        const isDoubles = /doubles/i.test(division);

        let startWave = '';
        if (WAVE_MODAL_CTX.timeMode === 'set') {
          const h = document.getElementById('waveHourSelect').value;
          const m = document.getElementById('waveMinSelect').value;
          const ampm = document.getElementById('waveAmPmSelect').value;
          startWave = h + ':' + m + ' ' + ampm;
        }

        const partner1 = (isDoubles || isRelay || document.getElementById('wavePartner1Field').style.display !== 'none')
          ? document.getElementById('wavePartner1Input').value.trim()
          : '';
        const partner2 = (isRelay || document.getElementById('waveRelayPartnersField').style.display !== 'none')
          ? document.getElementById('wavePartner2Input').value.trim()
          : '';
        const partner3 = (isRelay || document.getElementById('waveRelayPartnersField').style.display !== 'none')
          ? document.getElementById('wavePartner3Input').value.trim()
          : '';

        const btn = document.getElementById('waveSaveBtn');
        btn.disabled = true;
        btn.textContent = 'Saving...';

        const payload = {
          rowNumber: WAVE_MODAL_CTX.rowNumber,
          originalName: WAVE_MODAL_CTX.originalName,
          originalDay: WAVE_MODAL_CTX.originalDay,
          day: day,
          startWave: startWave,
          division: division,
          name: name,
          partner1: partner1,
          partner2: partner2,
          partner3: partner3,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Save Wave';
            closeWaveModal();
            APP_STATE = newState;
            renderAll();
            showToast('✅ Saved race wave for ' + name + (startWave ? ' (' + startWave + ')' : ''));
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Save Wave';
            showToast('❌ Could not save wave: ' + (err.message || err));
          })
          .saveRaceWaveEntry(payload);
      }

      function openDeleteWaveModal() {
        if (!WAVE_MODAL_CTX.rowNumber) return;
        document.getElementById('deleteWaveModalText').textContent =
          'Remove "' + (WAVE_MODAL_CTX.originalName || 'this entry') + '" from the Hyrox Anaheim 2026 times sheet?';
        const btn = document.getElementById('confirmDeleteWaveBtn');
        btn.disabled = false;
        btn.textContent = 'Yes, Delete';
        document.getElementById('deleteWaveModalBackdrop').classList.add('open');
      }

      function closeDeleteWaveModal() {
        document.getElementById('deleteWaveModalBackdrop').classList.remove('open');
      }

      function executeDeleteWave() {
        if (!WAVE_MODAL_CTX.rowNumber) return;
        const btn = document.getElementById('confirmDeleteWaveBtn');
        btn.disabled = true;
        btn.textContent = 'Deleting...';

        const payload = {
          rowNumber: WAVE_MODAL_CTX.rowNumber,
          name: WAVE_MODAL_CTX.originalName,
          day: WAVE_MODAL_CTX.originalDay,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Yes, Delete';
            closeDeleteWaveModal();
            closeWaveModal();
            APP_STATE = newState;
            renderAll();
            showToast('🗑️ Deleted race entry.');
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Yes, Delete';
            closeDeleteWaveModal();
            showToast('❌ Could not delete entry: ' + (err.message || err));
          })
          .deleteRaceWaveEntry(payload);
      }

      function populateTravelModalControls(selectedRowNum) {
        const travelData = APP_STATE.raceTravel || {
          entries: [],
          outboundDays: ['Wednesday 12/2', 'Thursday 12/3', 'Friday 12/4', 'Saturday 12/5', 'Sunday 12/6'],
          returnDays: ['Thursday 12/3', 'Friday 12/4', 'Saturday 12/5', 'Sunday 12/6', 'Monday 12/7'],
          dinnerLabels: ['Thurs Dinner', 'Friday Dinner', 'Saturday Dinner', 'Sunday Dinner', 'Monday Dinner']
        };
        const entries = travelData.entries || [];

        const existSel = document.getElementById('travelExistingSelect');
        let existHtml = '<option value="NEW"' + (!selectedRowNum ? ' selected' : '') + '>➕ Add New Travel Entry</option>';
        entries.forEach(e => {
          const isSel = Number(selectedRowNum) === Number(e.rowNumber) ? ' selected' : '';
          existHtml += '<option value="' + e.rowNumber + '"' + isSel + '>' + escapeHtmlClient(e.name) + '</option>';
        });
        existSel.innerHTML = existHtml;

        const outSel = document.getElementById('travelOutboundDaySelect');
        outSel.innerHTML = '<option value="">-- Select Outbound Day --</option>' +
          (travelData.outboundDays || []).map(d => '<option value="' + escapeHtmlClient(d) + '">' + escapeHtmlClient(d) + '</option>').join('');

        const retSel = document.getElementById('travelReturnDaySelect');
        retSel.innerHTML = '<option value="">-- Select Return Day --</option>' +
          (travelData.returnDays || []).map(d => '<option value="' + escapeHtmlClient(d) + '">' + escapeHtmlClient(d) + '</option>').join('');

        const chipSet = new Set(DEFAULT_HOTEL_CHIPS);
        entries.forEach(e => {
          const parsed = parseHotelDisplayAndUrl(e.lodging);
          if (parsed.name && parsed.name.toLowerCase() !== 'n/a') chipSet.add(parsed.name);
        });
        const chipsEl = document.getElementById('travelHotelChips');
        chipsEl.innerHTML = Array.from(chipSet).map((h, idx) =>
          '<button type="button" class="hotel-chip-btn" onclick="selectHotelChip(' + idx + ')" data-hotel="' + escapeHtmlClient(h) + '">🏨 ' + escapeHtmlClient(h) + '</button>'
        ).join('');
      }

      function selectHotelChip(idx) {
        const chipList = Array.from(new Set(DEFAULT_HOTEL_CHIPS));
        if (chipList[idx]) {
          document.getElementById('travelLodgingInput').value = chipList[idx];
        }
      }

      function renderTravelDinnerButtons() {
        const travelData = APP_STATE.raceTravel || {};
        const dinnerLabels = travelData.dinnerLabels || ['Thurs Dinner', 'Friday Dinner', 'Saturday Dinner', 'Sunday Dinner', 'Monday Dinner'];
        const container = document.getElementById('travelDinnerToggles');
        if (!container) return;

        const activeSet = new Set((TRAVEL_MODAL_CTX.dinners || []).map(d => d.toLowerCase()));
        container.innerHTML = dinnerLabels.map((dl, idx) => {
          const isActive = activeSet.has(dl.toLowerCase());
          const shortLabel = DINNER_SHORT_LABELS[dl] || dl;
          return '<button type="button" class="dinner-toggle-btn ' + (isActive ? 'active' : '') +
            '" onclick="toggleTravelDinner(' + idx + ')">' +
            (isActive ? '✅ ' : '') + escapeHtmlClient(shortLabel) + '</button>';
        }).join('');
      }

      function toggleTravelDinner(idx) {
        const travelData = APP_STATE.raceTravel || {};
        const dinnerLabels = travelData.dinnerLabels || ['Thurs Dinner', 'Friday Dinner', 'Saturday Dinner', 'Sunday Dinner', 'Monday Dinner'];
        const target = dinnerLabels[idx];
        if (!target) return;
        const cur = TRAVEL_MODAL_CTX.dinners || [];
        const foundIdx = cur.findIndex(d => d.toLowerCase() === target.toLowerCase());
        if (foundIdx >= 0) {
          cur.splice(foundIdx, 1);
        } else {
          cur.push(target);
        }
        TRAVEL_MODAL_CTX.dinners = cur;
        renderTravelDinnerButtons();
      }

      function loadTravelEntryIntoModal(rowNumber) {
        const entries = (APP_STATE.raceTravel && APP_STATE.raceTravel.entries) || [];
        const entry = entries.find(e => Number(e.rowNumber) === Number(rowNumber));

        if (!entry) {
          TRAVEL_MODAL_CTX.rowNumber = null;
          TRAVEL_MODAL_CTX.originalName = '';
          TRAVEL_MODAL_CTX.dinners = [];
          document.getElementById('travelModalTitle').textContent = '✈️ Add Travel & Dinners';
          document.getElementById('travelNameInput').value = '';
          document.getElementById('travelOutboundDaySelect').value = 'Wednesday 12/2';
          document.getElementById('travelOutboundDetails').value = '';
          document.getElementById('travelReturnDaySelect').value = 'Sunday 12/6';
          document.getElementById('travelReturnDetails').value = '';
          document.getElementById('travelLodgingInput').value = '';
          document.getElementById('travelOtherPlansInput').value = '';
          document.getElementById('travelDeleteWrap').style.display = 'none';
          renderTravelDinnerButtons();
          return;
        }

        TRAVEL_MODAL_CTX.rowNumber = entry.rowNumber;
        TRAVEL_MODAL_CTX.originalName = entry.name || '';
        TRAVEL_MODAL_CTX.dinners = (entry.dinners || []).slice();

        document.getElementById('travelModalTitle').textContent = '✈️ Edit Travel — ' + (entry.name || '');
        document.getElementById('travelNameInput').value = entry.name || '';
        document.getElementById('travelOutboundDaySelect').value = entry.outboundDay || '';
        document.getElementById('travelOutboundDetails').value = entry.outboundDetails || '';
        document.getElementById('travelReturnDaySelect').value = entry.returnDay || '';
        document.getElementById('travelReturnDetails').value = entry.returnDetails || '';
        document.getElementById('travelLodgingInput').value = entry.lodging || '';
        document.getElementById('travelOtherPlansInput').value = entry.otherPlans || '';
        document.getElementById('travelDeleteWrap').style.display = 'block';
        renderTravelDinnerButtons();
      }

      function onTravelExistingChange(val) {
        if (val === 'NEW' || !val) {
          loadTravelEntryIntoModal(null);
        } else {
          loadTravelEntryIntoModal(Number(val));
        }
      }

      function openTravelModal(rowNumber) {
        populateClubDatalist();
        populateTravelModalControls(rowNumber);
        loadTravelEntryIntoModal(rowNumber);
        document.getElementById('travelModalBackdrop').classList.add('open');
      }

      function closeTravelModal() {
        document.getElementById('travelModalBackdrop').classList.remove('open');
      }

      function submitTravelModal() {
        const name = document.getElementById('travelNameInput').value.trim();
        if (!name) {
          showToast('⚠️ Please enter your name(s).');
          return;
        }

        const outboundDay = document.getElementById('travelOutboundDaySelect').value;
        const outboundDetails = document.getElementById('travelOutboundDetails').value.trim();
        const returnDay = document.getElementById('travelReturnDaySelect').value;
        const returnDetails = document.getElementById('travelReturnDetails').value.trim();
        const lodging = document.getElementById('travelLodgingInput').value.trim();
        const otherPlans = document.getElementById('travelOtherPlansInput').value.trim();

        const btn = document.getElementById('travelSaveBtn');
        btn.disabled = true;
        btn.textContent = 'Saving...';

        const payload = {
          rowNumber: TRAVEL_MODAL_CTX.rowNumber,
          originalName: TRAVEL_MODAL_CTX.originalName,
          name: name,
          outboundDay: outboundDay,
          outboundDetails: outboundDetails,
          returnDay: returnDay,
          returnDetails: returnDetails,
          lodging: lodging,
          dinners: TRAVEL_MODAL_CTX.dinners || [],
          otherPlans: otherPlans,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Save Travel Info';
            closeTravelModal();
            APP_STATE = newState;
            renderAll();
            showToast('✅ Saved travel & dinner info for ' + name + '!');
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Save Travel Info';
            showToast('❌ Could not save travel info: ' + (err.message || err));
          })
          .saveRaceTravelEntry(payload);
      }

      function openDeleteTravelModal() {
        if (!TRAVEL_MODAL_CTX.rowNumber) return;
        document.getElementById('deleteTravelModalText').textContent =
          'Remove "' + (TRAVEL_MODAL_CTX.originalName || 'this entry') + '" from the Anaheim Travel sheet?';
        const btn = document.getElementById('confirmDeleteTravelBtn');
        btn.disabled = false;
        btn.textContent = 'Yes, Delete';
        document.getElementById('deleteTravelModalBackdrop').classList.add('open');
      }

      function closeDeleteTravelModal() {
        document.getElementById('deleteTravelModalBackdrop').classList.remove('open');
      }

      function executeDeleteTravel() {
        if (!TRAVEL_MODAL_CTX.rowNumber) return;
        const btn = document.getElementById('confirmDeleteTravelBtn');
        btn.disabled = true;
        btn.textContent = 'Deleting...';

        const payload = {
          rowNumber: TRAVEL_MODAL_CTX.rowNumber,
          name: TRAVEL_MODAL_CTX.originalName,
          coachEmail: getActiveCoachEmail(),
          useTestSheet: Boolean(APP_STATE.useTestSheet)
        };

        google.script.run
          .withSuccessHandler(newState => {
            btn.disabled = false;
            btn.textContent = 'Yes, Delete';
            closeDeleteTravelModal();
            closeTravelModal();
            APP_STATE = newState;
            renderAll();
            showToast('🗑️ Deleted travel entry.');
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = 'Yes, Delete';
            closeDeleteTravelModal();
            showToast('❌ Could not delete travel entry: ' + (err.message || err));
          })
          .deleteRaceTravelEntry(payload);
      }

      initApp();
    </script>
  </body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('F45 Morgan Hill HYROX - Training Sign-Up & Strava Sync')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles redirect back from Strava after user clicks Authorize
 */
function handleOAuthCallback(code, stateName, scope) {
  if (!scope || !scope.includes('activity:read')) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; text-align: center; margin-top: 50px; padding: 0 16px;">' +
      '<h2 style="color: #d93025;">Authorization Incomplete</h2>' +
      '<p>Please grant permission to view activity data so your runs can be recorded.</p>' +
      '</div>'
    ).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');

  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    code: code,
    grant_type: 'authorization_code'
  };

  const response = UrlFetchApp.fetch(SETTINGS.STRAVA_AUTH_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) {
    return HtmlService.createHtmlOutput(`<h3>Error connecting to Strava:</h3><p>${data.message || response.getContentText()}</p>`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  const athlete = data.athlete;
  const sheetName = stateName ? decodeURIComponent(stateName).trim() : `${athlete.firstname || ''} ${athlete.lastname || ''}`.trim();

  // Save athlete in _Strava_Athletes_Config
  const athleteRecord = {
    athleteId: String(athlete.id),
    sheetName: sheetName,
    stravaName: `${athlete.firstname || ''} ${athlete.lastname || ''}`.trim(),
    refreshToken: data.refresh_token,
    accessToken: data.access_token,
    expiresAt: data.expires_at,
    lastSynced: new Date()
  };

  saveAthleteToConfig(athleteRecord);

  // Immediately run full historical sync for this athlete so past weeks populate right away
  try {
    recalculateAthleteRuns(athleteRecord);
  } catch (err) {
    Logger.log('Initial sync error: ' + err.message);
  }

  return HtmlService.createHtmlOutput(`
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; text-align: center; margin-top: 60px; padding: 0 16px;">
      <h2 style="color: #FC4C02;">🎉 You're Connected, ${escapeHtml(sheetName)}!</h2>
      <p style="font-size: 16px; color: #444;">Your Strava runs will now automatically count on the <b>Run Accountability!</b> sheet.</p>
      <p style="color: #888; font-size: 13px;">You can now close this tab.</p>
    </div>
  `).setTitle('Connected Successfully')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// ==========================================
// 2. CONFIG SHEET & TOKEN MANAGEMENT
// ==========================================

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

/**
 * Robustly finds the target accountability sheet tab by:
 * 1. TAB_GID in Script Properties or default GID 1418914961 (immune to tab renames!)
 * 2. TAB_NAME in Script Properties
 * 3. Fallback: 'Run Accountability!' or searches for a tab with 'Accountability' in the name
 */
function getTargetSheet(ss) {
  if (!ss) ss = getSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  // 1. Check GID (Each tab in Google Sheets has a permanent numeric GID that never changes even if renamed)
  const targetGid = props.getProperty('TAB_GID') || '1418914961';
  const sheets = ss.getSheets();
  const byGid = sheets.find(sh => String(sh.getSheetId()) === String(targetGid));
  if (byGid) return byGid;

  // 2. Check custom tab name in Script Properties
  const customName = props.getProperty('TAB_NAME');
  if (customName) {
    const byName = ss.getSheetByName(customName);
    if (byName) return byName;
  }

  // 3. Fallback: SETTINGS.SHEET_TAB ('Run Accountability!')
  const byDefaultName = ss.getSheetByName(SETTINGS.SHEET_TAB);
  if (byDefaultName) return byDefaultName;

  // 4. Fuzzy search: Find any tab containing "Accountability" or "Run"
  const fuzzy = sheets.find(sh => /accountability/i.test(sh.getName()));
  if (fuzzy) return fuzzy;

  return sheets[0];
}

function getOrCreateConfigSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS.CONFIG_TAB);
  const headers = ['Athlete ID', 'Sheet Name', 'Strava Name', 'Refresh Token', 'Access Token', 'Expires At', 'Last Synced'];

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS.CONFIG_TAB);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function saveAthleteToConfig(athlete) {
  const sheet = getOrCreateConfigSheet();
  const values = sheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(athlete.athleteId) || String(values[i][1]).trim().toLowerCase() === athlete.sheetName.toLowerCase()) {
      targetRow = i + 1;
      break;
    }
  }

  const rowData = [
    String(athlete.athleteId),
    athlete.sheetName,
    athlete.stravaName,
    athlete.refreshToken,
    athlete.accessToken,
    athlete.expiresAt,
    athlete.lastSynced
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

function getAllAthletes() {
  const sheet = getOrCreateConfigSheet();
  const values = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) {
      list.push({
        rowIndex: i + 1,
        athleteId: String(values[i][0]),
        sheetName: String(values[i][1]).trim(),
        stravaName: String(values[i][2]),
        refreshToken: String(values[i][3]),
        accessToken: String(values[i][4]),
        expiresAt: Number(values[i][5])
      });
    }
  }
  return list;
}

/**
 * Utility function to view the list of connected athletes and their status in the Execution Log!
 */
function listConnectedAthletes() {
  const athletes = getAllAthletes();
  Logger.log('==============================================');
  Logger.log(`🏃 CONNECTED STRAVA ATHLETES (${athletes.length} total)`);
  Logger.log('==============================================');
  
  if (athletes.length === 0) {
    Logger.log('No athletes have connected yet.');
  } else {
    athletes.forEach((a, idx) => {
      Logger.log(`${idx + 1}. Sheet Name: "${a.sheetName}" | Strava Name: "${a.stravaName}" | Strava ID: ${a.athleteId}`);
    });
  }
  Logger.log('==============================================');
}

function getValidAccessToken(athlete) {
  const nowUnix = Math.floor(Date.now() / 1000);
  if (athlete.accessToken && athlete.expiresAt && athlete.expiresAt > (nowUnix + 120)) {
    return athlete.accessToken;
  }

  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');

  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: athlete.refreshToken,
    grant_type: 'refresh_token'
  };

  const res = UrlFetchApp.fetch(SETTINGS.STRAVA_AUTH_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    Logger.log(`Token refresh failed for ${athlete.sheetName}: ${data.message}`);
    return null;
  }

  const sheet = getOrCreateConfigSheet();
  sheet.getRange(athlete.rowIndex, 4, 1, 3).setValues([[
    data.refresh_token,
    data.access_token,
    data.expires_at
  ]]);

  return data.access_token;
}

// ==========================================
// 3. ADAPTIVE TABLE DETECTION (Moved table)
// ==========================================

/**
 * Dynamically locates the table header row and name column wherever a human places it.
 * Scans the sheet for date-like headers.
 */
function findTableLocation(sheet) {
  const maxSearchRows = Math.min(sheet.getLastRow(), 20);
  const maxSearchCols = Math.min(sheet.getLastColumn(), 25);
  if (maxSearchRows === 0 || maxSearchCols === 0) return null;

  const data = sheet.getRange(1, 1, maxSearchRows, maxSearchCols).getValues();

  let bestHeaderRow = -1;
  let maxDatesFound = 0;

  for (let r = 0; r < data.length; r++) {
    let dateCount = 0;
    for (let c = 0; c < data[r].length; c++) {
      if (parseDateInterval(data[r][c])) {
        dateCount++;
      }
    }
    if (dateCount > maxDatesFound) {
      maxDatesFound = dateCount;
      bestHeaderRow = r + 1; // 1-indexed
    }
  }

  if (bestHeaderRow === -1 || maxDatesFound === 0) {
    // Fallback: row 1
    bestHeaderRow = 1;
  }

  // Find the Name column (the column immediately preceding the first date column, or col 1)
  const headerRowVals = sheet.getRange(bestHeaderRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  let firstDateCol = 2;
  for (let c = 0; c < headerRowVals.length; c++) {
    if (parseDateInterval(headerRowVals[c])) {
      firstDateCol = c + 1;
      break;
    }
  }

  const nameCol = Math.max(1, firstDateCol - 1);

  // Extract athlete list from nameCol starting from bestHeaderRow + 1
  const numRows = Math.max(1, sheet.getLastRow() - bestHeaderRow);
  const colVals = sheet.getRange(bestHeaderRow + 1, nameCol, numRows, 1).getValues();
  const athletes = [];

  for (let i = 0; i < colVals.length; i++) {
    const val = String(colVals[i][0]).trim();
    if (val && !parseDateInterval(val)) {
      athletes.push({
        rawName: val,
        cleanName: cleanAthleteName(val),
        row: bestHeaderRow + 1 + i
      });
    }
  }

  return {
    headerRow: bestHeaderRow,
    nameCol: nameCol,
    firstDateCol: firstDateCol,
    athletes: athletes
  };
}

/**
 * Strips the ⚡ badge or whitespace to ensure clean matching
 */
function cleanAthleteName(name) {
  return String(name || '')
    .replace(/[⚡️🟧🏃]/g, '')
    .trim();
}

// ==========================================
// 4. FUZZY DATE HEADER PARSER & BUILDER
// ==========================================

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_MAP = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11
};

/**
 * Robust date parser handling various human formats:
 * - "Aug 31 - Sep 6"
 * - "Sep 7-13", "Sep 7 - 13"
 * - "9/7 - 9/13", "9/7-13"
 * - "September 7 to September 13"
 * - Actual Date objects (if sheets auto-converted cell format)
 */
function parseDateInterval(cellVal) {
  if (!cellVal) return null;
  const currentYear = new Date().getFullYear();

  if (cellVal instanceof Date) {
    // If a human entered a raw date representing the start of the week
    const start = new Date(cellVal.getFullYear(), cellVal.getMonth(), cellVal.getDate(), 0, 0, 0);
    const end = new Date(start.getTime() + 6 * 86400000 + 86399000);
    return { start, end };
  }

  const str = String(cellVal).trim();
  if (str.length < 3) return null;

  // Format 1: Text Month Day - Text Month Day (e.g. "Aug 31 - Sep 6" or "August 31 to September 6")
  let m = str.match(/([A-Za-z]+)\.?\s*(\d{1,2})\s*(?:[-–—to]+)\s*([A-Za-z]+)\.?\s*(\d{1,2})/i);
  if (m) {
    const m1 = MONTH_MAP[m[1].toLowerCase()];
    const d1 = parseInt(m[2], 10);
    const m2 = MONTH_MAP[m[3].toLowerCase()];
    const d2 = parseInt(m[4], 10);

    if (m1 !== undefined && m2 !== undefined && !isNaN(d1) && !isNaN(d2)) {
      // Handle year rollover (e.g. Dec 28 - Jan 3)
      let yEnd = currentYear;
      if (m1 === 11 && m2 === 0) yEnd = currentYear + 1;
      const start = new Date(currentYear, m1, d1, 0, 0, 0);
      const end = new Date(yEnd, m2, d2, 23, 59, 59);
      return { start, end };
    }
  }

  // Format 2: Text Month Day - Day (e.g. "Sep 7-13" or "Sep 7 - 13")
  m = str.match(/([A-Za-z]+)\.?\s*(\d{1,2})\s*(?:[-–—to]+)\s*(\d{1,2})/i);
  if (m) {
    const month = MONTH_MAP[m[1].toLowerCase()];
    const d1 = parseInt(m[2], 10);
    const d2 = parseInt(m[3], 10);

    if (month !== undefined && !isNaN(d1) && !isNaN(d2)) {
      const start = new Date(currentYear, month, d1, 0, 0, 0);
      const end = new Date(currentYear, month, d2, 23, 59, 59);
      return { start, end };
    }
  }

  // Format 3: Numeric M/D - M/D or M/D - D (e.g. "9/7 - 9/13" or "9/7-13")
  m = str.match(/(\d{1,2})\/(\d{1,2})\s*(?:[-–—to]+)\s*(?:(\d{1,2})\/)?(\d{1,2})/);
  if (m) {
    const m1 = parseInt(m[1], 10) - 1;
    const d1 = parseInt(m[2], 10);
    const m2 = m[3] ? parseInt(m[3], 10) - 1 : m1;
    const d2 = parseInt(m[4], 10);

    if (m1 >= 0 && m1 <= 11 && d1 >= 1 && d1 <= 31 && !isNaN(d2)) {
      const start = new Date(currentYear, m1, d1, 0, 0, 0);
      const end = new Date(currentYear, m2, d2, 23, 59, 59);
      return { start, end };
    }
  }

  return null;
}

/**
 * Returns Sunday-Saturday or Monday-Sunday range for a given date
 */
function getWeekRangeForDate(d, startOnMonday = true) {
  const date = new Date(d);
  const day = date.getDay(); // 0 is Sunday, 1 is Monday
  let diff = date.getDate() - day + (startOnMonday ? (day === 0 ? -6 : 1) : 0);
  
  const start = new Date(date.setDate(diff));
  start.setHours(0, 0, 0, 0);
  
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  end.setHours(23, 59, 59, 999);

  return { start, end };
}

/**
 * Formats a week interval matching Jade's style:
 * If same month: "Sep 21-27"
 * If spans month: "Aug 31 - Sep 6"
 */
function formatWeekHeader(start, end) {
  const mStart = MONTH_NAMES[start.getMonth()];
  const mEnd = MONTH_NAMES[end.getMonth()];
  const dStart = start.getDate();
  const dEnd = end.getDate();

  if (mStart === mEnd) {
    return `${mStart} ${dStart}-${dEnd}`;
  } else {
    return `${mStart} ${dStart} - ${mEnd} ${dEnd}`;
  }
}

/**
 * Inspects the table, parses all existing week columns, and automatically
 * adds a new column if the current week doesn't exist yet!
 */
function getOrAddWeekColumns(sheet, tableLoc) {
  const headerVals = sheet.getRange(tableLoc.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const weeks = [];

  for (let c = tableLoc.nameCol; c < headerVals.length; c++) {
    const parsed = parseDateInterval(headerVals[c]);
    if (parsed) {
      weeks.push({
        colIndex: c + 1,
        rawHeader: String(headerVals[c]),
        start: parsed.start,
        end: parsed.end
      });
    }
  }

  // Edge Case 2 & 3: Check if the CURRENT week exists in the table
  const now = new Date();
  const currentWeek = getWeekRangeForDate(now);
  const hasCurrentWeek = weeks.some(w => now >= w.start && now <= w.end);

  if (!hasCurrentWeek) {
    // Automatically add the new week column to the table!
    const newColIndex = sheet.getLastColumn() + 1;
    const headerTitle = formatWeekHeader(currentWeek.start, currentWeek.end);

    sheet.getRange(tableLoc.headerRow, newColIndex).setValue(headerTitle);
    sheet.getRange(tableLoc.headerRow, newColIndex).setFontWeight('bold');
    Logger.log(`✨ Automatically added new week column "${headerTitle}" at column ${newColIndex}`);

    weeks.push({
      colIndex: newColIndex,
      rawHeader: headerTitle,
      start: currentWeek.start,
      end: currentWeek.end
    });
  }

  return weeks;
}

// ==========================================
// 5. RUN COUNTING & CALCULATION ENGINE
// ==========================================

/**
 * Checks if an activity counts towards Run Accountability:
 * 1. Any Run (sport_type or type === 'Run') with >= 500m or >= 3 mins
 * 2. Any activity on Saturday or Sunday containing "Hyrox" (case-insensitive) in its title
 */
function isQualifyingActivity(act) {
  if (!act) return false;

  const actType = (act.type || '').toLowerCase();
  const sportType = (act.sport_type || '').toLowerCase();
  const name = (act.name || '').toLowerCase();

  // 1. Standard Run
  const isRun = actType === 'run' || sportType === 'run' || actType === 'virtualrun' || sportType === 'virtualrun';
  if (isRun) {
    const distance = act.distance || 0;
    const movingTime = act.moving_time || 0;
    if (distance >= SETTINGS.MIN_RUN_DISTANCE_METERS || movingTime >= SETTINGS.MIN_RUN_DURATION_SECONDS) {
      return true;
    }
  }

  // 2. Weekend "Hyrox" workout (Workout, Crossfit, WeightTraining, etc.)
  if (name.includes('hyrox')) {
    const actDate = new Date(act.start_date_local || act.start_date);
    const dayOfWeek = actDate.getDay(); // 0 = Sunday, 6 = Saturday
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      // Must be a genuine workout (at least 5 minutes)
      const duration = act.elapsed_time || act.moving_time || 0;
      if (duration >= 300) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Recalculates and updates the count of runs for an athlete across all active week columns
 */
function recalculateAthleteRuns(athlete) {
  const token = getValidAccessToken(athlete);
  if (!token) return;

  const ss = getSpreadsheet();
  const sheet = getTargetSheet(ss);
  if (!sheet) return;

  const tableLoc = findTableLocation(sheet);
  if (!tableLoc) {
    Logger.log('Could not locate table in ' + SETTINGS.SHEET_TAB);
    return;
  }

  const weekCols = getOrAddWeekColumns(sheet, tableLoc);
  if (weekCols.length === 0) return;

  // Find earliest start date among columns
  const earliestDate = weekCols.reduce((min, w) => w.start < min ? w.start : min, weekCols[0].start);
  const afterEpoch = Math.floor(earliestDate.getTime() / 1000) - 86400; // 1 day buffer

  // Fetch activities from Strava
  const url = `${SETTINGS.STRAVA_API_BASE}/athlete/activities?after=${afterEpoch}&per_page=100`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log(`Error fetching activities for ${athlete.sheetName}: ${res.getContentText()}`);
    return;
  }

  const activities = JSON.parse(res.getContentText());

  // Count runs per week
  const countsPerCol = {};
  weekCols.forEach(w => countsPerCol[w.colIndex] = 0);

  for (const act of activities) {
    if (!isQualifyingActivity(act)) continue;

    const actDate = new Date(act.start_date_local || act.start_date);
    for (const w of weekCols) {
      if (actDate >= w.start && actDate <= w.end) {
        countsPerCol[w.colIndex] = (countsPerCol[w.colIndex] || 0) + 1;
        break;
      }
    }
  }

  // Find athlete's row in the table (matching clean names ignoring emoji)
  let athleteRow = -1;
  const targetClean = cleanAthleteName(athlete.sheetName).toLowerCase();

  for (const a of tableLoc.athletes) {
    if (a.cleanName.toLowerCase() === targetClean) {
      athleteRow = a.row;
      // Ensure the cell displays the ⚡ badge!
      if (!a.rawName.includes('⚡')) {
        sheet.getRange(athleteRow, tableLoc.nameCol).setValue(`${a.cleanName} ⚡`);
      }
      break;
    }
  }

  // If athlete is not in the table yet, automatically append a row with the ⚡ badge!
  if (athleteRow === -1) {
    const lastRow = sheet.getLastRow();
    athleteRow = lastRow + 1;
    const clean = cleanAthleteName(athlete.sheetName);
    sheet.getRange(athleteRow, tableLoc.nameCol).setValue(`${clean} ⚡`);
    Logger.log(`➕ Added new member "${clean} ⚡" at row ${athleteRow}`);
  }

  // Update cells applying Math.max rule
  for (const w of weekCols) {
    const cell = sheet.getRange(athleteRow, w.colIndex);
    const existingVal = parseInt(cell.getValue(), 10) || 0;
    const stravaCount = countsPerCol[w.colIndex] || 0;

    const finalVal = Math.max(existingVal, stravaCount);
    if (finalVal > 0 && finalVal !== existingVal) {
      cell.setValue(finalVal);
    }
  }
}

/**
 * Sync function called by the 10-minute automated trigger
 */
function syncAllAthletesNow() {
  const athletes = getAllAthletes();
  Logger.log(`Starting automated sync for ${athletes.length} athletes...`);
  for (const a of athletes) {
    try {
      recalculateAthleteRuns(a);
    } catch (err) {
      Logger.log(`Failed to sync ${a.sheetName}: ${err.message}`);
    }
  }
  Logger.log('Sync completed.');
}

// ==========================================
// 6. AUTOMATED 10-MINUTE TRIGGER SETUP
// ==========================================

function setupAutoSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncAllAthletesNow') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('syncAllAthletesNow')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('✅ Success! 10-minute auto-sync trigger is active.');
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ==========================================
// 7. TRAINING SESSIONS SIGN-UP ENGINE & COACH TOOLS
// ==========================================

function getSpreadsheetForMode(useTestSheet) {
  const props = PropertiesService.getScriptProperties();
  const id = useTestSheet
    ? (props.getProperty('TEST_SPREADSHEET_ID') || SETTINGS.TEST_SPREADSHEET_ID)
    : props.getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

/**
 * Finds the TRAINING SESSIONS tab reliably by:
 * 1. TRAINING_TAB_GID script property or default GID 728305493
 * 2. TRAINING_TAB_NAME script property or 'TRAINING SESSIONS'
 * 3. Fuzzy search for 'training' and 'session'
 */
function getTrainingSheet(ss) {
  if (!ss) ss = getSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const targetGid = props.getProperty('TRAINING_TAB_GID') || SETTINGS.TRAINING_TAB_GID;
  const sheets = ss.getSheets();

  const byGid = sheets.find(sh => String(sh.getSheetId()) === String(targetGid));
  if (byGid) return byGid;

  const customName = props.getProperty('TRAINING_TAB_NAME') || SETTINGS.TRAINING_TAB;
  const byName = ss.getSheetByName(customName);
  if (byName) return byName;

  const fuzzy = sheets.find(sh => /training\s*session/i.test(sh.getName()));
  if (fuzzy) return fuzzy;

  return null;
}

/**
 * Dynamically locates the header row and maps column numbers (1-indexed) in TRAINING SESSIONS.
 * Handles banner/link rows added above the table header row.
 */
function findTrainingColumns(sheet) {
  const maxCols = Math.max(11, sheet.getLastColumn());
  const scanRows = Math.min(10, Math.max(1, sheet.getLastRow()));
  const topRows = sheet.getRange(1, 1, scanRows, maxCols).getValues();

  let headerRow = 1;
  let bestScore = -1;
  for (let r = 0; r < topRows.length; r++) {
    let score = 0;
    for (let c = 0; c < topRows[r].length; c++) {
      const cell = String(topRows[r][c] || '').trim().toUpperCase();
      if (cell.includes('DATE') || cell === 'TIME' || cell.includes('RSVP') ||
          cell === 'EVENT' || cell.includes('MEMBER SIGN') || cell.includes('WAITLIST')) {
        score++;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      headerRow = r + 1;
    }
  }

  const headerVals = topRows[headerRow - 1] || [];
  const cols = {
    headerRow: headerRow,
    dateCol: 1,
    timeCol: 2,
    rsvpCol: 3,
    eventCol: 4,
    locationCol: 5,
    costCol: 6,
    memberCol: 7,
    waitlistCol: 8,
    coordCol: 9,
    notesCol: 10,
    stationsCol: 11
  };

  headerVals.forEach((raw, idx) => {
    const h = String(raw || '').trim().toUpperCase();
    const colNum = idx + 1;
    if (h.includes('DATE')) cols.dateCol = colNum;
    else if (h === 'TIME') cols.timeCol = colNum;
    else if (h.includes('RSVP')) cols.rsvpCol = colNum;
    else if (h === 'EVENT') cols.eventCol = colNum;
    else if (h.includes('LOCATION')) cols.locationCol = colNum;
    else if (h.includes('COST')) cols.costCol = colNum;
    else if (h.includes('MEMBER SIGN')) cols.memberCol = colNum;
    else if (h.includes('NON-MEMBER') || h.includes('WAITLIST')) cols.waitlistCol = colNum;
    else if (h.includes('COORDINATED')) cols.coordCol = colNum;
    else if (h === 'NOTES') cols.notesCol = colNum;
    else if (h.includes('STATION')) cols.stationsCol = colNum;
  });

  return cols;
}

/**
 * Extracts a Date from human strings like:
 * - "SIMULATION DAY 2: Sunday 11/15/26"
 * - "TRAINING ON: Saturday 10/24/26 \n\n*NOTE..."
 * - "TRAINING ON: Sunday August 9, 2026"
 * - "Fri 7/17/2026"
 */
function parseTrainingSessionDate(cellVal) {
  if (!cellVal) return null;
  if (cellVal instanceof Date) {
    return new Date(cellVal.getFullYear(), cellVal.getMonth(), cellVal.getDate(), 23, 59, 59);
  }

  // Only inspect the first line so notes on line 2 (like "*NOTE NO TRAINING WEEKEND OF 10/16-10/18") don't confuse it
  const firstLine = String(cellVal).split(/\r?\n/)[0].trim();

  // Pattern 1: M/D/YY or M/D/YYYY
  let m = firstLine.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    const month = parseInt(m[1], 10) - 1;
    const day = parseInt(m[2], 10);
    let year = parseInt(m[3], 10);
    if (year < 100) year += 2000;
    return new Date(year, month, day, 23, 59, 59);
  }

  // Pattern 2: MonthName Day, YYYY or MonthName Day
  m = firstLine.match(/([A-Za-z]+)\s+(\d{1,2})(?:,\s*(\d{4}))?/);
  if (m) {
    const monthIdx = MONTH_MAP[m[1].toLowerCase()];
    if (monthIdx !== undefined) {
      const day = parseInt(m[2], 10);
      const year = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
      return new Date(year, monthIdx, day, 23, 59, 59);
    }
  }

  return null;
}

/**
 * Parses the MEMBER SIGN UP cell into either:
 * - { mode: 'slotted', banner: string, slots: [{ time: '10:45 AM', capacity: 2, entries: ['Enrique & Bob', 'Steve T'] }] }
 * - { mode: 'group', entries: ['Steve T', 'Josh'] }
 */
function parseMemberSignupCell(rawText) {
  const str = String(rawText || '').replace(/\r\n/g, '\n').trim();
  if (!str) {
    return { mode: 'group', entries: [] };
  }

  const lines = str.split('\n');
  const timeHeaderRegex = /^\s*(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)\s*$/;
  const numberedLineRegex = /^\s*(\d+)[\.\)]\s*(.*)$/;

  let timeHeaderCount = 0;
  let numberedCount = 0;

  for (const line of lines) {
    if (timeHeaderRegex.test(line)) timeHeaderCount++;
    else if (numberedLineRegex.test(line)) numberedCount++;
  }

  // A session is time-slotted if it has time headers and either multiple time headers or numbered slots under a time header
  const isSlotted = timeHeaderCount >= 2 || (timeHeaderCount >= 1 && numberedCount >= 1);

  if (isSlotted) {
    const bannerLines = [];
    const rawSlots = [];
    let currentSlot = null;

    for (const rawLine of lines) {
      const trimmed = rawLine.trim();
      if (!trimmed) continue;

      const timeMatch = trimmed.match(timeHeaderRegex);
      if (timeMatch) {
        currentSlot = {
          time: timeMatch[1].toUpperCase().replace(/\s+/g, ' ').trim(),
          numberedSpotsCount: 0,
          entries: []
        };
        rawSlots.push(currentSlot);
        continue;
      }

      if (!currentSlot) {
        bannerLines.push(trimmed);
      } else {
        const numMatch = trimmed.match(numberedLineRegex);
        if (numMatch) {
          const spotNum = parseInt(numMatch[1], 10);
          currentSlot.numberedSpotsCount = Math.max(currentSlot.numberedSpotsCount, spotNum);
          const entryName = numMatch[2].trim();
          if (entryName) {
            currentSlot.entries.push(entryName);
          }
        } else {
          // Unnumbered non-empty line inside a slot block
          currentSlot.entries.push(trimmed);
        }
      }
    }

    const banner = bannerLines.join(' ');
    let defaultCap = 2;
    if (/three|3/i.test(banner)) defaultCap = 3;
    else if (/four|4/i.test(banner)) defaultCap = 4;
    else if (/one|1/i.test(banner)) defaultCap = 1;

    const maxObservedSpots = rawSlots.reduce((m, s) => Math.max(m, s.numberedSpotsCount, s.entries.length), 0);
    if (maxObservedSpots > 0) defaultCap = maxObservedSpots;

    const slots = rawSlots.map(s => ({
      time: s.time,
      capacity: Math.max(s.numberedSpotsCount || defaultCap, s.entries.length),
      entries: s.entries
    }));

    return {
      mode: 'slotted',
      banner: banner || `ONLY ${defaultCap === 2 ? 'TWO' : defaultCap} GROUPS PER TIME SLOT`,
      slots: slots
    };
  }

  // Group Session Mode
  const entries = [];
  let maxNumberedSpot = 0;
  let emptyNumberedCount = 0;
  let explicitCapacity = 0;

  for (const rawLine of lines) {
    const trimmed = rawLine.trim();
    if (!trimmed) continue;
    // Ignore accidental single bare timestamp if no numbered slots exist
    if (lines.length === 1 && timeHeaderRegex.test(trimmed)) continue;
    // Ignore instruction banner lines (e.g. "If working with a partner, please list both names...")
    if (/if working with a partner/i.test(trimmed)) continue;

    const capBannerMatch = trimmed.match(/^(?:MAX|ONLY|UP TO)\s+(\d+)\s+(?:MEMBER\s+)?(?:SPOTS|SLOTS|ATHLETES|GROUPS)/i);
    if (capBannerMatch) {
      explicitCapacity = Math.max(explicitCapacity, parseInt(capBannerMatch[1], 10));
      continue;
    }

    const numMatch = trimmed.match(numberedLineRegex);
    if (numMatch) {
      const spotNum = parseInt(numMatch[1], 10);
      maxNumberedSpot = Math.max(maxNumberedSpot, spotNum);
      const cleaned = numMatch[2].trim();
      if (cleaned) {
        entries.push(cleaned);
      } else {
        emptyNumberedCount++;
      }
    } else {
      const cleaned = trimmed.replace(/^[-•*]\s*/, '').trim();
      if (cleaned) {
        entries.push(cleaned);
      }
    }
  }

  let capacity = 0;
  if (explicitCapacity > 0) {
    capacity = Math.max(explicitCapacity, entries.length);
  } else if (emptyNumberedCount > 0 && maxNumberedSpot > 0) {
    capacity = Math.max(maxNumberedSpot, entries.length);
  }

  return { mode: 'group', capacity: capacity, entries: entries };
}

/**
 * Serializes a parsed signup structure back into clean cell text for MEMBER SIGN UP
 */
function serializeMemberSignupCell(signup) {
  if (!signup) return '';

  if (signup.mode === 'slotted') {
    const blocks = [];
    if (signup.banner) {
      blocks.push(signup.banner.trim());
    }
    for (const slot of (signup.slots || [])) {
      const cap = Math.max(slot.capacity || 2, (slot.entries || []).length);
      const slotLines = [slot.time];
      for (let i = 0; i < cap; i++) {
        const val = (slot.entries && slot.entries[i]) ? String(slot.entries[i]).trim() : '';
        slotLines.push(val ? `${i + 1}. ${val}` : `${i + 1}.`);
      }
      blocks.push(slotLines.join('\n'));
    }
    return blocks.join('\n\n');
  }

  // Group session
  const entries = (signup.entries || []).map(e => String(e).trim()).filter(Boolean);
  const cap = Number(signup.capacity) || 0;
  if (cap > 0) {
    const total = Math.max(cap, entries.length);
    const lines = [`MAX ${total} SPOTS`];
    for (let i = 0; i < total; i++) {
      const val = entries[i] || '';
      lines.push(val ? `${i + 1}. ${val}` : `${i + 1}.`);
    }
    return lines.join('\n');
  }

  if (entries.length === 0) return '';
  return entries.map((e, i) => `${i + 1}. ${e}`).join('\n');
}

/**
 * Parses the NON-MEMBER WAITLIST cell (supports both newline-separated and comma-separated lists)
 */
function parseWaitlistCell(rawText) {
  const str = String(rawText || '').replace(/\r\n/g, '\n').trim();
  if (!str) return [];

  const hasNewlines = str.includes('\n');
  const parts = hasNewlines ? str.split('\n') : str.split(',');
  const entries = [];

  for (const rawPart of parts) {
    const cleaned = rawPart
      .trim()
      .replace(/^\d+[\.\)]\s*/, '')
      .replace(/^[-•]\s*/, '')
      .trim();
    if (cleaned) entries.push(cleaned);
  }
  return entries;
}

/**
 * Serializes waitlist entries as a clean numbered list
 */
function serializeWaitlistCell(entries) {
  const clean = (entries || []).map(e => String(e).trim()).filter(Boolean);
  if (clean.length === 0) return '';
  return clean.map((e, i) => `${i + 1}. ${e}`).join('\n');
}

/**
 * Builds the deduplicated club roster from Run Accountability! + _Strava_Athletes_Config
 */
function getClubRoster(ss) {
  const nameSet = new Set();
  try {
    const accSheet = getTargetSheet(ss);
    if (accSheet) {
      const loc = findTableLocation(accSheet);
      if (loc && loc.athletes) {
        loc.athletes.forEach(a => {
          if (a.cleanName) nameSet.add(a.cleanName);
        });
      }
    }
  } catch (e) {}

  try {
    const cfgSheet = ss.getSheetByName(SETTINGS.CONFIG_TAB);
    if (cfgSheet) {
      const vals = cfgSheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        const n = cleanAthleteName(vals[i][1]);
        if (n) nameSet.add(n);
      }
    }
  } catch (e) {}

  return Array.from(nameSet).sort((a, b) => a.localeCompare(b));
}

/**
 * Returns the lowercase list of authorized Coach emails:
 * - Default coaches: steve@steviet.org, beevan.jedi@gmail.com
 * - The spreadsheet owner's email (e.g. bjadesiao@gmail.com on the live sheet)
 * - Any comma-separated emails in optional Script Property COACH_EMAILS
 */
function getAuthorizedCoachEmails(ss) {
  const emailSet = new Set(['steve@steviet.org', 'beevan.jedi@gmail.com']);
  try {
    if (ss && typeof ss.getOwner === 'function') {
      const owner = ss.getOwner();
      if (owner && typeof owner.getEmail === 'function') {
        const ownerEmail = String(owner.getEmail() || '').trim().toLowerCase();
        if (ownerEmail) emailSet.add(ownerEmail);
      }
    }
  } catch (e) {}

  try {
    const props = PropertiesService.getScriptProperties();
    const extra = String(props.getProperty('COACH_EMAILS') || '');
    extra.split(',').forEach(em => {
      const clean = em.trim().toLowerCase();
      if (clean) emailSet.add(clean);
    });
  } catch (e) {}

  return Array.from(emailSet);
}

/**
 * Safely returns the active Google session email if available
 */
function getDetectedUserEmail() {
  try {
    if (typeof Session !== 'undefined' && Session && typeof Session.getActiveUser === 'function') {
      const u = Session.getActiveUser();
      if (u && typeof u.getEmail === 'function') {
        return String(u.getEmail() || '').trim().toLowerCase();
      }
    }
  } catch (e) {}
  return '';
}

/**
 * Checks if the active Google user or claimed coach email is an authorized Coach
 */
function isUserCoach(ss, claimedCoachEmail) {
  const allowed = getAuthorizedCoachEmails(ss);
  const detected = getDetectedUserEmail();
  if (detected && allowed.includes(detected)) return true;
  const claimed = String(claimedCoachEmail || '').trim().toLowerCase();
  if (claimed && allowed.includes(claimed)) return true;
  return false;
}

/**
 * Verifies a Coach email from the Coach Sign-In modal and returns updated state
 */
function verifyCoachEmail(email, useTestSheet) {
  const cleanEmail = String(email || '').trim().toLowerCase();
  const state = getTrainingAppState(useTestSheet, cleanEmail);
  return {
    isCoach: Boolean(state.isCoach),
    state: state
  };
}

/**
 * Returns the full state for the Training Sign-Up UI
 */
function getTrainingAppState(useTestSheet, claimedCoachEmail) {
  const ss = getSpreadsheetForMode(useTestSheet);
  const sheet = getTrainingSheet(ss);
  if (!sheet) {
    throw new Error('Could not find the TRAINING SESSIONS tab in the spreadsheet.');
  }

  const cols = findTrainingColumns(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(11, sheet.getLastColumn());
  const sessions = [];

  if (lastRow > cols.headerRow) {
    const numRows = lastRow - cols.headerRow;
    const data = sheet.getRange(cols.headerRow + 1, 1, numRows, lastCol).getValues();
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowNumber = cols.headerRow + 1 + i;
      const dateText = String(row[cols.dateCol - 1] || '').trim();
      const timeText = String(row[cols.timeCol - 1] || '').trim();
      const rsvpText = String(row[cols.rsvpCol - 1] || '').trim();
      const eventText = String(row[cols.eventCol - 1] || '').trim();
      const locationText = String(row[cols.locationCol - 1] || '').trim();
      const costText = String(row[cols.costCol - 1] || '').trim();
      const memberRaw = String(row[cols.memberCol - 1] || '');
      const waitlistRaw = String(row[cols.waitlistCol - 1] || '');
      const coordText = String(row[cols.coordCol - 1] || '').trim();
      const notesText = String(row[cols.notesCol - 1] || '').trim();

      // Skip blank rows or rows that have neither a date nor an event title
      if (!dateText && !eventText) continue;

      const parsedDate = parseTrainingSessionDate(row[cols.dateCol - 1]);
      // Skip non-session banner/link rows placed below the header row
      if (!parsedDate) {
        const combined = `${dateText} ${eventText}`;
        if (/https?:\/\/|script\.google\.com/i.test(combined) || (!timeText && !memberRaw.trim())) {
          continue;
        }
      }
      const isUpcoming = parsedDate ? (parsedDate >= todayStart) : (sessions.length === 0);

      sessions.push({
        rowNumber: rowNumber,
        dateText: dateText,
        timeText: timeText,
        rsvpText: rsvpText,
        eventText: eventText,
        locationText: locationText,
        costText: costText,
        coordText: coordText,
        notesText: notesText,
        isUpcoming: isUpcoming,
        signup: parseMemberSignupCell(memberRaw),
        waitlist: parseWaitlistCell(waitlistRaw)
      });
    }
  }

  let stravaClientId = '';
  try {
    stravaClientId = String(PropertiesService.getScriptProperties().getProperty('STRAVA_CLIENT_ID') || '').trim();
  } catch (e) {}

  return {
    sessions: sessions,
    roster: getClubRoster(ss),
    raceWaves: getRaceWavesState(ss),
    raceTravel: getRaceTravelState(ss),
    isCoach: isUserCoach(ss, claimedCoachEmail),
    detectedEmail: getDetectedUserEmail(),
    useTestSheet: Boolean(useTestSheet),
    spreadsheetUrl: ss.getUrl(),
    stravaClientId: stravaClientId,
    scriptUrl: getWebAppUrl()
  };
}

/**
 * Locates the exact session row inside the lock (verifies rowNumber or falls back to matching dateText)
 */
function resolveSessionRowIndex(sheet, cols, rowNumber, sessionKey) {
  const lastRow = sheet.getLastRow();
  if (rowNumber > cols.headerRow && rowNumber <= lastRow) {
    const checkDate = String(sheet.getRange(rowNumber, cols.dateCol).getValue() || '').trim();
    if (!sessionKey || checkDate === String(sessionKey).trim()) {
      return rowNumber;
    }
  }
  if (sessionKey) {
    const dates = sheet.getRange(cols.headerRow + 1, cols.dateCol, Math.max(1, lastRow - cols.headerRow), 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (String(dates[i][0] || '').trim() === String(sessionKey).trim()) {
        return cols.headerRow + 1 + i;
      }
    }
  }
  throw new Error('Could not locate the training session row in the sheet. Please refresh and try again.');
}

/**
 * Registers an athlete or doubles pair into a session (Member slot, Member group list, or Non-Member Waitlist)
 */
function registerTrainingSignup(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const entryText = String((payload && payload.entryText) || '').replace(/[\r\n]+/g, ' ').trim();
    if (!entryText) throw new Error('Name cannot be empty.');

    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getTrainingSheet(ss);
    const cols = findTrainingColumns(sheet);
    const targetRow = resolveSessionRowIndex(sheet, cols, Number(payload.rowNumber), payload.sessionKey);

    if (payload.listType === 'waitlist') {
      const cell = sheet.getRange(targetRow, cols.waitlistCol);
      const waitlist = parseWaitlistCell(cell.getValue());
      const norm = entryText.toLowerCase();
      if (waitlist.some(e => e.toLowerCase() === norm)) {
        throw new Error(`"${entryText}" is already on the Non-Member / Waitlist for this session.`);
      }
      waitlist.push(entryText);
      cell.setValue(serializeWaitlistCell(waitlist));
    } else {
      const cell = sheet.getRange(targetRow, cols.memberCol);
      const signup = parseMemberSignupCell(cell.getValue());
      const norm = entryText.toLowerCase();

      if (signup.mode === 'slotted') {
        const slotTime = String(payload.slotTime || '').trim().toUpperCase();
        const slot = signup.slots.find(s => s.time.toUpperCase() === slotTime);
        if (!slot) {
          throw new Error(`Time slot "${payload.slotTime}" was not found.`);
        }
        if (slot.entries.length >= slot.capacity) {
          throw new Error(`Sorry, the ${slot.time} time slot is already full (${slot.capacity}/${slot.capacity}). Please pick another slot!`);
        }
        if (slot.entries.some(e => e.toLowerCase() === norm)) {
          throw new Error(`"${entryText}" is already signed up in the ${slot.time} slot.`);
        }
        slot.entries.push(entryText);
      } else {
        const cap = Number(signup.capacity) || 0;
        if (cap > 0 && signup.entries.length >= cap) {
          throw new Error(`Sorry, all ${cap} member spots for this session are already full! Please join the Non-Member / Waitlist.`);
        }
        if (signup.entries.some(e => e.toLowerCase() === norm)) {
          throw new Error(`"${entryText}" is already signed up for this session.`);
        }
        signup.entries.push(entryText);
      }

      cell.setValue(serializeMemberSignupCell(signup));
    }

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet, payload && payload.coachEmail);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deregisters an athlete or pair from a session
 */
function deregisterTrainingSignup(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const entryText = String((payload && payload.entryText) || '').trim();
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getTrainingSheet(ss);
    const cols = findTrainingColumns(sheet);
    const targetRow = resolveSessionRowIndex(sheet, cols, Number(payload.rowNumber), payload.sessionKey);

    if (payload.listType === 'waitlist') {
      const cell = sheet.getRange(targetRow, cols.waitlistCol);
      const waitlist = parseWaitlistCell(cell.getValue());
      const idx = waitlist.findIndex(e => e.trim().toLowerCase() === entryText.toLowerCase());
      if (idx !== -1) {
        waitlist.splice(idx, 1);
      } else if (typeof payload.entryIndex === 'number' && waitlist[payload.entryIndex] !== undefined) {
        waitlist.splice(payload.entryIndex, 1);
      } else {
        throw new Error(`Could not find "${entryText}" on the waitlist.`);
      }
      cell.setValue(serializeWaitlistCell(waitlist));
    } else {
      const cell = sheet.getRange(targetRow, cols.memberCol);
      const signup = parseMemberSignupCell(cell.getValue());

      if (signup.mode === 'slotted') {
        const slotTime = String(payload.slotTime || '').trim().toUpperCase();
        const slot = signup.slots.find(s => s.time.toUpperCase() === slotTime);
        if (!slot) throw new Error(`Time slot "${payload.slotTime}" not found.`);

        const idx = slot.entries.findIndex(e => e.trim().toLowerCase() === entryText.toLowerCase());
        if (idx !== -1) {
          slot.entries.splice(idx, 1);
        } else if (typeof payload.entryIndex === 'number' && slot.entries[payload.entryIndex] !== undefined) {
          slot.entries.splice(payload.entryIndex, 1);
        } else {
          throw new Error(`Could not find "${entryText}" in slot ${slot.time}.`);
        }
      } else {
        const idx = signup.entries.findIndex(e => e.trim().toLowerCase() === entryText.toLowerCase());
        if (idx !== -1) {
          signup.entries.splice(idx, 1);
        } else if (typeof payload.entryIndex === 'number' && signup.entries[payload.entryIndex] !== undefined) {
          signup.entries.splice(payload.entryIndex, 1);
        } else {
          throw new Error(`Could not find "${entryText}" in member sign-ups.`);
        }
      }

      cell.setValue(serializeMemberSignupCell(signup));
    }

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet, payload && payload.coachEmail);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds structured time slots starting from a start time string (e.g. "10:30 AM"), interval in minutes, count, and optional existing entries per slot
 */
function buildTimeSlotStructure(firstSlotStr, intervalMins, numSlots, slotCap, existingSlotEntries) {
  const m = String(firstSlotStr || '10:30 AM').trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (!m) {
    throw new Error(`Invalid First Slot time "${firstSlotStr}". Use format like "10:30 AM".`);
  }
  let hours = parseInt(m[1], 10);
  let mins = parseInt(m[2], 10);
  const mer = (m[3] || 'AM').toUpperCase();

  if (mer === 'PM' && hours < 12) hours += 12;
  if (mer === 'AM' && hours === 12) hours = 0;

  let totalMins = hours * 60 + mins;
  const cap = Math.max(1, Number(slotCap) || 2);
  const count = Math.max(1, Math.min(40, Number(numSlots) || 12));
  const step = Math.max(1, Number(intervalMins) || 8);

  const slots = [];
  for (let i = 0; i < count; i++) {
    const h24 = Math.floor(totalMins / 60) % 24;
    const minPart = totalMins % 60;
    const ampm = h24 >= 12 ? 'PM' : 'AM';
    const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
    const timeLabel = `${h12}:${minPart < 10 ? '0' + minPart : minPart} ${ampm}`;
    const preserved = (existingSlotEntries && existingSlotEntries[i]) ? existingSlotEntries[i] : [];
    slots.push({
      time: timeLabel,
      capacity: Math.max(cap, preserved.length),
      entries: preserved
    });
    totalMins += step;
  }

  const wordMap = { 1: 'ONE', 2: 'TWO', 3: 'THREE', 4: 'FOUR' };
  const capWord = wordMap[cap] || String(cap);
  const banner = `ONLY ${capWord} ${cap === 1 ? 'GROUP' : 'GROUPS'} PER TIME SLOT`;

  return {
    mode: 'slotted',
    banner: banner,
    slots: slots
  };
}

/**
 * Generates time slot strings starting from a start time string (e.g. "10:30 AM"), interval in minutes, and count
 */
function generateTimeSlotTemplate(firstSlotStr, intervalMins, numSlots, slotCap) {
  return serializeMemberSignupCell(buildTimeSlotStructure(firstSlotStr, intervalMins, numSlots, slotCap, []));
}

/**
 * Coach Helper: Inserts a new training session row right under headers
 */
function coachCreateTrainingSession(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getTrainingSheet(ss);
    const cols = findTrainingColumns(sheet);

    const insertRow = cols.headerRow + 1;
    sheet.insertRowBefore(insertRow);

    const maxCol = Math.max(11, cols.stationsCol);
    const rowValues = new Array(maxCol).fill('');

    rowValues[cols.dateCol - 1] = String(payload.dateText || '').trim();
    rowValues[cols.timeCol - 1] = String(payload.timeText || '').trim();
    rowValues[cols.rsvpCol - 1] = String(payload.rsvpText || '').trim();
    rowValues[cols.eventCol - 1] = String(payload.eventText || 'HYROX training!').trim();
    rowValues[cols.locationCol - 1] = String(payload.locationText || 'F45').trim();
    rowValues[cols.costCol - 1] = String(payload.costText || SETTINGS.DEFAULT_COST_TEXT).trim();
    rowValues[cols.coordCol - 1] = String(payload.coordText || 'Jade').trim();
    rowValues[cols.notesCol - 1] = String(payload.notesText || '').trim();

    if (payload.mode === 'slotted') {
      rowValues[cols.memberCol - 1] = generateTimeSlotTemplate(
        payload.firstSlot || '10:30 AM',
        payload.intervalMins || 8,
        payload.numSlots || 12,
        payload.slotCap || 2
      );
    } else {
      const groupCap = Math.max(0, Number(payload.groupCap) || 0);
      if (groupCap > 0) {
        rowValues[cols.memberCol - 1] = serializeMemberSignupCell({
          mode: 'group',
          capacity: groupCap,
          entries: []
        });
      }
    }

    const targetRange = sheet.getRange(insertRow, 1, 1, maxCol);
    targetRange.setValues([rowValues]);
    targetRange.setFontWeight('normal');
    targetRange.setVerticalAlignment('top');
    targetRange.setWrap(true);

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Coach Helper: Updates an existing session's details, slots & schedule while preserving existing sign-ups in slot order
 */
function coachUpdateTrainingSession(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getTrainingSheet(ss);
    const cols = findTrainingColumns(sheet);
    const targetRow = resolveSessionRowIndex(sheet, cols, Number(payload.rowNumber), payload.sessionKey);

    // Update session metadata columns if provided in payload
    if (payload.dateText !== undefined) {
      const cleanDate = String(payload.dateText || '').trim();
      if (!cleanDate) throw new Error('Date Header cannot be empty.');
      sheet.getRange(targetRow, cols.dateCol).setValue(cleanDate);
    }
    if (payload.timeText !== undefined) {
      sheet.getRange(targetRow, cols.timeCol).setValue(String(payload.timeText || '').trim());
    }
    if (payload.rsvpText !== undefined) {
      sheet.getRange(targetRow, cols.rsvpCol).setValue(String(payload.rsvpText || '').trim());
    }
    if (payload.eventText !== undefined) {
      sheet.getRange(targetRow, cols.eventCol).setValue(String(payload.eventText || '').trim());
    }
    if (payload.locationText !== undefined) {
      sheet.getRange(targetRow, cols.locationCol).setValue(String(payload.locationText || '').trim());
    }
    if (payload.costText !== undefined) {
      sheet.getRange(targetRow, cols.costCol).setValue(String(payload.costText || '').trim());
    }
    if (payload.coordText !== undefined) {
      sheet.getRange(targetRow, cols.coordCol).setValue(String(payload.coordText || '').trim());
    }
    if (payload.notesText !== undefined) {
      sheet.getRange(targetRow, cols.notesCol).setValue(String(payload.notesText || '').trim());
    }

    const cell = sheet.getRange(targetRow, cols.memberCol);
    const currentSignup = parseMemberSignupCell(cell.getValue());

    if (payload.mode === 'slotted') {
      const numSlots = Math.max(1, Math.min(40, Number(payload.numSlots) || 12));
      const slotCap = Math.max(1, Number(payload.slotCap) || 2);
      const intervalMins = Math.max(1, Number(payload.intervalMins) || 8);
      const firstSlot = String(payload.firstSlot || '10:30 AM').trim();

      let existingSlotEntries = [];
      if (currentSignup.mode === 'slotted') {
        const curSlots = currentSignup.slots || [];
        let lastOccupiedSlot = 0;
        let maxSlotEntries = 0;
        curSlots.forEach((sl, idx) => {
          const arr = sl.entries || [];
          if (arr.length > 0) {
            lastOccupiedSlot = idx + 1;
            maxSlotEntries = Math.max(maxSlotEntries, arr.length);
          }
          existingSlotEntries.push(arr);
        });
        if (numSlots < lastOccupiedSlot) {
          throw new Error(`Cannot reduce to ${numSlots} slots because Slot #${lastOccupiedSlot} already has athletes signed up.`);
        }
        if (slotCap < maxSlotEntries) {
          throw new Error(`Cannot set per-slot capacity to ${slotCap} because a slot already has ${maxSlotEntries} signed up.`);
        }
      } else {
        const groupEntries = currentSignup.entries || [];
        if (groupEntries.length > numSlots * slotCap) {
          throw new Error(`Cannot fit ${groupEntries.length} existing sign-ups into ${numSlots} slots of ${slotCap}.`);
        }
        for (let i = 0; i < numSlots; i++) {
          existingSlotEntries.push(groupEntries.slice(i * slotCap, (i + 1) * slotCap));
        }
      }

      const updatedSignup = buildTimeSlotStructure(firstSlot, intervalMins, numSlots, slotCap, existingSlotEntries);
      cell.setValue(serializeMemberSignupCell(updatedSignup));
    } else {
      const groupCap = Math.max(0, Number(payload.groupCap) || 0);
      let existingEntries = [];
      if (currentSignup.mode === 'slotted') {
        (currentSignup.slots || []).forEach(sl => {
          (sl.entries || []).forEach(e => existingEntries.push(e));
        });
      } else {
        existingEntries = currentSignup.entries || [];
      }
      if (groupCap > 0 && groupCap < existingEntries.length) {
        throw new Error(`Cannot set total slots to ${groupCap} because ${existingEntries.length} athletes are already signed up.`);
      }
      cell.setValue(serializeMemberSignupCell({
        mode: 'group',
        capacity: groupCap,
        entries: existingEntries
      }));
    }

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Coach Helper: Permanently deletes / cancels a training session row from TRAINING SESSIONS
 */
function coachDeleteTrainingSession(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getTrainingSheet(ss);
    const cols = findTrainingColumns(sheet);
    const targetRow = resolveSessionRowIndex(sheet, cols, Number(payload.rowNumber), payload.sessionKey);

    sheet.deleteRow(targetRow);

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 8. HYROX ANAHEIM 2026 RACE HUB (START WAVES & TRAVEL)
// ==========================================

const RACE_DAYS_ORDER = ['Thursday', 'Friday', 'Saturday', 'Sunday'];
const DEFAULT_OUTBOUND_DAYS = ['Wednesday 12/2', 'Thursday 12/3', 'Friday 12/4', 'Saturday 12/5', 'Sunday 12/6'];
const DEFAULT_RETURN_DAYS = ['Thursday 12/3', 'Friday 12/4', 'Saturday 12/5', 'Sunday 12/6', 'Monday 12/7'];
const DEFAULT_DINNER_LABELS = ['Thurs Dinner', 'Friday Dinner', 'Saturday Dinner', 'Sunday Dinner', 'Monday Dinner'];

/**
 * Resiliently finds a sheet tab by exact trimmed name or regex pattern
 */
function getSheetByPattern(ss, exactName, regexPattern) {
  if (!ss || typeof ss.getSheets !== 'function') return null;
  const sheets = ss.getSheets() || [];
  const targetTrim = String(exactName || '').trim().toLowerCase();
  if (targetTrim) {
    const byExact = sheets.find(sh => String((sh && sh.getName && sh.getName()) || '').trim().toLowerCase() === targetTrim);
    if (byExact) return byExact;
  }
  if (regexPattern) {
    const byRegex = sheets.find(sh => regexPattern.test(String((sh && sh.getName && sh.getName()) || '')));
    if (byRegex) return byRegex;
  }
  return null;
}

function getRaceWavesSheet(ss) {
  return getSheetByPattern(ss, SETTINGS.RACE_WAVES_TAB, /Hyrox\s+Anaheim\s+2026\s+times/i);
}

function getRaceTravelSheet(ss) {
  return getSheetByPattern(ss, SETTINGS.RACE_TRAVEL_TAB, /Travel.*Anaheim.*2026/i);
}

function normalizeRaceDay(rawDay) {
  const s = String(rawDay || '').trim();
  if (/^thu/i.test(s)) return 'Thursday';
  if (/^fri/i.test(s)) return 'Friday';
  if (/^sat/i.test(s)) return 'Saturday';
  if (/^sun/i.test(s)) return 'Sunday';
  if (/^wed/i.test(s)) return 'Wednesday';
  return s;
}

/**
 * Parses and formats a wave time cell value (Date object or string like "9:50AM", "1:30 PM", "13:30")
 */
function parseAndFormatWaveTime(rawVal) {
  if (!rawVal && rawVal !== 0) {
    return { startWave: '', startMins: 9999 };
  }
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    const h24 = rawVal.getHours();
    const min = rawVal.getMinutes();
    const ampm = h24 >= 12 ? 'PM' : 'AM';
    const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
    const fmt = `${h12}:${min < 10 ? '0' + min : min} ${ampm}`;
    return { startWave: fmt, startMins: h24 * 60 + min };
  }
  const str = String(rawVal || '').trim();
  if (!str) return { startWave: '', startMins: 9999 };

  const m = str.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (m) {
    let h = parseInt(m[1], 10);
    const min = parseInt(m[2], 10);
    let mer = m[3] ? m[3].toUpperCase() : '';
    if (!mer) {
      if (h >= 13 && h <= 23) {
        mer = 'PM';
        h = h - 12;
      } else if (h === 12) {
        mer = 'PM';
      } else if (h >= 1 && h <= 6) {
        mer = 'PM';
      } else {
        mer = 'AM';
      }
    }
    let h24 = h;
    if (mer === 'PM' && h24 < 12) h24 += 12;
    if (mer === 'AM' && h24 === 12) h24 = 0;
    const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
    const fmt = `${h12}:${min < 10 ? '0' + min : min} ${mer}`;
    return { startWave: fmt, startMins: h24 * 60 + min };
  }

  return { startWave: str, startMins: 9999 };
}

/**
 * Dynamically locates the header row and columns (1-indexed) in 'Hyrox Anaheim 2026 times'
 */
function findRaceWavesColumns(sheet) {
  const maxCols = Math.max(7, sheet.getLastColumn());
  const scanRows = Math.min(15, Math.max(1, sheet.getLastRow()));
  const topRows = sheet.getRange(1, 1, scanRows, maxCols).getValues();

  let headerRow = 1;
  let bestScore = -1;
  for (let r = 0; r < topRows.length; r++) {
    let score = 0;
    for (let c = 0; c < topRows[r].length; c++) {
      const cell = String(topRows[r][c] || '').replace(/\s+/g, ' ').trim().toUpperCase();
      if (cell === 'DAY' || cell.includes('START WAVE') || cell === 'DIVISION' ||
          cell === 'NAME' || cell.includes('PARTNER')) {
        score++;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      headerRow = r + 1;
    }
  }

  const headerVals = topRows[headerRow - 1] || [];
  const cols = {
    headerRow: headerRow,
    dayCol: 1,
    waveCol: 2,
    divisionCol: 3,
    nameCol: 4,
    partner1Col: 5,
    partner2Col: 6,
    partner3Col: 7
  };

  headerVals.forEach((raw, idx) => {
    const h = String(raw || '').replace(/\s+/g, ' ').trim().toUpperCase();
    const colNum = idx + 1;
    if (h === 'DAY') cols.dayCol = colNum;
    else if (h.includes('START') || h.includes('WAVE')) cols.waveCol = colNum;
    else if (h.includes('DIVISION')) cols.divisionCol = colNum;
    else if (h === 'NAME' || h === 'ATHLETE') cols.nameCol = colNum;
    else if (h.includes('PARTNER 2')) cols.partner2Col = colNum;
    else if (h.includes('PARTNER 3')) cols.partner3Col = colNum;
    else if (h.includes('DOUBLES') || h.includes('PARTNER')) cols.partner1Col = colNum;
  });

  return cols;
}

/**
 * Reads all race wave entries from 'Hyrox Anaheim 2026 times'
 */
function getRaceWavesState(ss) {
  const sheet = getRaceWavesSheet(ss);
  if (!sheet) {
    return { entries: [], days: RACE_DAYS_ORDER };
  }

  const cols = findRaceWavesColumns(sheet);
  const lastRow = sheet.getLastRow();
  const maxCol = Math.max(7, cols.partner3Col, sheet.getLastColumn());
  const entries = [];

  if (lastRow > cols.headerRow) {
    const numRows = lastRow - cols.headerRow;
    const data = sheet.getRange(cols.headerRow + 1, 1, numRows, maxCol).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowNumber = cols.headerRow + 1 + i;
      const day = normalizeRaceDay(row[cols.dayCol - 1]);
      const waveParsed = parseAndFormatWaveTime(row[cols.waveCol - 1]);
      const division = String(row[cols.divisionCol - 1] || '').trim();
      const name = String(row[cols.nameCol - 1] || '').trim();
      const partner1 = String(row[cols.partner1Col - 1] || '').trim();
      const partner2 = String(row[cols.partner2Col - 1] || '').trim();
      const partner3 = String(row[cols.partner3Col - 1] || '').trim();

      if (!name && !division && !waveParsed.startWave && !partner1) continue;
      if (/https?:\/\/|script\.google\.com/i.test(`${day} ${name}`)) continue;

      entries.push({
        rowNumber: rowNumber,
        day: day || 'Thursday',
        startWave: waveParsed.startWave,
        startMins: waveParsed.startMins,
        division: division,
        name: name,
        partner1: partner1,
        partner2: partner2,
        partner3: partner3
      });
    }
  }

  return {
    entries: entries,
    days: RACE_DAYS_ORDER
  };
}

/**
 * Resolves the target row number in 'Hyrox Anaheim 2026 times' for updating or deleting
 */
function resolveRaceWaveRowIndex(sheet, cols, rowNumber, matchName, matchDay) {
  const lastRow = sheet.getLastRow();
  const normName = String(matchName || '').trim().toLowerCase();
  if (rowNumber > cols.headerRow && rowNumber <= lastRow) {
    const curName = String(sheet.getRange(rowNumber, cols.nameCol).getValue() || '').trim().toLowerCase();
    if (!normName || curName === normName) {
      return rowNumber;
    }
  }
  if (normName && lastRow > cols.headerRow) {
    const numRows = lastRow - cols.headerRow;
    const maxCol = Math.max(cols.nameCol, cols.dayCol);
    const data = sheet.getRange(cols.headerRow + 1, 1, numRows, maxCol).getValues();
    const normDay = normalizeRaceDay(matchDay).toLowerCase();
    for (let i = 0; i < data.length; i++) {
      const rName = String(data[i][cols.nameCol - 1] || '').trim().toLowerCase();
      const rDay = normalizeRaceDay(data[i][cols.dayCol - 1]).toLowerCase();
      if (rName === normName && (!normDay || rDay === normDay)) {
        return cols.headerRow + 1 + i;
      }
    }
    for (let i = 0; i < data.length; i++) {
      const rName = String(data[i][cols.nameCol - 1] || '').trim().toLowerCase();
      if (rName === normName) {
        return cols.headerRow + 1 + i;
      }
    }
  }
  throw new Error('Could not find that race wave entry in the sheet. Please refresh and try again.');
}

/**
 * Creates or updates an athlete/team row in 'Hyrox Anaheim 2026 times'
 */
function saveRaceWaveEntry(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getRaceWavesSheet(ss);
    if (!sheet) {
      throw new Error('Could not find the "Hyrox Anaheim 2026 times" tab in the spreadsheet.');
    }

    const name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Primary athlete name cannot be empty.');

    const day = normalizeRaceDay((payload && payload.day) || 'Thursday');
    const waveParsed = parseAndFormatWaveTime(payload && payload.startWave);
    const division = String((payload && payload.division) || 'Womens Open').trim();
    const partner1 = String((payload && payload.partner1) || '').trim();
    const partner2 = String((payload && payload.partner2) || '').trim();
    const partner3 = String((payload && payload.partner3) || '').trim();

    const cols = findRaceWavesColumns(sheet);
    const maxCol = Math.max(7, cols.partner3Col);
    const rowValues = new Array(maxCol).fill('');
    rowValues[cols.dayCol - 1] = day;
    rowValues[cols.waveCol - 1] = waveParsed.startWave;
    rowValues[cols.divisionCol - 1] = division;
    rowValues[cols.nameCol - 1] = name;
    rowValues[cols.partner1Col - 1] = partner1;
    rowValues[cols.partner2Col - 1] = partner2;
    rowValues[cols.partner3Col - 1] = partner3;

    const reqRow = Number(payload && payload.rowNumber) || 0;
    if (reqRow > 0) {
      const targetRow = resolveRaceWaveRowIndex(
        sheet,
        cols,
        reqRow,
        payload.originalName || name,
        payload.originalDay || day
      );
      sheet.getRange(targetRow, 1, 1, maxCol).setValues([rowValues]);
    } else {
      const lastRow = sheet.getLastRow();
      const targetDayIdx = RACE_DAYS_ORDER.indexOf(day);
      let insertAfterRow = cols.headerRow;
      let lastDataRow = cols.headerRow;

      if (lastRow > cols.headerRow) {
        const numRows = lastRow - cols.headerRow;
        const data = sheet.getRange(cols.headerRow + 1, 1, numRows, Math.max(cols.dayCol, cols.nameCol)).getValues();
        for (let i = 0; i < data.length; i++) {
          const rNum = cols.headerRow + 1 + i;
          const rDay = normalizeRaceDay(data[i][cols.dayCol - 1]);
          const rName = String(data[i][cols.nameCol - 1] || '').trim();
          if (rDay || rName) {
            lastDataRow = rNum;
            const rDayIdx = RACE_DAYS_ORDER.indexOf(rDay);
            if (targetDayIdx === -1 || (rDayIdx !== -1 && rDayIdx <= targetDayIdx)) {
              insertAfterRow = rNum;
            }
          }
        }
      }

      const targetRow = insertAfterRow + 1;
      if (insertAfterRow < lastDataRow) {
        sheet.insertRowBefore(targetRow);
      }
      sheet.getRange(targetRow, 1, 1, maxCol).setValues([rowValues]);
    }

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet, payload && payload.coachEmail);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes an athlete/team row from 'Hyrox Anaheim 2026 times'
 */
function deleteRaceWaveEntry(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getRaceWavesSheet(ss);
    if (!sheet) {
      throw new Error('Could not find the "Hyrox Anaheim 2026 times" tab in the spreadsheet.');
    }

    const cols = findRaceWavesColumns(sheet);
    const targetRow = resolveRaceWaveRowIndex(
      sheet,
      cols,
      Number(payload && payload.rowNumber) || 0,
      payload && payload.name,
      payload && payload.day
    );
    sheet.deleteRow(targetRow);

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet, payload && payload.coachEmail);
  } finally {
    lock.releaseLock();
  }
}

const FULL_WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function formatTravelHeaderCell(rawVal) {
  if (!rawVal && rawVal !== 0) return '';
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    const dayName = FULL_WEEKDAYS[rawVal.getDay()];
    const m = rawVal.getMonth() + 1;
    const d = rawVal.getDate();
    return `${dayName} ${m}/${d}`;
  }
  return String(rawVal || '').replace(/\s+/g, ' ').trim();
}

/**
 * Dynamically locates the main header row, sub-header row, and column mappings in 'Travel ✈️ Hyrox Anaheim 2026!! '
 */
function findRaceTravelColumns(sheet) {
  const maxCols = Math.max(21, sheet.getLastColumn());
  const scanRows = Math.min(12, Math.max(2, sheet.getLastRow()));
  const topRows = sheet.getRange(1, 1, scanRows, maxCols).getValues();

  let subHeaderRow = 2;
  let bestSubScore = -1;
  for (let r = 0; r < topRows.length; r++) {
    let score = 0;
    for (let c = 0; c < topRows[r].length; c++) {
      const cell = formatTravelHeaderCell(topRows[r][c]);
      if (/\d{1,2}\/\d{1,2}/.test(cell) || /dinner/i.test(cell)) {
        score++;
      }
    }
    if (score > bestSubScore) {
      bestSubScore = score;
      subHeaderRow = r + 1;
    }
  }

  const mainHeaderRow = Math.max(1, subHeaderRow - 1);
  const mainVals = topRows[mainHeaderRow - 1] || [];
  const subVals = topRows[subHeaderRow - 1] || [];

  let nameCol = 1;
  let outboundStartCol = 2;
  let returnStartCol = 8;
  let lodgingCol = 14;
  let dinnerStartCol = 16;
  let otherPlansCol = 21;

  for (let c = 0; c < mainVals.length; c++) {
    const h = String(mainVals[c] || '').replace(/\s+/g, ' ').trim().toUpperCase();
    const colNum = c + 1;
    if (h === 'NAME') nameCol = colNum;
    else if (h.includes('DEPARTURE FROM HOME') || h.includes('OUTBOUND')) outboundStartCol = colNum;
    else if (h.includes('DEPARTURE FROM LA') || h.includes('RETURN')) returnStartCol = colNum;
    else if (h.includes('LODGING') || h.includes('HOTEL')) lodgingCol = colNum;
    else if (h.includes('DINNER')) dinnerStartCol = colNum;
    else if (h.includes('OTHER PLANS')) otherPlansCol = colNum;
  }

  const outboundCols = [];
  const returnCols = [];
  const dinnerCols = [];

  for (let c = 0; c < subVals.length; c++) {
    const label = formatTravelHeaderCell(subVals[c]);
    if (!label) continue;
    const colNum = c + 1;
    if (/dinner/i.test(label)) {
      dinnerCols.push({ label: label, col: colNum });
    } else if (colNum >= outboundStartCol && colNum < returnStartCol) {
      outboundCols.push({ label: label, col: colNum });
    } else if (colNum >= returnStartCol && colNum < lodgingCol) {
      returnCols.push({ label: label, col: colNum });
    }
  }

  if (outboundCols.length === 0) {
    DEFAULT_OUTBOUND_DAYS.forEach((d, idx) => outboundCols.push({ label: d, col: 2 + idx }));
  }
  if (returnCols.length === 0) {
    DEFAULT_RETURN_DAYS.forEach((d, idx) => returnCols.push({ label: d, col: 8 + idx }));
  }
  if (dinnerCols.length === 0) {
    DEFAULT_DINNER_LABELS.forEach((d, idx) => dinnerCols.push({ label: d, col: 16 + idx }));
  }

  return {
    mainHeaderRow: mainHeaderRow,
    subHeaderRow: subHeaderRow,
    nameCol: nameCol,
    outboundCols: outboundCols,
    returnCols: returnCols,
    lodgingCol: lodgingCol,
    dinnerCols: dinnerCols,
    otherPlansCol: otherPlansCol
  };
}

function isTruthyDinnerCell(val) {
  const s = String(val || '').trim().toLowerCase();
  if (!s) return false;
  if (s === 'no' || s === 'n' || s === 'false' || s === '0' || s === '-') return false;
  return true;
}

/**
 * Reads all athlete travel, hotel, and dinner rows from 'Travel ✈️ Hyrox Anaheim 2026!! '
 */
function getRaceTravelState(ss) {
  const sheet = getRaceTravelSheet(ss);
  if (!sheet) {
    return {
      entries: [],
      outboundDays: DEFAULT_OUTBOUND_DAYS,
      returnDays: DEFAULT_RETURN_DAYS,
      dinnerLabels: DEFAULT_DINNER_LABELS
    };
  }

  const cols = findRaceTravelColumns(sheet);
  const lastRow = sheet.getLastRow();
  const maxCol = Math.max(21, cols.otherPlansCol, sheet.getLastColumn());
  const entries = [];

  if (lastRow > cols.subHeaderRow) {
    const numRows = lastRow - cols.subHeaderRow;
    const data = sheet.getRange(cols.subHeaderRow + 1, 1, numRows, maxCol).getValues();
    const validReturnDaySet = new Set(cols.returnCols.map(rc => rc.label.toLowerCase()));

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowNumber = cols.subHeaderRow + 1 + i;
      const name = String(row[cols.nameCol - 1] || '').trim();
      if (!name) continue;
      if (/https?:\/\/|script\.google\.com/i.test(name)) continue;

      let rawOutbound = [];
      for (const oc of cols.outboundCols) {
        const text = String(row[oc.col - 1] || '').trim();
        if (text) rawOutbound.push({ day: oc.label, details: text });
      }

      let rawReturn = [];
      for (const rc of cols.returnCols) {
        const text = String(row[rc.col - 1] || '').trim();
        if (text) rawReturn.push({ day: rc.label, details: text });
      }

      // If an athlete duplicated their return trip text in both the outbound Sunday column and return Sunday column, deduplicate
      if (rawOutbound.length > 1 && rawReturn.length > 0) {
        const retTexts = new Set(rawReturn.map(r => r.details.replace(/\s+/g, ' ').trim().toLowerCase()));
        rawOutbound = rawOutbound.filter(o => !retTexts.has(o.details.replace(/\s+/g, ' ').trim().toLowerCase()));
      }

      // If an athlete entered both their outbound flight and return flight inside the first date block (leaving returnCols blank), split cleanly
      if (rawOutbound.length >= 2 && rawReturn.length === 0) {
        const tail = rawOutbound.slice(1);
        if (tail.every(t => validReturnDaySet.has(t.day.toLowerCase()))) {
          rawReturn = tail;
          rawOutbound = [rawOutbound[0]];
        }
      }

      const lodging = String(row[cols.lodgingCol - 1] || '').trim();
      const dinners = [];
      for (const dc of cols.dinnerCols) {
        if (isTruthyDinnerCell(row[dc.col - 1])) {
          dinners.push(dc.label);
        }
      }
      const otherPlans = String(row[cols.otherPlansCol - 1] || '').trim();

      entries.push({
        rowNumber: rowNumber,
        name: name,
        outbound: rawOutbound,
        outboundDay: rawOutbound[0] ? rawOutbound[0].day : '',
        outboundDetails: rawOutbound.map(o => o.details).join('\n\n'),
        returnTrip: rawReturn,
        returnDay: rawReturn[0] ? rawReturn[0].day : '',
        returnDetails: rawReturn.map(r => r.details).join('\n\n'),
        lodging: lodging,
        dinners: dinners,
        otherPlans: otherPlans
      });
    }
  }

  return {
    entries: entries,
    outboundDays: cols.outboundCols.map(c => c.label),
    returnDays: cols.returnCols.map(c => c.label),
    dinnerLabels: cols.dinnerCols.map(c => c.label)
  };
}

/**
 * Resolves the target row number in 'Travel ✈️ Hyrox Anaheim 2026!! '
 */
function resolveRaceTravelRowIndex(sheet, cols, rowNumber, matchName) {
  const lastRow = sheet.getLastRow();
  const normName = String(matchName || '').trim().toLowerCase();
  if (rowNumber > cols.subHeaderRow && rowNumber <= lastRow) {
    const curName = String(sheet.getRange(rowNumber, cols.nameCol).getValue() || '').trim().toLowerCase();
    if (!normName || curName === normName) {
      return rowNumber;
    }
  }
  if (normName && lastRow > cols.subHeaderRow) {
    const numRows = lastRow - cols.subHeaderRow;
    const names = sheet.getRange(cols.subHeaderRow + 1, cols.nameCol, numRows, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0] || '').trim().toLowerCase() === normName) {
        return cols.subHeaderRow + 1 + i;
      }
    }
  }
  return -1;
}

/**
 * Creates or updates an athlete's travel, hotel & dinner row in 'Travel ✈️ Hyrox Anaheim 2026!! '
 */
function saveRaceTravelEntry(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getRaceTravelSheet(ss);
    if (!sheet) {
      throw new Error('Could not find the "Travel ✈️ Hyrox Anaheim 2026!!" tab in the spreadsheet.');
    }

    const name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Name cannot be empty.');

    const cols = findRaceTravelColumns(sheet);
    const reqRow = Number(payload && payload.rowNumber) || 0;
    let targetRow = resolveRaceTravelRowIndex(sheet, cols, reqRow, payload.originalName || name);

    if (targetRow === -1) {
      // Check if an entry with this name already exists
      targetRow = resolveRaceTravelRowIndex(sheet, cols, 0, name);
    }

    if (targetRow === -1) {
      // Find the last row with a name, or start at subHeaderRow + 2 (leaving the single spacer row 3 intact if empty)
      const lastRow = sheet.getLastRow();
      let lastNamedRow = cols.subHeaderRow + 1;
      if (lastRow > cols.subHeaderRow) {
        const names = sheet.getRange(cols.subHeaderRow + 1, cols.nameCol, lastRow - cols.subHeaderRow, 1).getValues();
        for (let i = 0; i < names.length; i++) {
          if (String(names[i][0] || '').trim()) {
            lastNamedRow = cols.subHeaderRow + 1 + i;
          }
        }
      }
      targetRow = lastNamedRow + 1;
    }

    const maxCol = Math.max(21, cols.otherPlansCol, sheet.getLastColumn());
    const existingRow = targetRow <= sheet.getLastRow()
      ? sheet.getRange(targetRow, 1, 1, maxCol).getValues()[0]
      : new Array(maxCol).fill('');
    const rowValues = existingRow.slice(0, maxCol);
    while (rowValues.length < maxCol) rowValues.push('');

    rowValues[cols.nameCol - 1] = name;

    // Outbound columns
    const outboundDay = String((payload && payload.outboundDay) || '').trim();
    const outboundDetails = String((payload && payload.outboundDetails) || '').trim();
    cols.outboundCols.forEach(oc => { rowValues[oc.col - 1] = ''; });
    if (outboundDetails) {
      const targetOutCol = cols.outboundCols.find(oc => oc.label.toLowerCase() === outboundDay.toLowerCase()) || cols.outboundCols[0];
      if (targetOutCol) rowValues[targetOutCol.col - 1] = outboundDetails;
    }

    // Return columns
    const returnDay = String((payload && payload.returnDay) || '').trim();
    const returnDetails = String((payload && payload.returnDetails) || '').trim();
    cols.returnCols.forEach(rc => { rowValues[rc.col - 1] = ''; });
    if (returnDetails) {
      const targetRetCol = cols.returnCols.find(rc => rc.label.toLowerCase() === returnDay.toLowerCase()) || cols.returnCols[0];
      if (targetRetCol) rowValues[targetRetCol.col - 1] = returnDetails;
    }

    // Lodging
    rowValues[cols.lodgingCol - 1] = String((payload && payload.lodging) || '').trim();

    // Dinners
    const selectedDinners = new Set(
      Array.isArray(payload && payload.dinners)
        ? payload.dinners.map(d => String(d || '').trim().toLowerCase())
        : []
    );
    cols.dinnerCols.forEach(dc => {
      rowValues[dc.col - 1] = selectedDinners.has(dc.label.toLowerCase()) ? 'Y' : '';
    });

    // Other Plans
    rowValues[cols.otherPlansCol - 1] = String((payload && payload.otherPlans) || '').trim();

    const targetRange = sheet.getRange(targetRow, 1, 1, maxCol);
    targetRange.setValues([rowValues]);
    targetRange.setVerticalAlignment('top');
    targetRange.setWrap(true);

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet, payload && payload.coachEmail);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes an athlete's travel row from 'Travel ✈️ Hyrox Anaheim 2026!! '
 */
function deleteRaceTravelEntry(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const useTestSheet = Boolean(payload && payload.useTestSheet);
    const ss = getSpreadsheetForMode(useTestSheet);
    const sheet = getRaceTravelSheet(ss);
    if (!sheet) {
      throw new Error('Could not find the "Travel ✈️ Hyrox Anaheim 2026!!" tab in the spreadsheet.');
    }

    const cols = findRaceTravelColumns(sheet);
    const targetRow = resolveRaceTravelRowIndex(
      sheet,
      cols,
      Number(payload && payload.rowNumber) || 0,
      payload && payload.name
    );
    if (targetRow === -1) {
      throw new Error('Could not find that travel entry in the sheet. Please refresh and try again.');
    }
    sheet.deleteRow(targetRow);

    SpreadsheetApp.flush();
    return getTrainingAppState(useTestSheet, payload && payload.coachEmail);
  } finally {
    lock.releaseLock();
  }
}

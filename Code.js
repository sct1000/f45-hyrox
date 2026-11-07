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

/**
 * Handles Web App GET requests:
 * 1. Strava Webhook subscription verification (?hub.mode=subscribe&...)
 * 2. Strava OAuth callback (?code=...&state=...)
 * 3. Unified Web App with tabs for Training Sign-Up (?page=training) & Strava Sync (?page=strava)
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

  const initialTab = (params.page || '').toLowerCase() === 'strava' ? 'strava' : (
    (params.page || '').toLowerCase() === 'training' ? 'training' : 'training'
  );
  const useTestSheet = params.test === '1' || params.test === 'true';

  return renderUnifiedWebApp(initialTab, useTestSheet);
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
          <button id="tabBtnTraining" class="nav-tab" onclick="switchTab('training')">🏋️ Training Sign-Up</button>
          <button id="tabBtnStrava" class="nav-tab" onclick="switchTab('strava')">🏃 Strava Run Sync</button>
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
        renderAll();
      }

      function switchTab(tab) {
        ACTIVE_TAB = tab;
        document.getElementById('tabTraining').style.display = tab === 'training' ? 'block' : 'none';
        document.getElementById('tabStrava').style.display = tab === 'strava' ? 'block' : 'none';
        document.getElementById('tabBtnTraining').classList.toggle('active', tab === 'training');
        document.getElementById('tabBtnStrava').classList.toggle('active', tab === 'strava');
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

  return {
    sessions: sessions,
    roster: getClubRoster(ss),
    isCoach: isUserCoach(ss, claimedCoachEmail),
    detectedEmail: getDetectedUserEmail(),
    useTestSheet: Boolean(useTestSheet),
    spreadsheetUrl: ss.getUrl()
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

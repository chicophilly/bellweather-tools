/**
 * Sherpa Auth — Shared MSAL authentication helper for CORE tools
 * 
 * Deploy to: chicophilly.github.io/bellweather-tools/lib/sherpa-auth.js
 * 
 * Loaded by any CORE tool that needs to call Microsoft Graph API.
 * Provides a single sign-in experience shared across tools via sessionStorage.
 *
 * USAGE:
 *   <script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script>
 *   <script src="https://chicophilly.github.io/bellweather-tools/lib/sherpa-auth.js"></script>
 *   
 *   await SherpaAuth.init();               // call once on page load
 *   if (!SherpaAuth.isSignedIn()) {
 *     await SherpaAuth.signIn();           // redirects for sign-in
 *   }
 *   const token = await SherpaAuth.getToken();
 *   // ... use token for Graph API calls
 *
 * API:
 *   SherpaAuth.init()           — Handle redirect promise, restore session. Await before using.
 *   SherpaAuth.isSignedIn()     — boolean. Is there a valid account?
 *   SherpaAuth.signIn()         — Trigger loginRedirect. Page will reload after.
 *   SherpaAuth.signOut()        — Trigger logoutRedirect.
 *   SherpaAuth.getToken()       — Returns access token string. Acquires silently or via popup.
 *   SherpaAuth.graph(url, opts) — Convenience wrapper: fetch(url) with auth header. Returns JSON.
 *   SherpaAuth.getAccount()     — Returns the current account object (name, email, etc).
 */

(function(global) {
  'use strict';

  // ─────────────────────────────────────────────────────────────────────────
  // CONFIG — update here if Azure App Registration changes
  // ─────────────────────────────────────────────────────────────────────────
  const CONFIG = {
    clientId: '84b37f0e-b149-4089-9bdc-fc654715db23',
    tenantId: 'ff0e3f46-615a-4a22-8a80-4bb83d627785',
    scopes:   ['Files.ReadWrite.All', 'Sites.ReadWrite.All'],
    graphBase: 'https://graph.microsoft.com/v1.0',
  };

  // ─────────────────────────────────────────────────────────────────────────
  // STATE
  // ─────────────────────────────────────────────────────────────────────────
  let msalApp = null;
  let initialized = false;

  // ─────────────────────────────────────────────────────────────────────────
  // INIT
  // Must be called before any other method. Handles redirect return, loads account.
  // ─────────────────────────────────────────────────────────────────────────
  async function init() {
    if (initialized) return;
    if (typeof msal === 'undefined') {
      throw new Error('SherpaAuth: msal-browser library not loaded. Include ' +
        '<script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script> ' +
        'before sherpa-auth.js');
    }

    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: {
        // sessionStorage so auth persists within a tab but not across browser restarts.
        // Shared across tools in the same origin (chicophilly.github.io).
        cacheLocation: 'sessionStorage',
      },
    });

    try {
      await msalApp.handleRedirectPromise();
    } catch (e) {
      console.error('SherpaAuth: redirect handling failed', e);
    }
    initialized = true;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SIGN-IN STATE
  // ─────────────────────────────────────────────────────────────────────────
  function ensureInit() {
    if (!initialized) {
      throw new Error('SherpaAuth: call await SherpaAuth.init() before using other methods');
    }
  }

  function isSignedIn() {
    ensureInit();
    return msalApp.getAllAccounts().length > 0;
  }

  function getAccount() {
    ensureInit();
    return msalApp.getAllAccounts()[0] || null;
  }

  function signIn() {
    ensureInit();
    return msalApp.loginRedirect({ scopes: CONFIG.scopes });
  }

  function signOut() {
    ensureInit();
    return msalApp.logoutRedirect({ postLogoutRedirectUri: window.location.href });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // TOKEN ACQUISITION
  // ─────────────────────────────────────────────────────────────────────────
  async function getToken() {
    ensureInit();
    const account = msalApp.getAllAccounts()[0];
    if (!account) throw new Error('SherpaAuth: not signed in');
    try {
      const result = await msalApp.acquireTokenSilent({ scopes: CONFIG.scopes, account });
      return result.accessToken;
    } catch (silentErr) {
      // Silent acquisition failed (e.g. expired refresh token). Fall back to popup.
      const result = await msalApp.acquireTokenPopup({ scopes: CONFIG.scopes });
      return result.accessToken;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // GRAPH API WRAPPER
  // Convenience: authenticated fetch against Graph, returns JSON by default.
  // Pass {raw: true} to get the raw Response (for blob downloads).
  // ─────────────────────────────────────────────────────────────────────────
  async function graph(url, opts = {}) {
    const token = await getToken();
    const headers = {
      Authorization: `Bearer ${token}`,
      ...(opts.raw ? {} : { 'Content-Type': 'application/json' }),
      ...(opts.headers || {}),
    };
    const fetchOpts = { ...opts, headers };
    delete fetchOpts.raw;
    const response = await fetch(url, fetchOpts);
    if (!response.ok) {
      let errBody = {};
      try { errBody = await response.json(); } catch {}
      throw new Error(errBody.error?.message || `Graph API error: HTTP ${response.status}`);
    }
    return opts.raw ? response : response.json();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // EXPOSE
  // ─────────────────────────────────────────────────────────────────────────
  global.SherpaAuth = {
    init,
    isSignedIn,
    getAccount,
    signIn,
    signOut,
    getToken,
    graph,
    // Exposed so callers can build Graph URLs without hardcoding the base
    GRAPH_BASE: CONFIG.graphBase,
    SCOPES: CONFIG.scopes,
  };

})(window);

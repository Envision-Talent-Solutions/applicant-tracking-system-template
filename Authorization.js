/**
 * Authorization.js
 * Handles granular OAuth consent for Apps Script web apps.
 *
 * BACKGROUND: Starting January 2026, Google is rolling out granular OAuth consent
 * for published Apps Script web apps. Users can selectively grant/deny individual
 * scopes rather than all-or-nothing approval.
 *
 * This module provides:
 * - Per-scope authorization checking using getAuthorizationInfo(authMode, scopes)
 * - Detection of which specific scopes are granted vs denied
 * - Graceful feature degradation when optional scopes are denied
 * - Automatic prompting for critical scopes using requireScopes()
 * - User-friendly re-authorization flows
 *
 * Reference: https://developers.google.com/apps-script/concepts/scopes
 *
 * OAuth Scopes Used by this Application:
 * - spreadsheets.currentonly : Core functionality (critical)
 * - forms.currentonly        : Form management (optional)
 * - script.scriptapp         : Installable triggers (optional)
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const OAUTH_SCOPES = {
  SPREADSHEETS: 'https://www.googleapis.com/auth/spreadsheets.currentonly',
  FORMS: 'https://www.googleapis.com/auth/forms.currentonly',
  SCRIPT: 'https://www.googleapis.com/auth/script.scriptapp'
};

const FEATURE_CONFIG = {
  SPREADSHEET_CORE: {
    name: 'Spreadsheet Access',
    description: 'Read and write candidate/requisition data',
    scopes: [OAUTH_SCOPES.SPREADSHEETS],
    critical: true
  },
  FORMS: {
    name: 'Google Forms',
    description: 'Create and manage candidate submission forms',
    scopes: [OAUTH_SCOPES.FORMS],
    critical: false
  },
  TRIGGERS: {
    name: 'Script Triggers',
    description: 'Automated background processing',
    scopes: [OAUTH_SCOPES.SCRIPT],
    critical: false
  }
};

// Critical scopes that must be granted for the app to function
const CRITICAL_SCOPES = Object.values(FEATURE_CONFIG)
  .filter(f => f.critical)
  .flatMap(f => f.scopes);

// Error patterns that indicate authorization/permission issues
const AUTH_ERROR_PATTERNS = [
  /authorization/i,
  /permission/i,
  /access.+denied/i,
  /not authorized/i,
  /requires authorization/i,
  /you do not have permission/i
];

// =============================================================================
// CORE AUTHORIZATION FUNCTIONS
// =============================================================================

/**
 * Checks if an error message indicates an authorization problem.
 * @param {string} errorMessage - The error message to check
 * @returns {boolean} True if this appears to be an auth error
 */
function isAuthorizationError(errorMessage) {
  if (!errorMessage) return false;
  return AUTH_ERROR_PATTERNS.some(pattern => pattern.test(errorMessage));
}

/**
 * Gets authorization info for specific scopes.
 * This is the primary method for checking granular consent.
 *
 * @param {string[]} scopes - Array of scope URLs to check
 * @param {ScriptApp.AuthMode} authMode - Auth mode to check (default: FULL)
 * @returns {Object} Authorization info with status and granted scopes
 */
function getAuthInfoForScopes(scopes, authMode) {
  authMode = authMode || ScriptApp.AuthMode.FULL;

  try {
    const authInfo = ScriptApp.getAuthorizationInfo(authMode, scopes);

    return {
      status: authInfo.getAuthorizationStatus(),
      isAuthorized: authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.NOT_REQUIRED,
      grantedScopes: authInfo.getAuthorizedScopes() || [],
      authUrl: authInfo.getAuthorizationUrl() || null,
      requestedScopes: scopes,
      error: null
    };
  } catch (e) {
    // API call failed - treat as unauthorized
    return {
      status: ScriptApp.AuthorizationStatus.REQUIRED,
      isAuthorized: false,
      grantedScopes: [],
      authUrl: null,
      requestedScopes: scopes,
      error: e.message
    };
  }
}

/**
 * Gets a detailed report of all scope authorizations.
 *
 * @returns {Object} Complete authorization report
 */
function getAuthorizationReport() {
  const report = {
    timestamp: new Date().toISOString(),
    features: {},
    scopes: {},
    summary: {
      totalFeatures: 0,
      authorizedFeatures: 0,
      deniedFeatures: 0,
      criticalDenied: []
    }
  };

  // Check each scope individually
  for (const [scopeName, scopeUrl] of Object.entries(OAUTH_SCOPES)) {
    const authInfo = getAuthInfoForScopes([scopeUrl]);
    report.scopes[scopeName] = {
      url: scopeUrl,
      granted: authInfo.grantedScopes.includes(scopeUrl),
      status: String(authInfo.status)
    };
  }

  // Check each feature
  for (const [featureKey, feature] of Object.entries(FEATURE_CONFIG)) {
    const featureScopes = feature.scopes;
    const authInfo = getAuthInfoForScopes(featureScopes);

    // A feature is authorized if ALL its required scopes are granted
    const allScopesGranted = featureScopes.every(scope =>
      authInfo.grantedScopes.includes(scope)
    );

    report.features[featureKey] = {
      name: feature.name,
      description: feature.description,
      critical: feature.critical,
      authorized: allScopesGranted,
      requiredScopes: featureScopes,
      grantedScopes: authInfo.grantedScopes,
      missingScopes: featureScopes.filter(s => !authInfo.grantedScopes.includes(s)),
      authUrl: authInfo.authUrl
    };

    report.summary.totalFeatures++;
    if (allScopesGranted) {
      report.summary.authorizedFeatures++;
    } else {
      report.summary.deniedFeatures++;
      if (feature.critical) {
        report.summary.criticalDenied.push(featureKey);
      }
    }
  }

  return report;
}

/**
 * Checks if a specific feature is authorized (all its scopes granted).
 *
 * @param {string} featureKey - Key from FEATURE_CONFIG
 * @returns {Object} Feature authorization status
 */
function checkFeatureAuthorization(featureKey) {
  const feature = FEATURE_CONFIG[featureKey];
  if (!feature) {
    return {
      authorized: false,
      error: 'Unknown feature: ' + featureKey
    };
  }

  const authInfo = getAuthInfoForScopes(feature.scopes);
  const allScopesGranted = feature.scopes.every(scope =>
    authInfo.grantedScopes.includes(scope)
  );

  return {
    feature: featureKey,
    name: feature.name,
    description: feature.description,
    critical: feature.critical,
    authorized: allScopesGranted,
    requiredScopes: feature.scopes,
    grantedScopes: authInfo.grantedScopes,
    missingScopes: feature.scopes.filter(s => !authInfo.grantedScopes.includes(s)),
    authUrl: authInfo.authUrl
  };
}

// =============================================================================
// SCOPE ENFORCEMENT
// =============================================================================

/**
 * Requires specific scopes before proceeding.
 * Will prompt user and halt execution if scopes not granted.
 * Use this for critical operations that cannot proceed without specific scopes.
 *
 * @param {string[]} scopes - Array of scope URLs to require
 */
function requireScopes(scopes) {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, scopes);
}

/**
 * Requires all critical scopes before proceeding.
 * Call this at the start of critical operations.
 */
function requireCriticalScopes() {
  if (CRITICAL_SCOPES.length > 0) {
    ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, CRITICAL_SCOPES);
  }
}

/**
 * Requires scopes for a specific feature before proceeding.
 *
 * @param {string} featureKey - Key from FEATURE_CONFIG
 */
function requireFeatureScopes(featureKey) {
  const feature = FEATURE_CONFIG[featureKey];
  if (feature && feature.scopes.length > 0) {
    ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, feature.scopes);
  }
}

// =============================================================================
// FEATURE GUARDS
// =============================================================================

/**
 * Creates a guard function for a feature.
 * Returns false and shows a message if the feature's scopes aren't granted.
 *
 * @param {string} featureKey - Key from FEATURE_CONFIG
 * @returns {Function} Guard function returning true if feature is authorized
 */
function createFeatureGuard(featureKey) {
  return function() {
    const status = checkFeatureAuthorization(featureKey);

    if (!status.authorized) {
      const feature = FEATURE_CONFIG[featureKey];
      const missingCount = status.missingScopes.length;

      let message;
      if (feature.critical) {
        message = `${feature.name} is required but not authorized. ` +
                  `Please grant access via System Admin > Authorization.`;
      } else {
        message = `${feature.name} requires additional permissions (${missingCount} scope${missingCount > 1 ? 's' : ''} missing). ` +
                  `This feature is disabled.`;
      }

      try {
        SpreadsheetApp.getActive().toast(message, 'Authorization Required', 8);
      } catch (e) {
        // Spreadsheet access may also be denied
        try {
          SpreadsheetApp.getUi().alert('Authorization Required', message, SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e2) {
          // UI access also denied - critical failure
        }
      }

      return false;
    }

    return true;
  };
}

// Pre-built guards for all features
const AuthGuards = {
  spreadsheet: createFeatureGuard('SPREADSHEET_CORE'),
  forms: createFeatureGuard('FORMS'),
  triggers: createFeatureGuard('TRIGGERS')
};

// =============================================================================
// GRACEFUL DEGRADATION WRAPPERS
// =============================================================================

/**
 * Wraps a function with feature authorization checking.
 * Returns fallback value if feature isn't authorized instead of failing.
 *
 * @param {string} featureKey - Feature key for authorization check
 * @param {Function} fn - Function to execute if authorized
 * @param {Object} options - Options for handling unauthorized state
 * @returns {Object} Result with success status, data, and authorization info
 */
function withFeatureAuth(featureKey, fn, options = {}) {
  const {
    fallbackValue = null,
    showToast = true,
    requireAuth = false  // If true, use requireScopes instead of soft check
  } = options;

  const feature = FEATURE_CONFIG[featureKey];
  if (!feature) {
    return {
      success: false,
      data: fallbackValue,
      error: 'Unknown feature: ' + featureKey,
      authorized: false
    };
  }

  // If requireAuth is true, use hard requirement (will prompt user)
  if (requireAuth) {
    requireFeatureScopes(featureKey);
  }

  // Check authorization
  const authStatus = checkFeatureAuthorization(featureKey);

  if (!authStatus.authorized) {
    if (showToast) {
      const message = `${feature.name} is not authorized. Missing ${authStatus.missingScopes.length} permission(s).`;
      try {
        SpreadsheetApp.getActive().toast(message, 'Feature Disabled', 6);
      } catch (e) {
        // Ignore toast errors
      }
    }

    return {
      success: false,
      data: fallbackValue,
      error: `${feature.name} not authorized`,
      authorized: false,
      missingScopes: authStatus.missingScopes,
      authUrl: authStatus.authUrl
    };
  }

  // Feature is authorized, execute the function
  try {
    const result = fn();
    return {
      success: true,
      data: result,
      error: null,
      authorized: true
    };
  } catch (e) {
    // Check if this is actually an authorization error at runtime
    // (can happen if scope was revoked between check and execution)
    if (isAuthorizationError(e.message)) {
      if (showToast) {
        try {
          SpreadsheetApp.getActive().toast(
            `${feature.name} permission was revoked. Please re-authorize.`,
            'Authorization Revoked',
            8
          );
        } catch (toastErr) {
          // Ignore
        }
      }

      return {
        success: false,
        data: fallbackValue,
        error: e.message,
        authorized: false,  // Auth was revoked
        runtimeAuthError: true,
        missingScopes: feature.scopes  // Assume all scopes for this feature need re-auth
      };
    }

    // Non-auth runtime error
    return {
      success: false,
      data: fallbackValue,
      error: e.message,
      authorized: true,  // Was authorized, but failed for other reason
      runtimeError: true
    };
  }
}

// =============================================================================
// USER PROMPTS AND UI
// =============================================================================

/**
 * Checks critical authorization on startup and prompts if needed.
 * Call this from onOpen().
 *
 * Note: onOpen runs in AuthMode.LIMITED which restricts some operations.
 * This function handles both LIMITED (onOpen) and FULL modes gracefully.
 *
 * @param {ScriptApp.AuthMode} authMode - Optional auth mode (default: tries to detect)
 * @returns {boolean} True if all critical features are authorized
 */
function checkCriticalAuthorization(authMode) {
  // In LIMITED mode (onOpen), we can only do basic checks
  // In FULL mode (user-triggered), we can prompt for authorization
  const isLimitedMode = authMode === ScriptApp.AuthMode.LIMITED;

  try {
    const report = getAuthorizationReport();

    if (report.summary.criticalDenied.length === 0) {
      return true;  // All critical features authorized
    }

    // Critical features are missing
    // In LIMITED mode, we can't prompt - just return false
    if (isLimitedMode) {
      // Try to show a toast as a hint (may fail in LIMITED mode)
      try {
        SpreadsheetApp.getActive().toast(
          'Some features require authorization. Go to System Admin > Authorization.',
          'Authorization Needed',
          10
        );
      } catch (e) {
        // Can't show toast in LIMITED mode, that's fine
      }
      return false;
    }

    // FULL mode - we can prompt user
    const ui = SpreadsheetApp.getUi();
    let message = 'This application requires the following permissions to function:\n\n';

    report.summary.criticalDenied.forEach(featureKey => {
      const feature = report.features[featureKey];
      message += `- ${feature.name}: ${feature.description}\n`;
      if (feature.missingScopes.length > 0) {
        message += `  (Missing: ${feature.missingScopes.map(s => s.split('/').pop()).join(', ')})\n`;
      }
    });

    message += '\nWould you like to authorize now?';

    const response = ui.alert('Authorization Required', message, ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
      // Use requireScopes to force the authorization prompt
      try {
        requireCriticalScopes();

        // IMPORTANT: Re-check after requireScopes() because with granular consent
        // the user may have only granted some of the requested scopes
        const postAuthReport = getAuthorizationReport();
        if (postAuthReport.summary.criticalDenied.length > 0) {
          // User didn't grant all critical scopes
          const stillMissing = postAuthReport.summary.criticalDenied
            .map(k => postAuthReport.features[k].name)
            .join(', ');

          ui.alert(
            'Authorization Incomplete',
            `The following critical permissions were not granted: ${stillMissing}\n\n` +
            'The application may not function correctly.',
            ui.ButtonSet.OK
          );
          return false;
        }

        return true;
      } catch (e) {
        // User cancelled or error occurred
        return false;
      }
    }

    return false;
  } catch (e) {
    // If we can't even check authorization, return false
    return false;
  }
}

/**
 * Shows a detailed authorization status dialog.
 * Menu handler for "Check Authorization Status".
 */
function showAuthorizationStatus() {
  const report = getAuthorizationReport();
  const ui = SpreadsheetApp.getUi();

  let message = 'Authorization Status Report\n';
  message += 'Generated: ' + report.timestamp + '\n\n';

  // Summary
  message += `Features: ${report.summary.authorizedFeatures}/${report.summary.totalFeatures} authorized\n\n`;

  // Feature details
  message += 'Feature Status:\n';
  message += '-'.repeat(45) + '\n';

  for (const [key, feature] of Object.entries(report.features)) {
    const status = feature.authorized ? '[OK]' : '[X]';
    const criticalTag = feature.critical ? ' [REQUIRED]' : '';
    message += `${status} ${feature.name}${criticalTag}\n`;

    if (!feature.authorized && feature.missingScopes.length > 0) {
      const scopeNames = feature.missingScopes.map(s => s.split('/').pop());
      message += `    Missing: ${scopeNames.join(', ')}\n`;
    }
  }

  // Scope details
  message += '\n' + '-'.repeat(45) + '\n';
  message += 'Individual Scopes:\n';

  for (const [name, scope] of Object.entries(report.scopes)) {
    const status = scope.granted ? '[OK]' : '[X]';
    message += `${status} ${name}\n`;
  }

  if (report.summary.deniedFeatures > 0) {
    message += '\n' + '-'.repeat(45) + '\n';
    message += 'Use "Authorize Script" to grant missing permissions.';
  }

  ui.alert('Authorization Status', message, ui.ButtonSet.OK);
}

/**
 * Prompts user to authorize missing scopes.
 * Menu handler for "Authorize Script".
 */
function promptForAuthorizationIfNeeded() {
  const report = getAuthorizationReport();

  // Check if everything is already authorized
  if (report.summary.deniedFeatures === 0) {
    SpreadsheetApp.getUi().alert(
      'Already Authorized',
      'All permissions are granted. All features are available.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ui = SpreadsheetApp.getUi();

  // Build message showing what's missing
  let message = 'The following features need authorization:\n\n';

  for (const [key, feature] of Object.entries(report.features)) {
    if (!feature.authorized) {
      const tag = feature.critical ? ' (Required)' : ' (Optional)';
      message += `- ${feature.name}${tag}\n`;
      message += `  ${feature.description}\n`;
    }
  }

  message += '\nClick OK to authorize. You will be prompted to grant permissions.';

  const response = ui.alert('Authorization Required', message, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    // Collect all missing scopes
    const missingScopes = new Set();
    for (const feature of Object.values(report.features)) {
      if (!feature.authorized) {
        feature.missingScopes.forEach(s => missingScopes.add(s));
      }
    }

    if (missingScopes.size > 0) {
      try {
        // This will trigger the OAuth prompt
        ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, Array.from(missingScopes));

        // IMPORTANT: Re-check after requireScopes - with granular consent,
        // user may have granted some scopes but denied others
        const postAuthReport = getAuthorizationReport();

        if (postAuthReport.summary.deniedFeatures === 0) {
          // All features now authorized
          ui.alert(
            'Authorization Complete',
            'All permissions have been granted. All features are now available.',
            ui.ButtonSet.OK
          );
        } else if (postAuthReport.summary.deniedFeatures < report.summary.deniedFeatures) {
          // Some features were authorized
          const nowAvailable = Object.entries(postAuthReport.features)
            .filter(([key, f]) => f.authorized && !report.features[key].authorized)
            .map(([_, f]) => f.name);

          const stillMissing = Object.entries(postAuthReport.features)
            .filter(([_, f]) => !f.authorized)
            .map(([_, f]) => f.name);

          let resultMsg = 'Authorization partially complete.\n\n';
          if (nowAvailable.length > 0) {
            resultMsg += `Now available: ${nowAvailable.join(', ')}\n\n`;
          }
          if (stillMissing.length > 0) {
            resultMsg += `Still disabled: ${stillMissing.join(', ')}\n`;
            resultMsg += '\nYou can authorize additional features later from this menu.';
          }

          ui.alert('Authorization Updated', resultMsg, ui.ButtonSet.OK);
        } else {
          // No change - user denied all
          ui.alert(
            'No Permissions Granted',
            'No additional permissions were granted. Some features remain disabled.\n\n' +
            'You can try again later from System Admin > Authorization.',
            ui.ButtonSet.OK
          );
        }
      } catch (e) {
        ui.alert(
          'Authorization Cancelled',
          'The authorization process was cancelled.\n\n' +
          'You can try again from System Admin > Authorization > Authorize Script.',
          ui.ButtonSet.OK
        );
      }
    }
  }
}


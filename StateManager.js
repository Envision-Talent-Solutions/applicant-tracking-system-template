/**
 * @file StateManager.gs - Centralized wrapper for DocumentProperties.
 * 
 * SECURITY NOTE: Document Properties are stored in plain text and accessible
 * to all users with edit access. For high-security environments:
 * - Consider using a separate database for sensitive data
 * - Add validation when reading critical properties
 * - Implement audit logging for property changes
 * - Limit edit access to trusted users only
 */

const StateManager = {
  // ---------- Internal Property Access ----------
  
  /**
   * Gets a property value
   * @param {string} key - Property key
   * @returns {string|null} Property value or null
   */
  _getProperty: (key) => {
    try {
      return PropertiesService.getDocumentProperties().getProperty(key);
    } catch (e) {
      logWarn('Failed to get property', { key, error: e.message });
      return null;
    }
  },
  
  /**
   * Sets a property value
   * @param {string} key - Property key
   * @param {string} value - Property value
   */
  _setProperty: (key, value) => {
    try {
      PropertiesService.getDocumentProperties().setProperty(key, value);
    } catch (e) {
      logWarn('Failed to set property', { key, error: e.message });
    }
  },
  
  /**
   * Deletes a property
   * @param {string} key - Property key
   */
  _deleteProperty: (key) => {
    try {
      PropertiesService.getDocumentProperties().deleteProperty(key);
    } catch (e) {
      logWarn('Failed to delete property', { key, error: e.message });
    }
  },
  
  /**
   * Gets all properties
   * @returns {Object} All properties as key-value pairs
   */
  _getProperties: () => {
    try {
      return PropertiesService.getDocumentProperties().getProperties();
    } catch (e) {
      logWarn('Failed to get all properties', { error: e.message });
      return {};
    }
  },

  // ---------- Recursion Guard ----------
  
  /**
   * Sets a flag to prevent recursive function calls
   */
  setRecursionGuard: () => StateManager._setProperty(CACHE_KEYS.REC_GUARD(), '1'),
  
  /**
   * Checks if recursion guard is active
   * @returns {boolean} True if guard is set
   */
  isRecursionGuardSet: () => StateManager._getProperty(CACHE_KEYS.REC_GUARD()) !== null,
  
  /**
   * Removes the recursion guard
   */
  deleteRecursionGuard: () => StateManager._deleteProperty(CACHE_KEYS.REC_GUARD()),

  // ---------- Job ID Sequencer ----------
  
  /**
   * Gets the current job ID sequence number for a year
   * @param {number} year - The year (e.g., 2025)
   * @returns {number} Current sequence number
   */
  getJobIdSequence: (year) => {
    const key = `${PROP_JOBSEQ_PREFIX}${year}`;
    const value = StateManager._getProperty(key);
    return Number(value || '0');
  },
  
  /**
   * Sets the job ID sequence number for a year
   * @param {number} year - The year (e.g., 2025)
   * @param {number} value - The sequence number
   */
  setJobIdSequence: (year, value) => {
    const key = `${PROP_JOBSEQ_PREFIX}${year}`;
    StateManager._setProperty(key, String(value));
  },
  
  /**
   * Checks if job ID sequence has been initialized for a year
   * @param {number} year - The year to check
   * @returns {boolean} True if initialized
   */
  isJobIdSequenceInitialized: (year) => {
    const key = `${PROP_JOBSEQ_PREFIX}${year}`;
    return StateManager._getProperty(key) !== null;
  },

  // ---------- Link Hygiene Dirty Flags ----------
  
  /**
   * Marks a link cell as needing cleanup
   * @param {string} sheetName - Name of the sheet
   * @param {string} header - Header name
   * @param {number} row - Row number
   */
  markLinkDirty: (sheetName, header, row) => {
    const key = `ATS:linkdirty:${sheetName}:${header}:${row}`;
    StateManager._setProperty(key, String(Date.now()));
  },
  
  /**
   * Gets all dirty link properties
   * @returns {Object} Object with dirty link keys and timestamps
   */
  getDirtyLinkProperties: () => {
    const allProps = StateManager._getProperties();
    const dirtyProps = {};
    for (const key in allProps) {
      if (key.startsWith('ATS:linkdirty:')) {
        dirtyProps[key] = allProps[key];
      }
    }
    return dirtyProps;
  },
  
  /**
   * Deletes a specific dirty link property
   * @param {string} key - The property key to delete
   */
  deleteLinkDirtyProperty: (key) => StateManager._deleteProperty(key),
  
  // ---------- Google Form ID ----------
  
  /**
   * Stores the Google Form ID for candidate submissions
   * @param {string} id - The form ID
   */
  setFormId: (id) => StateManager._setProperty(PROP_FORM_ID, id),
  
  /**
   * Retrieves the stored Google Form ID
   * @returns {string|null} The form ID or null
   */
  getFormId: () => StateManager._getProperty(PROP_FORM_ID),

  // ---------- Settings Hash (for Validation Optimization) ----------
  
  /**
   * Gets the stored hash of SETTINGS sheet configuration
   * Used to detect if validations need rebuilding
   * @returns {string|null} The settings hash or null
   */
  getSettingsHash: () => StateManager._getProperty(PROP_SETTINGS_HASH),
  
  /**
   * Stores the hash of SETTINGS sheet configuration
   * @param {string} hash - The new settings hash
   */
  setSettingsHash: (hash) => StateManager._setProperty(PROP_SETTINGS_HASH, hash),

  // ---------- Generic Property Access ----------
  
  /**
   * Generic property setter (for queue operations, etc.)
   * @param {string} key - Property key
   * @param {string} value - Property value
   */
  setProperty: (key, value) => StateManager._setProperty(key, value),
  
  /**
   * Generic property getter
   * @param {string} key - Property key
   * @returns {string|null} Property value or null
   */
  getProperty: (key) => StateManager._getProperty(key),
  
  /**
   * Generic property deleter
   * @param {string} key - Property key
   */
  deleteProperty: (key) => StateManager._deleteProperty(key),
};
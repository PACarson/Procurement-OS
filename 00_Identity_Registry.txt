/**
 * ============================================================
 * SHARED KERNEL — 00_Identity_Registry.gs
 * ============================================================
 * Scope        : Cross-OS shared module (Layer 1 — Identity,
 *                 per Domain OS Blueprint V2)
 * Consumers    : Inventory OS (21–26), Procurement OS (60–69),
 *                 and any future Domain OS
 * Version      : V0.1 (Lightweight Identity Stub)
 * Last Updated : 2026-06-29
 *
 * PURPOSE
 * ------------------------------------------------------------
 * Provide ONE shared, minimal canonical-identity lookup so that
 * Inventory OS / Procurement OS (and future OS) never invent
 * their own item-name matching logic.
 *
 * "Everything is an Identity before it becomes a record."
 *
 * THIS IS DELIBERATELY MINIMAL
 * ------------------------------------------------------------
 * - NO SKU system
 * - NO barcode system
 * - NO fuzzy / AI-based name normalization
 * - Matching is exact, case-insensitive, trimmed string match
 *   against canonical_name + aliases only.
 *
 * This module is a TEMPORARY foundation. It is designed to be
 * swapped for a full Identity Engine V2 later WITHOUT changing
 * the public function signatures below (resolve / getById /
 * findByName / addAlias / list). Callers must only ever use
 * these functions — never read/write the IDENTITY_REGISTRY
 * sheet directly (same Layered Access rule as every other OS).
 *
 * EVENT SOURCING / CQRS COMPATIBILITY
 * ------------------------------------------------------------
 * - IDENTITY_REGISTRY sheet is treated as the Read Model.
 * - create() is a single, append-only write — identity_id and
 *   canonical_name are never mutated in place once created.
 * - addAlias() only ever appends to the aliases list.
 * - This keeps the data shape replay-safe: a future Identity
 *   Engine V2 can treat every row here as the equivalent of a
 *   single IDENTITY_CREATED event and rebuild on top of it.
 *
 * DOES NOT BLOCK INVENTORY OS
 * ------------------------------------------------------------
 * resolve() is a single call, no setup required beyond this
 * file. 21_InventoryEngine.gs and 60_ProcurementEngine.gs can
 * start calling it immediately.
 * ============================================================
 */

var IdentityRegistry = (function () {

  var SHEET_NAME = 'IDENTITY_REGISTRY';
  var HEADERS = ['identity_id', 'canonical_name', 'aliases', 'category', 'unit', 'created_at'];
  var ID_PREFIX = 'ID-';
  var ID_PAD = 6;
  var ALIAS_DELIMITER = ' | ';

  // ----------------------------------------------------------
  // Internal: Sheet access
  // (ONLY this module touches SpreadsheetApp — Layered Access)
  // ----------------------------------------------------------

  function _getSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  function _readAllRows_() {
    var sheet = _getSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    // Batch read — never read row-by-row / cell-by-cell.
    return sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  }

  // ----------------------------------------------------------
  // Internal: serialization helpers
  // ----------------------------------------------------------

  function _parseAliases_(raw) {
    if (!raw) return [];
    return String(raw)
      .split('|')
      .map(function (s) { return s.trim(); })
      .filter(Boolean);
  }

  function _serializeAliases_(aliases) {
    return (aliases || []).join(ALIAS_DELIMITER);
  }

  function _normalize_(text) {
    return String(text || '').trim().toLowerCase();
  }

  function _rowToRecord_(row) {
    return {
      identity_id: row[0],
      canonical_name: row[1],
      aliases: _parseAliases_(row[2]),
      category: row[3],
      unit: row[4],
      created_at: row[5]
    };
  }

  function _nextIdentityId_(rows) {
    var max = 0;
    rows.forEach(function (row) {
      var match = /^ID-(\d+)$/.exec(row[0]);
      if (match) {
        var n = parseInt(match[1], 10);
        if (n > max) max = n;
      }
    });
    return ID_PREFIX + String(max + 1).padStart(ID_PAD, '0');
  }

  // ----------------------------------------------------------
  // Public: lookups
  // ----------------------------------------------------------

  /** Get a single identity by identity_id. Returns null if not found. */
  function getById(identityId) {
    if (!identityId) return null;
    var rows = _readAllRows_();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === identityId) return _rowToRecord_(rows[i]);
    }
    return null;
  }

  /**
   * Find an identity by canonical_name OR any alias.
   * Exact match only (case-insensitive, trimmed). No fuzzy matching.
   * Returns null if not found.
   */
  function findByName(name) {
    var target = _normalize_(name);
    if (!target) return null;
    var rows = _readAllRows_();
    for (var i = 0; i < rows.length; i++) {
      var record = _rowToRecord_(rows[i]);
      if (_normalize_(record.canonical_name) === target) return record;
      for (var j = 0; j < record.aliases.length; j++) {
        if (_normalize_(record.aliases[j]) === target) return record;
      }
    }
    return null;
  }

  /** Return all identities. Use sparingly — admin / debug only. */
  function list() {
    return _readAllRows_().map(_rowToRecord_);
  }

  // ----------------------------------------------------------
  // Public: writes
  // ----------------------------------------------------------

  /**
   * Create a brand-new identity stub. Most callers should use
   * resolve() instead, which avoids creating duplicates.
   */
  function create(input) {
    input = input || {};
    var name = input.canonicalName || input.canonical_name;
    if (!name) {
      throw new Error('IdentityRegistry.create() requires canonicalName');
    }

    var rows = _readAllRows_();
    var record = {
      identity_id: _nextIdentityId_(rows),
      canonical_name: name,
      aliases: input.aliases || [],
      category: input.category || '',
      unit: input.unit || '',
      created_at: new Date().toISOString()
    };

    _getSheet_().appendRow([
      record.identity_id,
      record.canonical_name,
      _serializeAliases_(record.aliases),
      record.category,
      record.unit,
      record.created_at
    ]);

    return record;
  }

  /**
   * Add a new alias to an existing identity (no-op if it already
   * exists). Returns the updated record, or null if identityId
   * was not found.
   */
  function addAlias(identityId, alias) {
    if (!identityId || !alias) return getById(identityId);

    var sheet = _getSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    var range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    var rows = range.getValues();

    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === identityId) {
        var aliases = _parseAliases_(rows[i][2]);
        var norm = _normalize_(alias);
        var alreadyExists = aliases.some(function (a) { return _normalize_(a) === norm; });
        if (!alreadyExists) {
          aliases.push(alias);
          sheet.getRange(2 + i, 3).setValue(_serializeAliases_(aliases));
        }
        return _rowToRecord_(sheet.getRange(2 + i, 1, 1, HEADERS.length).getValues()[0]);
      }
    }
    return null;
  }

  /**
   * PRIMARY ENTRY POINT for Inventory OS / Procurement OS.
   *
   * Resolve a canonical identity from a name. If no match exists
   * (against canonical_name or aliases), auto-create a stub
   * identity and return it.
   *
   * input = {
   *   canonicalName: string   (required)
   *   aliases:       string[] (optional — also checked, and
   *                            attached if a new stub is created)
   *   category:      string   (optional)
   *   unit:          string   (optional)
   * }
   */
  function resolve(input) {
    input = input || {};
    var name = input.canonicalName || input.canonical_name;
    if (!name) {
      throw new Error('IdentityRegistry.resolve() requires canonicalName');
    }

    var existing = findByName(name);
    if (existing) return existing;

    if (input.aliases && input.aliases.length) {
      for (var i = 0; i < input.aliases.length; i++) {
        var byAlias = findByName(input.aliases[i]);
        if (byAlias) return byAlias;
      }
    }

    // No match anywhere — auto-create a lightweight stub.
    return create(input);
  }

  // ----------------------------------------------------------
  // Public API
  // ----------------------------------------------------------
  return {
    resolve: resolve,
    getById: getById,
    findByName: findByName,
    create: create,
    addAlias: addAlias,
    list: list
  };

})();

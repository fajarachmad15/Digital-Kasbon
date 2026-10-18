/**
 * ============================================================================
 * MODULE: SupabaseClient.gs
 * PROJECT: Digital Kasbon Petty Cash - Enterprise GAS Suite
 * DESCRIPTION: Lightweight PostgREST HTTP client via UrlFetchApp
 * ============================================================================
 */

const SupabaseClient = (function () {
  function getHeaders(extraHeaders) {
    const apiKey = AppConfig.get('SUPABASE_KEY');
    const headers = {
      'apikey': apiKey,
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    };
    if (extraHeaders) {
      for (const k in extraHeaders) {
        if (Object.prototype.hasOwnProperty.call(extraHeaders, k)) {
          headers[k] = extraHeaders[k];
        }
      }
    }
    return headers;
  }

  function buildUrl(tableOrView, queryString) {
    const baseUrl = AppConfig.get('SUPABASE_URL').replace(/\/$/, '');
    const qs = queryString ? ('?' + queryString.replace(/^\?/, '')) : '';
    return baseUrl + '/rest/v1/' + tableOrView + qs;
  }

  /**
   * SELECT rows from a table or view.
   * @param {string} tableOrView - e.g. 'v_kasbon_dashboard' or 'kasbon_requests'
   * @param {string} queryString - PostgREST filter e.g. 'kode_store=eq.A644&order=timestamp_req.desc'
   * @return {Array<Object>}
   */
  function select(tableOrView, queryString) {
    const url = buildUrl(tableOrView, queryString);
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: getHeaders(),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : [];
    }
    throw new Error('Supabase SELECT Error (' + code + '): ' + text);
  }

  /**
   * SELECT a single row or return null.
   */
  function selectSingle(tableOrView, queryString) {
    const rows = select(tableOrView, queryString + '&limit=1');
    return rows && rows.length > 0 ? rows[0] : null;
  }

  /**
   * INSERT one or more rows into a table.
   * @param {string} table
   * @param {Object|Array<Object>} payload
   * @param {string=} onConflictCol - optional column for upsert
   * @return {Array<Object>}
   */
  function insert(table, payload, onConflictCol) {
    const qs = onConflictCol ? ('on_conflict=' + encodeURIComponent(onConflictCol)) : '';
    const url = buildUrl(table, qs);
    const preferHeader = onConflictCol
      ? 'resolution=merge-duplicates,return=representation'
      : 'return=representation';

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: getHeaders({ 'Prefer': preferHeader }),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : [];
    }
    throw new Error('Supabase INSERT Error (' + code + '): ' + text);
  }

  /**
   * PATCH (Update) rows matching a PostgREST query filter.
   * @param {string} table
   * @param {string} matchQuery - e.g. 'no_pengajuan=eq.KBA380-1225-001'
   * @param {Object} updatePayload
   * @return {Array<Object>}
   */
  function update(table, matchQuery, updatePayload) {
    if (!matchQuery) {
      throw new Error('Supabase UPDATE requires a non-empty matchQuery filter.');
    }
    const url = buildUrl(table, matchQuery);
    const response = UrlFetchApp.fetch(url, {
      method: 'patch',
      headers: getHeaders({ 'Prefer': 'return=representation' }),
      payload: JSON.stringify(updatePayload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : [];
    }
    throw new Error('Supabase UPDATE Error (' + code + '): ' + text);
  }

  /**
   * DELETE rows matching a PostgREST query filter.
   * @param {string} table
   * @param {string} matchQuery - e.g. 'id=eq.fae8e2b2-610d-47e6-8949-57ddb5116aaf'
   * @return {Array<Object>}
   */
  function remove(table, matchQuery) {
    if (!matchQuery) {
      throw new Error('Supabase DELETE requires a non-empty matchQuery filter.');
    }
    const url = buildUrl(table, matchQuery);
    const response = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: getHeaders({ 'Prefer': 'return=representation' }),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : [];
    }
    throw new Error('Supabase DELETE Error (' + code + '): ' + text);
  }

  /**
   * Writes an audit trail entry into kasbon_audit_logs (non-blocking on error).
   */
  function logAudit(entry) {
    try {
      insert('kasbon_audit_logs', {
        no_pengajuan: entry.no_pengajuan,
        step_number: entry.step_number || 1,
        action_type: entry.action_type || 'UPDATE',
        actor_nik: entry.actor_nik || '-',
        actor_name: entry.actor_name || '-',
        actor_email: entry.actor_email || '-',
        actor_role: entry.actor_role || '-',
        previous_status: entry.previous_status || '-',
        new_status: entry.new_status || '-',
        notes: entry.notes || '-'
      });
    } catch (e) {
      console.warn('Audit log warning:', e.message);
    }
  }

  /**
   * Writes a master data / admin audit trail entry into kasbon_admin_logs (non-blocking on error).
   */
  function logAdminAudit(entry) {
    try {
      insert('kasbon_admin_logs', {
        action_type: entry.action_type || 'ADMIN_ACTION',
        target_entity: entry.target_entity || 'SYSTEM',
        target_id: entry.target_id || '-',
        actor_nik: entry.actor_nik || '-',
        actor_name: entry.actor_name || '-',
        actor_email: entry.actor_email || '-',
        details: entry.details || '-'
      });
    } catch (e) {
      console.warn('Admin audit log warning:', e.message);
    }
  }

  return {
    select: select,
    selectSingle: selectSingle,
    insert: insert,
    update: update,
    remove: remove,
    logAudit: logAudit,
    logAdminAudit: logAdminAudit
  };
})();

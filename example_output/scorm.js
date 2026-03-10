
/**
 * SCORM 2004 3rd Edition API Wrapper
 * Finds the LMS API in parent frames and exposes a clean interface.
 */
const SCORM = (() => {
  let _api = null;
  let _initialized = false;
  let _terminated = false;

  function findAPI(win) {
    let attempts = 0;
    while (win && !win.API_1484_11 && attempts < 10) {
      if (win.parent && win.parent !== win) {
        win = win.parent;
      } else {
        break;
      }
      attempts++;
    }
    return (win && win.API_1484_11) ? win.API_1484_11 : null;
  }

  function getAPI() {
    if (_api) return _api;
    _api = findAPI(window);
    if (!_api && window.opener) _api = findAPI(window.opener);
    return _api;
  }

  function init() {
    if (_initialized || _terminated) return;
    const api = getAPI();
    if (!api) {
      console.warn('[SCORM] No LMS API found – running in standalone mode.');
      return;
    }
    const result = api.Initialize('');
    if (result === 'true' || result === true) {
      _initialized = true;
      // Only set defaults if not resuming a previous session
      const existingStatus = api.GetValue('cmi.completion_status');
      if (!existingStatus || existingStatus === 'not attempted' || existingStatus === 'unknown') {
        api.SetValue('cmi.completion_status', 'incomplete');
        api.SetValue('cmi.success_status', 'unknown');
        api.SetValue('cmi.progress_measure', '0');
      }
      api.Commit('');
      console.log('[SCORM] Initialized.');
    } else {
      console.error('[SCORM] Initialize failed:', api.GetLastError(), api.GetErrorString(api.GetLastError()));
    }
  }

  function setProgress(measure) {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    const val = Math.min(1, Math.max(0, measure));
    api.SetValue('cmi.progress_measure', String(val.toFixed(4)));
    api.Commit('');
    console.log('[SCORM] Progress:', (val * 100).toFixed(1) + '%');
  }

  function setLocation(loc) {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    api.SetValue('cmi.location', String(loc));
    api.Commit('');
  }

  function getLocation() {
    if (!_initialized) return '';
    const api = getAPI();
    if (!api) return '';
    return api.GetValue('cmi.location') || '';
  }

  function setSuspendData(data) {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    api.SetValue('cmi.suspend_data', data);
    api.Commit('');
  }

  function getSuspendData() {
    if (!_initialized) return '';
    const api = getAPI();
    if (!api) return '';
    return api.GetValue('cmi.suspend_data') || '';
  }

  function complete() {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    api.SetValue('cmi.progress_measure', '1');
    api.SetValue('cmi.completion_status', 'completed');
    api.SetValue('cmi.success_status', 'passed');
    api.Commit('');
    console.log('[SCORM] Completed!');
  }

  function terminate() {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    _terminated = true;
    api.Terminate('');
    console.log('[SCORM] Terminated.');
  }

  return { init, setProgress, setLocation, getLocation, setSuspendData, getSuspendData, complete, terminate };
})();

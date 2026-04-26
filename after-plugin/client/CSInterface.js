/**
 * Lion Workspace — Professional CSInterface Wrapper
 * Inspired by AutoCut's resilient architecture patterns:
 *   - Retry with exponential backoff on evalScript failures
 *   - Request queue to prevent concurrent evalScript flooding
 *   - Timeout handling for hung ExtendScript calls
 *   - Health monitoring with connection state tracking
 *   - Step-based initialization logging
 */

(function(global) {
    'use strict';

    // ─── Logger ───────────────────────────────────────────────
    var LOG_LEVELS = { debug: 0, info: 1, warn: 2, error: 3 };
    var LOG_COLORS = {
        debug: 'color:#636363',
        info:  'color:#bef264',
        warn:  'color:#f59e0b',
        error: 'color:#f43f5e'
    };
    var currentLogLevel = LOG_LEVELS.info;

    var LWLog = {
        setLevel: function(level) { currentLogLevel = LOG_LEVELS[level] || 0; },
        _log: function(level, category, msg, data) {
            if (LOG_LEVELS[level] < currentLogLevel) return;
            var prefix = '[LW:' + category + ']';
            var style = LOG_COLORS[level] || '';
            if (data !== undefined) {
                console.log('%c' + prefix + ' ' + msg, style, data);
            } else {
                console.log('%c' + prefix + ' ' + msg, style);
            }
        },
        debug: function(cat, msg, data) { this._log('debug', cat, msg, data); },
        info:  function(cat, msg, data) { this._log('info',  cat, msg, data); },
        warn:  function(cat, msg, data) { this._log('warn',  cat, msg, data); },
        error: function(cat, msg, data) { this._log('error', cat, msg, data); }
    };

    global.LWLog = LWLog;

    // ─── Step-based Init Tracker (like AutoCut) ──────────────
    var initSteps = [];
    var InitTracker = {
        currentStep: null,
        start: function(name) {
            this.currentStep = name;
            var entry = { name: name, startedAt: Date.now(), status: 'running' };
            initSteps.push(entry);
            LWLog.info('init', '→ ' + name);
            return entry;
        },
        complete: function(name) {
            for (var i = initSteps.length - 1; i >= 0; i--) {
                if (initSteps[i].name === name) {
                    initSteps[i].status = 'done';
                    initSteps[i].duration = Date.now() - initSteps[i].startedAt;
                    LWLog.info('init', '✓ ' + name + ' (' + initSteps[i].duration + 'ms)');
                    break;
                }
            }
            this.currentStep = null;
        },
        fail: function(name, error) {
            for (var i = initSteps.length - 1; i >= 0; i--) {
                if (initSteps[i].name === name) {
                    initSteps[i].status = 'failed';
                    initSteps[i].error = error;
                    LWLog.error('init', '✗ ' + name + ': ' + error);
                    break;
                }
            }
            this.currentStep = null;
        },
        getSteps: function() { return initSteps.slice(); },
        getSummary: function() {
            return initSteps.map(function(s) {
                return s.status + ' ' + s.name + (s.duration ? ' (' + s.duration + 'ms)' : '');
            }).join('\n');
        }
    };

    global.LWInit = InitTracker;

    // ─── CSInterface Core ────────────────────────────────────
    var MAX_RETRIES = 3;
    var RETRY_BASE_DELAY = 200; // ms, exponential: 200, 400, 800
    var EVAL_TIMEOUT = 15000;   // 15s timeout for evalScript
    var QUEUE_CONCURRENCY = 1;  // sequential execution

    var _cepAvailable = !!(global.__adobe_cep__);
    var _queue = [];
    var _queueRunning = 0;
    var _healthState = {
        cepOk: _cepAvailable,
        lastSuccessAt: 0,
        lastFailAt: 0,
        consecutiveFails: 0,
        totalCalls: 0,
        totalRetries: 0,
        totalTimeouts: 0
    };

    function CSInterface() {}

    // ─── Health API ──────────────────────────────────────────
    CSInterface.prototype.getHealth = function() {
        return {
            cepAvailable: _cepAvailable,
            healthy: _healthState.consecutiveFails < 3,
            lastSuccess: _healthState.lastSuccessAt,
            lastFail: _healthState.lastFailAt,
            consecutiveFails: _healthState.consecutiveFails,
            stats: {
                totalCalls: _healthState.totalCalls,
                totalRetries: _healthState.totalRetries,
                totalTimeouts: _healthState.totalTimeouts
            }
        };
    };

    CSInterface.prototype.isHealthy = function() {
        return _cepAvailable && _healthState.consecutiveFails < 5;
    };

    // ─── Core evalScript with retry + timeout ────────────────
    function _rawEvalScript(script, callback) {
        if (!global.__adobe_cep__) {
            if (callback) callback('CEP_NOT_AVAILABLE');
            return;
        }
        try {
            global.__adobe_cep__.evalScript(script, callback || function() {});
        } catch (e) {
            LWLog.error('cep', 'evalScript threw: ' + e.message);
            if (callback) callback('EVAL_EXCEPTION:' + e.message);
        }
    }

    function _isErrorResult(result) {
        if (!result) return false;
        var r = String(result);
        return r === 'EvalScript_Error' ||
               r === 'CEP_NOT_AVAILABLE' ||
               r.indexOf('EVAL_EXCEPTION:') === 0;
    }

    function _evalWithRetry(script, callback, retries, attempt) {
        retries = (typeof retries === 'number') ? retries : MAX_RETRIES;
        attempt = attempt || 0;
        _healthState.totalCalls++;

        var timedOut = false;
        var responded = false;

        // Timeout guard
        var timer = setTimeout(function() {
            if (responded) return;
            timedOut = true;
            _healthState.totalTimeouts++;
            _healthState.consecutiveFails++;
            _healthState.lastFailAt = Date.now();
            LWLog.warn('cep', 'evalScript timeout (' + EVAL_TIMEOUT + 'ms): ' + script.substring(0, 60));

            if (attempt < retries) {
                var delay = RETRY_BASE_DELAY * Math.pow(2, attempt);
                _healthState.totalRetries++;
                LWLog.info('cep', 'Retry ' + (attempt + 1) + '/' + retries + ' in ' + delay + 'ms');
                setTimeout(function() {
                    _evalWithRetry(script, callback, retries, attempt + 1);
                }, delay);
            } else {
                LWLog.error('cep', 'All retries exhausted for: ' + script.substring(0, 60));
                if (callback) callback('TIMEOUT');
            }
        }, EVAL_TIMEOUT);

        _rawEvalScript(script, function(result) {
            if (timedOut) return; // Already timed out, ignore late response
            responded = true;
            clearTimeout(timer);

            if (_isErrorResult(result) && attempt < retries) {
                var delay = RETRY_BASE_DELAY * Math.pow(2, attempt);
                _healthState.totalRetries++;
                _healthState.consecutiveFails++;
                _healthState.lastFailAt = Date.now();
                LWLog.warn('cep', 'evalScript error, retry ' + (attempt + 1) + '/' + retries + ' in ' + delay + 'ms: ' + result);
                setTimeout(function() {
                    _evalWithRetry(script, callback, retries, attempt + 1);
                }, delay);
            } else {
                if (!_isErrorResult(result)) {
                    _healthState.consecutiveFails = 0;
                    _healthState.lastSuccessAt = Date.now();
                }
                if (callback) callback(result);
            }
        });
    }

    // ─── Queue System ────────────────────────────────────────
    function _processQueue() {
        if (_queueRunning >= QUEUE_CONCURRENCY || _queue.length === 0) return;
        _queueRunning++;

        var item = _queue.shift();
        _evalWithRetry(item.script, function(result) {
            _queueRunning--;
            if (item.callback) item.callback(result);
            _processQueue(); // Process next in queue
        }, item.retries);
    }

    // ─── Public evalScript (queued + retry) ──────────────────
    CSInterface.prototype.evalScript = function(script, callback, options) {
        options = options || {};

        if (!_cepAvailable) {
            LWLog.debug('cep', 'CEP not available, skipping: ' + script.substring(0, 40));
            if (callback) callback('CEP_NOT_AVAILABLE');
            return;
        }

        if (options.immediate) {
            // Skip queue for critical calls (e.g., cancel operations)
            _evalWithRetry(script, callback, options.retries);
        } else {
            _queue.push({
                script: script,
                callback: callback,
                retries: (typeof options.retries === 'number') ? options.retries : MAX_RETRIES
            });
            _processQueue();
        }
    };

    // Backwards-compatible simple eval (no queue, single retry)
    CSInterface.prototype.evalScriptSimple = function(script, callback) {
        _rawEvalScript(script, callback);
    };

    // ─── System Paths ────────────────────────────────────────
    CSInterface.prototype.getSystemPath = function(type) {
        try {
            var path = global.__adobe_cep__.getSystemPath(type);
            return path;
        } catch (e) {
            LWLog.error('cep', 'getSystemPath failed: ' + e.message);
            return '';
        }
    };

    CSInterface.SYSTEM_PATH = {
        USER_DATA: 'userData',
        COMMON_FILES: 'commonFiles',
        MY_DOCUMENTS: 'myDocuments',
        APPLICATION: 'application',
        EXTENSION: 'extension',
        HOST_APPLICATION: 'hostApplication'
    };

    // ─── Resilient HTTP Client (like AutoCut's axios layer) ──
    var LWHttp = {
        MAX_RETRIES: 3,
        RETRY_DELAY: 500,
        _consecutiveFails: 0,
        _lastSuccessAt: 0,

        request: function(method, url, data, options, callback) {
            if (typeof options === 'function') { callback = options; options = {}; }
            options = options || {};
            var retries = (typeof options.retries === 'number') ? options.retries : this.MAX_RETRIES;
            var timeout = options.timeout || 10000;
            var attempt = 0;
            var self = this;

            function doRequest() {
                var xhr = new XMLHttpRequest();
                xhr.timeout = timeout;

                xhr.onload = function() {
                    self._consecutiveFails = 0;
                    self._lastSuccessAt = Date.now();
                    if (callback) callback(null, xhr.status, xhr.responseText);
                };

                xhr.onerror = function() {
                    self._consecutiveFails++;
                    if (attempt < retries) {
                        attempt++;
                        var delay = self.RETRY_DELAY * Math.pow(2, attempt - 1);
                        LWLog.warn('http', 'Request failed, retry ' + attempt + '/' + retries + ' in ' + delay + 'ms: ' + url);
                        setTimeout(doRequest, delay);
                    } else {
                        LWLog.error('http', 'All retries failed: ' + url);
                        if (callback) callback('NETWORK_ERROR', 0, null);
                    }
                };

                xhr.ontimeout = function() {
                    self._consecutiveFails++;
                    if (attempt < retries) {
                        attempt++;
                        var delay = self.RETRY_DELAY * Math.pow(2, attempt - 1);
                        LWLog.warn('http', 'Request timeout, retry ' + attempt + '/' + retries + ': ' + url);
                        setTimeout(doRequest, delay);
                    } else {
                        LWLog.error('http', 'All retries timed out: ' + url);
                        if (callback) callback('TIMEOUT', 0, null);
                    }
                };

                xhr.open(method, url, true);

                // Set headers
                if (options.headers) {
                    for (var key in options.headers) {
                        if (options.headers.hasOwnProperty(key)) {
                            xhr.setRequestHeader(key, options.headers[key]);
                        }
                    }
                }

                if (data && typeof data === 'object' && !(data instanceof FormData)) {
                    xhr.setRequestHeader('Content-Type', 'application/json');
                    xhr.send(JSON.stringify(data));
                } else {
                    xhr.send(data || null);
                }
            }

            doRequest();
        },

        get: function(url, options, callback) {
            this.request('GET', url, null, options, callback);
        },

        post: function(url, data, options, callback) {
            this.request('POST', url, data, options, callback);
        },

        isHealthy: function() {
            return this._consecutiveFails < 5;
        }
    };

    global.LWHttp = LWHttp;

    // ─── Connection Manager (adaptive polling) ───────────────
    var LWConnection = {
        _state: 'disconnected', // 'connected', 'disconnected', 'reconnecting'
        _pollInterval: 1000,    // Normal: 1s
        _maxPollInterval: 5000, // When disconnected: 5s
        _minPollInterval: 1000,
        _listeners: [],
        _apiBase: '',

        init: function(apiBase) {
            this._apiBase = apiBase;
            LWLog.info('conn', 'Connection manager initialized: ' + apiBase);
        },

        getState: function() { return this._state; },

        getPollInterval: function() { return this._pollInterval; },

        onStateChange: function(fn) { this._listeners.push(fn); },

        _notify: function(newState, oldState) {
            for (var i = 0; i < this._listeners.length; i++) {
                try { this._listeners[i](newState, oldState); } catch (e) {}
            }
        },

        markConnected: function() {
            var old = this._state;
            if (old !== 'connected') {
                this._state = 'connected';
                this._pollInterval = this._minPollInterval;
                LWLog.info('conn', 'Connected to Lion Workspace');
                this._notify('connected', old);
            }
        },

        markDisconnected: function() {
            var old = this._state;
            if (old !== 'disconnected') {
                this._state = 'disconnected';
                // Adaptive: increase poll interval when disconnected (saves resources)
                this._pollInterval = Math.min(this._pollInterval * 1.5, this._maxPollInterval);
                LWLog.warn('conn', 'Disconnected (poll interval: ' + Math.round(this._pollInterval) + 'ms)');
                this._notify('disconnected', old);
            } else {
                // Already disconnected, keep increasing interval
                this._pollInterval = Math.min(this._pollInterval * 1.2, this._maxPollInterval);
            }
        },

        // Reset poll interval back to fast when user interacts
        resetPollRate: function() {
            this._pollInterval = this._minPollInterval;
        }
    };

    global.LWConnection = LWConnection;

    // ─── Export ──────────────────────────────────────────────
    global.CSInterface = CSInterface;

})(typeof window !== 'undefined' ? window : this);

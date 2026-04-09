/* Minimal CSInterface wrapper for Adobe CEP extensions */
function CSInterface() {}

CSInterface.prototype.evalScript = function(script, callback) {
    if (window.__adobe_cep__) {
        window.__adobe_cep__.evalScript(script, callback || function() {});
    } else if (callback) {
        callback('CEP_NOT_AVAILABLE');
    }
};

CSInterface.prototype.getSystemPath = function(type) {
    try {
        var path = window.__adobe_cep__.getSystemPath(type);
        return path;
    } catch (e) { return ''; }
};

CSInterface.SYSTEM_PATH = {
    USER_DATA: 'userData',
    COMMON_FILES: 'commonFiles',
    MY_DOCUMENTS: 'myDocuments',
    APPLICATION: 'application',
    EXTENSION: 'extension',
    HOST_APPLICATION: 'hostApplication'
};

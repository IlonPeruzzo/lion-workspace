// ============================================================================
// Lion Workspace — After Effects ExtendScript host
// ============================================================================
// O timer, pomodoro e pipeline são controlados via HTTP (porta 9847) que fala
// direto com o app Lion Workspace. Este script só expõe utilidades específicas
// do After Effects (nome do projeto ativo, comp ativa, etc.).

#target aftereffects

// Retorna a versão do CSInterface (sanity check)
function lwAEPing() {
    try {
        var v = parseFloat(app.version) || 0;
        return 'OK:AE' + v;
    } catch (e) {
        return 'ERR:' + e.toString();
    }
}

// Retorna o nome do projeto ativo (.aep) — usado opcionalmente pelo plugin
function lwAEGetProjectName() {
    try {
        if (!app.project || !app.project.file) return '';
        var f = app.project.file;
        var name = f.name.replace(/\.aep$/i, '').replace(/\.aepx$/i, '');
        return decodeURIComponent(name);
    } catch (e) {
        return '';
    }
}

// Retorna o nome da composição ativa (se houver)
function lwAEGetActiveCompName() {
    try {
        if (!app.project || !app.project.activeItem) return '';
        var item = app.project.activeItem;
        if (item instanceof CompItem) return item.name;
        return '';
    } catch (e) {
        return '';
    }
}

// Retorna informações úteis sobre o projeto/comp atual em JSON
function lwAEGetSessionInfo() {
    try {
        var info = {
            projectName: lwAEGetProjectName(),
            activeComp: lwAEGetActiveCompName(),
            aeVersion: parseFloat(app.version) || 0,
            numItems: app.project ? app.project.numItems : 0
        };
        return 'OK:' + JSON.stringify(info);
    } catch (e) {
        return 'ERR:' + e.toString();
    }
}

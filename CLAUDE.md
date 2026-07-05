# Lion Workspace — notas do projeto (Claude Code)

App **Electron** (gestão pra editores de vídeo BR) + **2 plugins CEP** (Adobe Premiere e After Effects).
Dono: **Ilon Peruzzo** (@ilonn no Discord). Idioma: **PT-BR**.

## Estrutura
- `main.js` — processo principal Electron (Node). Sobe um **servidor HTTP local na porta 9847** que os plugins CEP consomem. Também: Discord OAuth, Supabase, spawn do motion tracker (Python), PDF, etc.
- `index.html` — renderer do app (dashboard, clientes, config, remover fundo, etc.).
- `preload.js` — contextBridge (APIs expostas ao renderer).
- `premiere-plugin/client/index.html` + `premiere-plugin/host/index.jsx` — plugin CEP do **Premiere**.
- `after-plugin/client/index.html` + `after-plugin/host/index.jsx` — plugin CEP do **After Effects**.
- `motion_tracker.py` — rastreamento OpenCV (rodado via spawn de Python).
- `build/` — ícones (`icon.png` = logo do leão verde).
- Backend separado (não neste repo): `../lionwork-saas` (Supabase edge functions + dashboard Next.js).

## Como rodar / comandos
- **App em dev:** `npx electron .` na pasta `controle-videos` (carrega `index.html`/`main.js` do source direto).
- **Recarregar o renderer do app:** `Ctrl+R` (habilitado no main.js).
- **Editar plugin → sincronizar pro CEP:** depois de editar `premiere-plugin/client/index.html` ou `after-plugin/client/index.html`, copiar pra:
  - `%APPDATA%\Adobe\CEP\extensions\com.lionworkspace.premiere\client\index.html`
  - `%APPDATA%\Adobe\CEP\extensions\com.lionworkspace.after\client\index.html`
  - Depois **recarregar o painel** (fechar/abrir). O `.jsx` (ExtendScript) só recarrega ao **reiniciar o Premiere/AE**.
- **Debug do plugin CEP:** PlayerDebugMode ligado. Porta de debug do Premiere = **7777** (veja `.debug` na pasta da extensão). DevTools em `http://localhost:7777/json/list`.
- **Testar HTML isolado:** subir `python -m http.server 8766` e abrir no navegador.

## Regras (IMPORTANTES)
- **Só buildar / bumpar versão quando o Ilon mandar.** No resto, mudanças **offline**: editar arquivos + sync pro CEP, sem `npm run build`.
- **Verificar antes de dizer "pronto"** — testar/preview, não só editar e afirmar que funciona.
- PT-BR nas strings de UI (com acento). Comentários de código nos plugins CEP costumam ir **sem acento** (evita bug de encoding).

## Padrões técnicos (pra não reintroduzir bugs conhecidos)
- **CSP de imagem no CEP:** `<img src=http://...>` direto é bloqueado. Use `loadImgViaXhr(url, imgEl)` (baixa como blob via XHR e seta `display:block`). Ex.: avatar do Discord via `API + '/discord-avatar'`. Logo do leão vai embutida como **data URI**.
- **Status unificado dos plugins:** `lwStatusRender(el, text, color)` — ícone SVG + animação + cor por tipo. Mapeia cor→tipo (erro=vermelho/rose, sucesso=lime, aviso=gold, loading=violeta). Texto terminando em `…`/`...` = loading (barrinha). **Fade-out automático ~4s** (loading persiste até ser trocado). Todas as funções de status (`cpShowStatus`, `mtSetStatus`, `_bgMsg`, etc.) passam por ele. Falha = vermelho. Confetti no paste: `lwConfetti(anchorEl)`.
- **Anchor Point / Motion no Premiere (jsx):** `setValue` na Motion intrínseca só re-renderiza se `forceUIRefresh()` cutucar o playhead **um frame inteiro** (NÃO +1 tick). "Compensar posição" seta anchor **e** position juntos pro clip não pular. O toggle clip/sequence é ignorado no código.
- **Copy/Paste — escolher pasta:** use `window.cep.fs.showOpenDialog(false, true, ...)` (API nativa do CEP). NÃO use `Folder.selectDlg` (não existe) nem evalScript.

## Estado recente
- Header dos plugins: avatar Discord + nome + @handle + estrela premium (loop) + leão verde + engrenagem; reorder de painéis por drag nas configs.
- Anchor Point do Premiere corrigido (forceUIRefresh frame inteiro).
- Sistema de status repaginado (cores + ícones circle + check-circle animado + confetti + fade-out).
- Botão "Remover Fundo" (`.magic-cta`) → lime sólido.
- **Em andamento:** tela de login do plugin com a logo do leão em cima + popup "Login OK" com foto (igual ao app).

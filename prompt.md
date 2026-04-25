Refatore o arquivo Index.html de um Google Apps Script dashboard de pedidos do Mercado Livre. Mantenha toda a lógica JavaScript intacta — altere apenas CSS, estrutura HTML e pequenas adições de funcionalidade.

Redesign Visual
Substitua completamente o bloco :root por este sistema de cores:

:root {
  --bg: #F4F6F9; --surface: #FFFFFF; --surface2: #F8F9FB;
  --navy: #1E2D4E; --navy-light: #2C3E6B; --blue: #3B5BDB;
  --yellow: #F5C518; --yellow-light: #FFF8DC; --yellow-dark: #E0A800;
  --gray-50: #F8F9FA; --gray-100: #F1F3F5; --gray-200: #E9ECEF;
  --gray-300: #DEE2E6; --gray-400: #CED4DA; --gray-500: #ADB5BD;
  --gray-600: #6C757D; --gray-700: #495057; --gray-800: #343A40;
  --text: #1A2236; --text-muted: #6C757D;
  --badge-pending-bg: #FFF3CD; --badge-pending-text: #7D5200;
  --badge-confirm-bg: #D1FAE5; --badge-confirm-text: #065F46;
  --badge-cancel-bg:  #FFE4E4; --badge-cancel-text:  #7D0000;
  --badge-pronto-bg:  #DBEAFE; --badge-pronto-text:  #1E3A8A;
  --badge-naofin-bg:  #FEF3C7; --badge-naofin-text:  #78350F;
  --badge-outro-bg:   #F1F3F5; --badge-outro-text:   #495057;
  --shadow-sm: 0 1px 3px rgba(0,0,0,.08), 0 1px 2px rgba(0,0,0,.05);
  --shadow-md: 0 4px 12px rgba(0,0,0,.10), 0 2px 6px rgba(0,0,0,.06);
  --radius: 10px; --radius2: 8px;
}

Remover do HTML e CSS: <div class="bgfx">, <div class="scanlines"> e todos os estilos .bgfx, .scanlines, variáveis --cyan, --magenta, --green, --red, --orange, --glow-*. Remover todos os box-shadow neon/glow.

Redesign dos componentes:

body: background: var(--bg); color: var(--text);
.wrap: max-width: 1280px; padding: 0 16px 48px;
header: fundo branco, sem border-radius, sem blur/glassmorphism. Adicionar border-bottom: 1px solid var(--gray-200); box-shadow: var(--shadow-sm); position: sticky; top: 0;
.badgeLogo: background: var(--navy); border-left: 4px solid var(--yellow); border-radius: 8px; box-shadow: none; (remover todos os radial-gradient e glow)
.brand h1: font-size: 14px; font-weight: 700; color: var(--navy); letter-spacing: 0; text-transform: none;
.chip: background: var(--gray-50); border: 1px solid var(--gray-300); border-radius: var(--radius); box-shadow: none;
.btn: background: var(--navy); color: #fff; border: none; border-radius: var(--radius); font-size: 13px; font-weight: 600; letter-spacing: 0; text-transform: none; box-shadow: none; transition: background .15s; — no hover: background: var(--navy-light);
.sourceBar: background: var(--surface); border: 1px solid var(--gray-200); border-radius: var(--radius); box-shadow: var(--shadow-sm);
.segBtn: background: var(--gray-100); border: 1px solid var(--gray-300); border-radius: 6px; color: var(--gray-700); font-size: 13px; text-transform: none; letter-spacing: 0; box-shadow: none;
.segBtn.active: background: var(--navy); color: #fff; border-color: var(--navy); box-shadow: none;
.kpis: grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
.card: background: var(--surface); border: 1px solid var(--gray-200); box-shadow: var(--shadow-sm); border-radius: var(--radius);
.kpi .label: font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: .06em; font-weight: 600;
.kpi .value: font-size: 30px; font-weight: 700; color: var(--text);
.kpi .hint: font-size: 11px; color: var(--gray-500);
Substituir filter: brightness(1.10) no estado ativo do KPI por uma classe CSS .kpi-active com border-left: 3px solid var(--yellow) !important; box-shadow: var(--shadow-md) !important;
thead th: background: var(--navy); color: rgba(255,255,255,.9); font-size: 12px; letter-spacing: .08em; text-transform: uppercase; padding: 11px 14px; cursor: pointer; user-select: none; — hover: background: var(--navy-light);
tbody tr: background: var(--surface); — hover: background: #EEF2FF;
tbody td: padding: 10px 14px; font-size: 13px; color: var(--text); border-bottom: 1px solid var(--gray-100);
tbody tr.row-multi: background: var(--gray-50) !important; — hover: var(--gray-100) !important
.badge: remover borda e .dot. Novo estilo: display: inline-flex; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; letter-spacing: .04em; text-transform: uppercase; border: none;
Substituir classes de badge por:
.b-pendente { background: var(--badge-pending-bg); color: var(--badge-pending-text); }
.b-pago { background: var(--badge-confirm-bg); color: var(--badge-confirm-text); }
.b-entregue { background: var(--badge-pronto-bg); color: var(--badge-pronto-text); }
.b-enviado { background: var(--badge-naofin-bg); color: var(--badge-naofin-text); }
.b-cancelado{ background: var(--badge-cancel-bg); color: var(--badge-cancel-text); }
.b-outro { background: var(--badge-outro-bg); color: var(--badge-outro-text); }
.tableHeader h2: font-size: 14px; font-weight: 600; color: var(--navy); letter-spacing: 0; text-transform: none;
.pill: background: var(--gray-100); border: 1px solid var(--gray-200); border-radius: 6px; color: var(--gray-700);
.pill strong: color: var(--navy);

Melhorias Funcionais
1. Toast de erro — adicionar ao <body>:

<div id="toast" class="toast" aria-live="polite"></div>

CSS:

.toast {
  position: fixed; bottom: 24px; right: 24px; z-index: 999;
  background: #C92A2A; color: #fff; padding: 12px 18px; border-radius: 8px;
  font-size: 13px; font-weight: 500; box-shadow: 0 4px 16px rgba(0,0,0,.18);
  opacity: 0; pointer-events: none; transition: opacity .25s; max-width: 360px;
}
.toast.show { opacity: 1; pointer-events: auto; }

JS — adicionar função e substituir alert():

const showToast = (msg, duration = 4500) => {
  const t = $("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove("show"), duration);
};
// Substituir .withFailureHandler((err) => alert(...)) por:
.withFailureHandler((err) => showToast("Falha ao carregar: " + (err?.message ?? err)))

2. Spinner de loading — adicionar dentro de .tableWrap antes da <div style="width:100%...">:

<div id="loadingOverlay" class="loading-overlay" style="display:none">
  <div class="spinner"></div><span>Carregando dados...</span>
</div>

CSS:

.loading-overlay { display: flex; align-items: center; justify-content: center; gap: 10px; padding: 48px 16px; color: var(--text-muted); font-size: 13px; }
.spinner { width: 20px; height: 20px; border: 2px solid var(--gray-200); border-top-color: var(--navy); border-radius: 50%; animation: spin .7s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }

JS — em fetchPedidos() adicionar no início: $("loadingOverlay").style.display = "flex"; $("tbody").innerHTML = ""; $("thead").innerHTML = ""; e nos handlers .withSuccessHandler e .withFailureHandler adicionar: $("loadingOverlay").style.display = "none";

3. Coluna Fonte (quando TODOS) — em setThead(), nas views que não são FALTAM_VIEW, adicionar condicionalmente <th data-sort="sheet" style="width:110px;">Fonte</th> quando ACTIVE_SOURCE === "TODOS". Em rowHtml(), adicionar célula <td><span class="badge b-outro" style="font-size:10px;">${safe(p.sheet ?? "—")}</span></td> quando ACTIVE_SOURCE === "TODOS" e view não for FALTAM_VIEW. Atualizar emptyColspan() para somar +1 nesse caso.

4. Ordenação por coluna — adicionar estado global let SORT = { col: null, dir: 1 };. Em cada <th> adicionar data-sort="nomeDoCampo" e <span class="sort-icon">↕</span> (quando ativo: ↑ ou ↓). CSS: .sort-icon { opacity: .4; font-size: 10px; margin-left: 4px; } / th.sort-asc .sort-icon, th.sort-desc .sort-icon { opacity: 1; }. Adicionar listener nos th via delegação no thead (usar { once: true } e re-registrar após cada renderTable). Em applyAllFilters(), antes de renderTable(list), aplicar sort: if (SORT.col) list.sort((a,b) => { const av = String(a[SORT.col]??"").toLowerCase(); const bv = String(b[SORT.col]??"").toLowerCase(); return av < bv ? -SORT.dir : av > bv ? SORT.dir : 0; });. Resetar SORT = { col: null, dir: 1 } ao trocar de KPI ou de fonte.


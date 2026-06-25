<!-- #include file="conexao.asp" -->

<%
If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

VAR_USUARIO_NOME     = "" & Session("name")
VAR_USUARIO_FUNCAO   = "" & Session("funcao")
VAR_USUARIO_NIVEL    = "" & Session("nivel")
VAR_USUARIO_MATRICULA = "" & Session("matricula")

If VAR_USUARIO_FUNCAO <> "INSPETOR" And VAR_USUARIO_FUNCAO <> "SUPERVISOR" And VAR_USUARIO_NIVEL <> "ADM" Then
    Response.Redirect("form_home.asp")
    Response.End
End If

Function HtmlSafe(valor)
    HtmlSafe = Server.HTMLEncode("" & valor)
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações – Dashboard de Leitura</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <script src="libs/js/jquery.js"></script>
    <!-- Chart.js CDN (será substituído pela lib local se disponível) -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>

    <style>
        /* =====================================================
           VARIÁVEIS – mesmo design token do sistema
        ===================================================== */
        :root {
            --vale-teal:          #00857A;
            --vale-teal-dark:     #006B63;
            --vale-yellow:        #F6B800;
            --text-dark:          #263238;
            --text-muted:         #607D8B;
            --card-bg:            rgba(255,255,255,0.92);
            --sidebar-bg:         rgba(255,255,255,0.86);
            --sidebar-collapsed:  88px;
            --sidebar-expanded:   292px;
            --danger:             #C62828;
            --success:            #2E7D32;
        }

        * { box-sizing: border-box; }

        html, body {
            width: 100%;
            min-height: 100vh;
            min-height: 100dvh;
            margin: 0;
            overflow-x: hidden;
        }

        body {
            font-family: Arial, Helvetica, sans-serif;
            background:
                linear-gradient(rgba(0,60,55,0.16), rgba(0,60,55,0.16)),
                url("libs/img/login_bg.jpg") center center / cover no-repeat fixed;
            color: var(--text-dark);
        }

        /* =====================================================
           SIDEBAR – idêntico ao restante do sistema
        ===================================================== */
        .sidebar {
            position: fixed; left: 0; top: 0;
            width: var(--sidebar-collapsed);
            height: 100vh; height: 100dvh;
            background: var(--sidebar-bg);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            box-shadow: 5px 0 24px rgba(0,0,0,0.18);
            z-index: 100;
            transition: width 0.25s ease-in-out;
            display: flex; flex-direction: column;
            overflow: hidden;
            padding: 14px 12px;
        }
        .sidebar.open { width: var(--sidebar-expanded); }

        .sidebar-header {
            flex-shrink: 0;
            display: flex; flex-direction: column;
            align-items: center; gap: 14px;
            padding-bottom: 14px;
            border-bottom: 1px solid rgba(0,133,122,0.12);
        }

        .menu-toggle {
            width: 56px; height: 56px;
            border: none; border-radius: 18px;
            background: var(--vale-teal); color: #fff;
            font-size: 26px; font-weight: 800;
            cursor: pointer; transition: all 0.2s ease-in-out;
        }
        .menu-toggle:hover { background: var(--vale-teal-dark); transform: translateY(-1px); }

        .sidebar-logo-area {
            width: 100%; display: flex;
            justify-content: center; align-items: center;
            min-height: 54px;
        }
        .sidebar-logo-area img {
            width: 58px; max-width: 58px; height: auto;
            transition: all 0.2s ease-in-out;
        }
        .sidebar.open .sidebar-logo-area img { width: 118px; max-width: 118px; }

        .sidebar-nav {
            flex: 1 1 auto;
            overflow-y: auto; overflow-x: hidden;
            display: flex; flex-direction: column;
            gap: 11px;
            padding-top: 14px; padding-bottom: 14px;
        }

        .sidebar-footer {
            flex-shrink: 0;
            padding-top: 12px;
            border-top: 1px solid rgba(0,133,122,0.12);
        }

        .sidebar-link, .external-toggle {
            width: 100%; min-height: 58px;
            border: none; border-radius: 18px;
            background: rgba(255,255,255,0.78);
            color: var(--text-dark); text-decoration: none;
            display: flex; align-items: center; gap: 14px;
            padding: 0 12px; cursor: pointer;
            box-shadow: 0 7px 16px rgba(0,0,0,0.10);
            transition: all 0.2s ease-in-out;
            font-family: Arial, Helvetica, sans-serif;
        }
        .sidebar-link:hover, .external-toggle:hover {
            background: #fff;
            transform: translateX(4px);
            box-shadow: 0 12px 25px rgba(0,0,0,0.18);
        }
        .sidebar-link.active {
            background: var(--vale-teal);
            color: #fff;
        }
        .sidebar-link.active .sidebar-title { color: #fff; }
        .sidebar:not(.open) .sidebar-link,
        .sidebar:not(.open) .external-toggle {
            justify-content: center;
            padding-left: 0; padding-right: 0;
        }

        .sidebar-icon {
            width: 38px; height: 38px;
            border-radius: 14px;
            background: var(--vale-teal); color: #fff;
            display: flex; align-items: center; justify-content: center;
            flex-shrink: 0; overflow: hidden;
            font-weight: 800; font-size: 17px;
        }
        .sidebar-icon img { width: 100%; height: 100%; object-fit: cover; }

        .sidebar-text {
            display: none; flex-direction: column;
            text-align: left; line-height: 1.15;
            white-space: nowrap; min-width: 0;
        }
        .sidebar.open .sidebar-text { display: flex; }

        .sidebar-title { font-size: 13px; font-weight: 800; color: var(--vale-teal-dark); }
        .sidebar-subtitle { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
        .sidebar-footer .sidebar-icon { background: var(--danger); }

        .external-group { width: 100%; }
        .external-arrow {
            display: none; margin-left: auto;
            color: var(--vale-teal-dark); font-size: 15px;
            transition: transform 0.2s ease-in-out;
        }
        .sidebar.open .external-arrow { display: inline-block; }
        .external-group.open .external-arrow { transform: rotate(180deg); }
        .external-list {
            display: none; flex-direction: column;
            gap: 8px; margin-top: 8px;
            padding-left: 4px; padding-right: 4px;
        }
        .sidebar.open .external-group.open .external-list { display: flex; }
        .sidebar:not(.open) .external-list { display: none !important; }
        .external-link {
            min-height: 48px; border-radius: 15px;
            background: rgba(255,255,255,0.68);
            color: var(--text-dark); text-decoration: none;
            display: flex; align-items: center; gap: 11px;
            padding: 0 12px; transition: all 0.2s ease-in-out;
        }
        .external-link:hover { background: #fff; transform: translateX(4px); }
        .external-link img, .external-logo {
            width: 31px; height: 31px;
            border-radius: 11px; object-fit: cover; flex-shrink: 0;
        }
        .external-link span { font-size: 13px; font-weight: 800; color: var(--vale-teal-dark); }

        .external-icon-stack-folder {
            width: 42px !important; height: 42px !important;
            min-width: 42px !important; min-height: 42px !important;
            max-width: 42px !important; max-height: 42px !important;
            position: relative !important; display: block !important;
            background: transparent !important; border-radius: 14px !important;
            overflow: visible !important; padding: 0 !important; flex-shrink: 0 !important;
        }
        .external-icon-stack-folder .folder-card {
            position: absolute !important;
            width: 29px !important; height: 29px !important;
            border-radius: 10px !important; background: #fff !important;
            display: flex !important; align-items: center !important;
            justify-content: center !important;
            box-shadow: 0 3px 8px rgba(0,0,0,0.22) !important;
            border: 2px solid rgba(255,255,255,0.95) !important;
            overflow: hidden !important;
        }
        .external-icon-stack-folder .folder-card img {
            width: 100% !important; height: 100% !important;
            object-fit: cover !important; border: none !important;
            border-radius: 8px !important;
        }
        .external-icon-stack-folder .folder-card-back  { left:1px; top:5px;  z-index:1; opacity:.92; transform:rotate(-8deg) scale(.92); }
        .external-icon-stack-folder .folder-card-middle{ left:7px; top:2px;  z-index:2; opacity:.96; transform:rotate(5deg) scale(.96); }
        .external-icon-stack-folder .folder-card-front { left:11px; top:10px; z-index:3; transform:rotate(0deg) scale(1); }

        /* =====================================================
           CONTEÚDO PRINCIPAL
        ===================================================== */
        .main-content {
            min-height: 100vh;
            margin-left: var(--sidebar-collapsed);
            width: calc(100vw - var(--sidebar-collapsed));
            padding: clamp(20px, 2.8vw, 40px);
            transition: margin-left 0.25s ease-in-out, width 0.25s ease-in-out;
        }
        body.sidebar-open .main-content {
            margin-left: var(--sidebar-expanded);
            width: calc(100vw - var(--sidebar-expanded));
        }

        .page-header {
            width: min(100%, 1400px);
            margin: 0 auto 22px auto;
            display: flex; flex-direction: column;
            align-items: center; text-align: center; gap: 10px;
        }
        .page-title {
            font-size: clamp(20px, 2.5vw, 32px);
            font-weight: 800; color: #fff;
            text-shadow: 0 2px 10px rgba(0,0,0,0.35);
            margin: 0;
        }
        .user-info { font-size: 14px; color: rgba(255,255,255,0.88); }

        /* =====================================================
           FILTROS
        ===================================================== */
        .filters-card {
            width: min(100%, 1400px);
            margin: 0 auto 22px auto;
            background: var(--card-bg);
            border-radius: 22px;
            padding: 18px 22px;
            box-shadow: 0 8px 28px rgba(0,0,0,0.13);
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            align-items: flex-end;
        }
        .filters-card label {
            font-size: 12px; font-weight: 700;
            color: var(--text-muted);
            display: flex; flex-direction: column; gap: 4px;
            min-width: 140px;
        }
        .filters-card input,
        .filters-card select {
            padding: 8px 12px;
            border: 1.5px solid rgba(0,133,122,0.25);
            border-radius: 10px;
            font-size: 13px;
            color: var(--text-dark);
            background: #fff;
            outline: none;
            transition: border-color 0.2s;
        }
        .filters-card input:focus,
        .filters-card select:focus { border-color: var(--vale-teal); }
        .btn-filtrar, .btn-limpar, .btn-exportar {
            padding: 10px 20px;
            border: none; border-radius: 12px;
            font-size: 13px; font-weight: 700;
            cursor: pointer; transition: all 0.2s;
        }
        .btn-filtrar  { background: var(--vale-teal); color: #fff; }
        .btn-filtrar:hover { background: var(--vale-teal-dark); transform: translateY(-1px); }
        .btn-limpar   { background: rgba(0,133,122,0.08); color: var(--vale-teal-dark); }
        .btn-limpar:hover { background: rgba(0,133,122,0.16); }
        .btn-exportar { background: var(--vale-yellow); color: var(--text-dark); }
        .btn-exportar:hover { opacity: 0.88; transform: translateY(-1px); }

        /* =====================================================
           KPI CARDS
        ===================================================== */
        .kpi-grid {
            width: min(100%, 1400px);
            margin: 0 auto 24px auto;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 16px;
        }
        .kpi-card {
            background: var(--card-bg);
            border-radius: 22px;
            padding: 20px 18px;
            box-shadow: 0 8px 28px rgba(0,0,0,0.13);
            text-align: center;
            position: relative;
            overflow: hidden;
        }
        .kpi-card::before {
            content: "";
            position: absolute; top: 0; left: 0; right: 0;
            height: 4px;
            background: var(--vale-teal);
            border-radius: 22px 22px 0 0;
        }
        .kpi-card.kpi-danger::before { background: var(--danger); }
        .kpi-card.kpi-success::before { background: var(--success); }
        .kpi-card.kpi-yellow::before { background: var(--vale-yellow); }

        .kpi-value {
            font-size: 2.2rem; font-weight: 800;
            color: var(--vale-teal-dark); line-height: 1;
            margin-bottom: 6px;
        }
        .kpi-card.kpi-danger .kpi-value  { color: var(--danger); }
        .kpi-card.kpi-success .kpi-value { color: var(--success); }
        .kpi-card.kpi-yellow .kpi-value  { color: #9a7200; }

        .kpi-label {
            font-size: 12px; font-weight: 700;
            color: var(--text-muted); text-transform: uppercase;
            letter-spacing: 0.04em;
        }

        /* =====================================================
           GRID DE GRÁFICOS
        ===================================================== */
        .charts-grid {
            width: min(100%, 1400px);
            margin: 0 auto 26px auto;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
        }
        .chart-card {
            background: var(--card-bg);
            border-radius: 22px;
            padding: 22px 20px;
            box-shadow: 0 8px 28px rgba(0,0,0,0.13);
        }
        .chart-card.chart-wide {
            grid-column: 1 / -1;
        }
        .chart-card-title {
            font-size: 14px; font-weight: 800;
            color: var(--vale-teal-dark);
            margin: 0 0 16px 0;
            display: flex; align-items: center; gap: 8px;
        }
        .chart-card-title::before {
            content: "";
            display: inline-block;
            width: 4px; height: 18px;
            background: var(--vale-teal);
            border-radius: 4px;
        }
        .chart-wrap { position: relative; height: 260px; }
        .chart-wrap canvas { max-height: 260px; }

        /* =====================================================
           TABELA HISTÓRICO
        ===================================================== */
        .section-historico {
            width: min(100%, 1400px);
            margin: 0 auto 32px auto;
        }
        .section-title {
            font-size: 18px; font-weight: 800;
            color: #fff;
            text-shadow: 0 2px 8px rgba(0,0,0,0.3);
            margin: 0 0 14px 0;
            display: flex; align-items: center; gap: 10px;
        }

        .table-card {
            background: var(--card-bg);
            border-radius: 22px;
            padding: 0;
            box-shadow: 0 8px 28px rgba(0,0,0,0.13);
            overflow: hidden;
        }
        .table-controls {
            padding: 16px 20px;
            display: flex; justify-content: space-between; align-items: center;
            border-bottom: 1px solid rgba(0,133,122,0.10);
            flex-wrap: wrap; gap: 10px;
        }
        .table-info { font-size: 13px; color: var(--text-muted); }

        .hist-table {
            width: 100%;
            border-collapse: collapse;
        }
        .hist-table thead tr {
            background: linear-gradient(135deg, var(--vale-teal), var(--vale-teal-dark));
            color: #fff;
        }
        .hist-table th {
            padding: 13px 16px;
            font-size: 12px; font-weight: 700;
            text-align: left; white-space: nowrap;
            letter-spacing: 0.04em;
        }
        .hist-table tbody tr {
            border-bottom: 1px solid rgba(0,133,122,0.07);
            transition: background 0.15s;
        }
        .hist-table tbody tr:hover { background: rgba(0,133,122,0.05); }
        .hist-table td {
            padding: 11px 16px;
            font-size: 13px; color: var(--text-dark);
            vertical-align: middle;
        }
        .hist-cod {
            font-size: 11px; font-weight: 800;
            color: var(--vale-teal-dark);
            background: rgba(0,133,122,0.10);
            padding: 3px 7px; border-radius: 7px;
        }
        .badge-sim {
            display: inline-block;
            background: rgba(46,125,50,0.12);
            color: var(--success);
            font-size: 11px; font-weight: 700;
            padding: 3px 8px; border-radius: 8px;
        }
        .badge-nao {
            display: inline-block;
            background: rgba(198,40,40,0.10);
            color: var(--danger);
            font-size: 11px; font-weight: 700;
            padding: 3px 8px; border-radius: 8px;
        }
        .btn-detalhe {
            padding: 6px 14px;
            border: 1.5px solid var(--vale-teal);
            background: transparent; color: var(--vale-teal-dark);
            border-radius: 10px; font-size: 12px; font-weight: 700;
            cursor: pointer; transition: all 0.18s;
        }
        .btn-detalhe:hover { background: var(--vale-teal); color: #fff; }

        /* Paginação */
        .pagination {
            display: flex; justify-content: center;
            gap: 6px; padding: 16px;
        }
        .page-btn {
            min-width: 36px; height: 36px;
            border: 1.5px solid rgba(0,133,122,0.25);
            background: #fff; color: var(--vale-teal-dark);
            border-radius: 10px; font-size: 13px; font-weight: 700;
            cursor: pointer; transition: all 0.18s;
        }
        .page-btn:hover, .page-btn.active {
            background: var(--vale-teal); color: #fff;
            border-color: var(--vale-teal);
        }
        .page-btn:disabled { opacity: 0.4; cursor: default; }

        /* =====================================================
           MODAL DETALHE REGISTRO
        ===================================================== */
        .modal-overlay {
            display: none;
            position: fixed; inset: 0;
            background: rgba(0,0,0,0.52);
            backdrop-filter: blur(4px);
            z-index: 200;
            align-items: center; justify-content: center;
        }
        .modal-overlay.open { display: flex; }
        .modal-box {
            background: #fff;
            border-radius: 24px;
            width: min(94vw, 780px);
            max-height: 88vh;
            overflow-y: auto;
            box-shadow: 0 24px 60px rgba(0,0,0,0.32);
            padding: 0;
        }
        .modal-header {
            background: linear-gradient(135deg, var(--vale-teal), var(--vale-teal-dark));
            color: #fff;
            padding: 18px 22px;
            border-radius: 24px 24px 0 0;
            display: flex; justify-content: space-between; align-items: center;
        }
        .modal-title { font-size: 15px; font-weight: 800; }
        .modal-close {
            background: rgba(255,255,255,0.18);
            border: none; color: #fff;
            width: 34px; height: 34px;
            border-radius: 10px; font-size: 18px;
            cursor: pointer; transition: background 0.15s;
        }
        .modal-close:hover { background: rgba(255,255,255,0.32); }
        .modal-body { padding: 22px; }
        .modal-chart-wrap { position: relative; height: 220px; margin-bottom: 20px; }
        .modal-chart-wrap canvas { max-height: 220px; }

        .modal-sub {
            font-size: 13px; font-weight: 800;
            color: var(--vale-teal-dark);
            margin: 18px 0 10px 0;
            padding-bottom: 6px;
            border-bottom: 2px solid rgba(0,133,122,0.12);
        }
        .modal-table {
            width: 100%; border-collapse: collapse;
            font-size: 12px;
        }
        .modal-table th {
            background: rgba(0,133,122,0.08);
            padding: 8px 10px;
            text-align: left; font-weight: 700;
            color: var(--text-muted);
        }
        .modal-table td { padding: 8px 10px; border-bottom: 1px solid rgba(0,0,0,0.05); }
        .modal-table tr:last-child td { border-bottom: none; }

        .modal-footer { padding: 14px 22px; text-align: right; border-top: 1px solid rgba(0,0,0,0.07); }
        .btn-fechar {
            padding: 10px 24px;
            background: var(--vale-teal); color: #fff;
            border: none; border-radius: 12px;
            font-size: 14px; font-weight: 700;
            cursor: pointer; transition: background 0.2s;
        }
        .btn-fechar:hover { background: var(--vale-teal-dark); }

        /* =====================================================
           LOADING SPINNER
        ===================================================== */
        .spinner {
            display: inline-block;
            width: 36px; height: 36px;
            border: 4px solid rgba(0,133,122,0.2);
            border-top-color: var(--vale-teal);
            border-radius: 50%;
            animation: spin 0.75s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .loading-center {
            display: flex; justify-content: center;
            align-items: center; padding: 40px;
        }

        /* =====================================================
           RESPONSIVO
        ===================================================== */
        @media (max-width: 700px) {
            .charts-grid { grid-template-columns: 1fr; }
            .kpi-grid    { grid-template-columns: repeat(2, 1fr); }
            .hist-table th:nth-child(4),
            .hist-table td:nth-child(4) { display: none; }
        }
    </style>
</head>

<body>

    <!-- =====================================================
         SIDEBAR
    ===================================================== -->
    <aside id="sidebar" class="sidebar">

        <div class="sidebar-header">
            <button type="button" class="menu-toggle" onclick="toggleSidebar()" title="Abrir ou fechar menu">☰</button>
            <div class="sidebar-logo-area">
                <img src="libs/img/logo-vale.png" alt="Logo Vale"
                     onerror="this.onerror=null;this.src='libs/img/Logotipo_Vale.png';">
            </div>
        </div>

        <nav class="sidebar-nav">

            <a href="form_home.asp" class="sidebar-link">
                <span class="sidebar-icon"><img src="libs/img/img_link/Home.jpg" alt="Home"></span>
                <span class="sidebar-text"><span class="sidebar-title">Home</span></span>
            </a>

            <a href="form_visualizar_registro.asp" class="sidebar-link">
                <span class="sidebar-icon"><img src="libs/img/img_link/visualizar.jpg" alt="Visualizar"></span>
                <span class="sidebar-text"><span class="sidebar-title">Visualizar registro</span></span>
            </a>

            <% If VAR_USUARIO_FUNCAO = "SUPERVISOR" Or VAR_USUARIO_NIVEL = "ADM" Then %>
            <a href="form_formulario.asp" class="sidebar-link">
                <span class="sidebar-icon">+</span>
                <span class="sidebar-text"><span class="sidebar-title">Novo registro</span></span>
            </a>
            <% End If %>

            <a href="form_grafico.asp" class="sidebar-link">
                <span class="sidebar-icon"><img src="libs/img/img_link/relatorios.jpg" alt="Relatórios"></span>
                <span class="sidebar-text"><span class="sidebar-title">Relatórios</span></span>
            </a>

            <a href="form_relatorios.asp" class="sidebar-link active">
                <span class="sidebar-icon">📊</span>
                <span class="sidebar-text">
                    <span class="sidebar-title">Dashboard</span>
                    <span class="sidebar-subtitle">Leitura e métricas</span>
                </span>
            </a>

            <a href="https://efvmworkplace/central%20de%20kaizen/index.php" class="sidebar-link" target="_blank">
                <span class="sidebar-icon"><img src="libs/img/img_link/kaizen.jpg" alt="Kaizen"></span>
                <span class="sidebar-text"><span class="sidebar-title">Kaizen</span></span>
            </a>

            <a href="https://linktr.ee/FaleComVale" class="sidebar-link" target="_blank">
                <span class="sidebar-icon">#</span>
                <span class="sidebar-text"><span class="sidebar-title">#FaleCOM</span></span>
            </a>

            <a href="dss_online.asp" class="sidebar-link">
                <span class="sidebar-icon">D</span>
                <span class="sidebar-text"><span class="sidebar-title">DSS Online</span></span>
            </a>

            <div id="linksUteisGroup" class="external-group">
                <button type="button" class="external-toggle" onclick="toggleLinksUteis()">
                    <span class="sidebar-icon external-icon-stack-folder">
                        <span class="folder-card folder-card-back"><img src="libs/img/img_link/gdb.jpg" alt="GDB"></span>
                        <span class="folder-card folder-card-middle"><img src="libs/img/img_link/edados.jpg" alt="E-Dados"></span>
                        <span class="folder-card folder-card-front"><img src="libs/img/img_link/irislogo.png" alt="Iris"></span>
                    </span>
                    <span class="sidebar-text">
                        <span class="sidebar-title">Links úteis</span>
                        <span class="sidebar-subtitle">Sistemas externos</span>
                    </span>
                    <span class="external-arrow">▾</span>
                </button>
                <div class="external-list">
                    <a href="https://equipfer.valenet.valeglobal.net/equipfer/" class="external-link" target="_blank">
                        <img src="libs/img/img_link/tooplate_image_01.jpg" alt="Equipfer"><span>Equipfer</span>
                    </a>
                    <a href="https://gdb.valeglobal.net/gdb/view/login/login.faces" class="external-link" target="_blank">
                        <img src="libs/img/img_link/gdb.jpg" alt="GDB"><span>GDB</span>
                    </a>
                    <a href="https://performancemanager4.successfactors.com/sf/home" class="external-link" target="_blank">
                        <img src="libs/img/img_link/prontidao1.jpg" alt="VES"><span>VES</span>
                    </a>
                    <a href="https://lbrportalfolha.valenet.valeglobal.net/portalrh/Produtos/SAAA/Principal2.aspx?amb_selecionado=0&abrir_nova_janela=N&eh_mdesigner=N&nome_portal=616653596455764672655738516965596E57664E38413D3D" class="external-link" target="_blank">
                        <img src="libs/img/img_link/edados.jpg" alt="E-Dados"><span>E-Dados</span>
                    </a>
                    <a href="https://iris.valeglobal.net/login" class="external-link" target="_blank">
                        <img src="libs/img/img_link/irislogo.png" alt="Iris"><span>Iris</span>
                    </a>
                </div>
            </div>

        </nav>

        <div class="sidebar-footer">
            <a class="sidebar-link sair" href="fecha_session.asp">
                <span class="sidebar-icon">S</span>
                <span class="sidebar-text">
                    <span class="sidebar-title">Sair</span>
                    <span class="sidebar-subtitle">Encerrar sessão</span>
                </span>
            </a>
        </div>

    </aside>

    <!-- =====================================================
         CONTEÚDO PRINCIPAL
    ===================================================== -->
    <main class="main-content">

        <header class="page-header">
            <h1 class="page-title">📊 Dashboard de Leitura</h1>
            <div class="user-info">
                Bem-vindo(a), <strong><%=HtmlSafe(VAR_USUARIO_NOME)%></strong>
                &nbsp;·&nbsp; <%=HtmlSafe(VAR_USUARIO_FUNCAO)%>
            </div>
        </header>

        <!-- FILTROS -->
        <div class="filters-card">
            <label>
                Data inicial
                <input type="text" id="f_data_ini" placeholder="DD/MM/AAAA" maxlength="10">
            </label>
            <label>
                Data final
                <input type="text" id="f_data_fim" placeholder="DD/MM/AAAA" maxlength="10">
            </label>
            <label>
                Tipo
                <input type="text" id="f_tipo" placeholder="ex: AVISO">
            </label>
            <label>
                Código
                <input type="number" id="f_cod" placeholder="ex: 42" min="1">
            </label>
            <button class="btn-filtrar" onclick="aplicarFiltros()">Aplicar filtros</button>
            <button class="btn-limpar"  onclick="limparFiltros()">Limpar</button>
            <button class="btn-exportar" onclick="exportarCSV()">⬇ Exportar CSV</button>
        </div>

        <!-- KPI CARDS -->
        <div class="kpi-grid" id="kpiGrid">
            <div class="loading-center"><div class="spinner"></div></div>
        </div>

        <!-- GRÁFICOS -->
        <div class="charts-grid">

            <div class="chart-card chart-wide">
                <p class="chart-card-title">Leituras confirmadas por dia</p>
                <div class="chart-wrap"><canvas id="chartLeiturasdia"></canvas></div>
            </div>

            <div class="chart-card">
                <p class="chart-card-title">Registros por tipo</p>
                <div class="chart-wrap"><canvas id="chartPorTipo"></canvas></div>
            </div>

            <div class="chart-card">
                <p class="chart-card-title">Top 10 usuários com mais pendências</p>
                <div class="chart-wrap"><canvas id="chartPendencias"></canvas></div>
            </div>

            <div class="chart-card chart-wide">
                <p class="chart-card-title">Top 10 registros com maior tempo médio de leitura (dias)</p>
                <div class="chart-wrap" style="height:320px;"><canvas id="chartTempoMedio"></canvas></div>
            </div>

        </div>

        <!-- TABELA HISTÓRICO -->
        <section class="section-historico">
            <h2 class="section-title">📋 Histórico de Informações</h2>
            <div class="table-card">
                <div class="table-controls">
                    <div class="table-info" id="tableInfo">Carregando...</div>
                    <div id="paginationTop" class="pagination" style="padding:0;"></div>
                </div>
                <div id="tableWrap">
                    <div class="loading-center"><div class="spinner"></div></div>
                </div>
                <div id="paginationBottom" class="pagination"></div>
            </div>
        </section>

    </main>

    <!-- =====================================================
         MODAL DETALHE
    ===================================================== -->
    <div id="modalDetalhe" class="modal-overlay">
        <div class="modal-box">
            <div class="modal-header">
                <div class="modal-title" id="modalDetalheTitle">Detalhe do registro</div>
                <button type="button" class="modal-close" onclick="fecharModal()">✕</button>
            </div>
            <div class="modal-body" id="modalDetalheBody">
                <div class="loading-center"><div class="spinner"></div></div>
            </div>
            <div class="modal-footer">
                <button class="btn-fechar" onclick="fecharModal()">Fechar</button>
            </div>
        </div>
    </div>

    <!-- =====================================================
         SCRIPTS
    ===================================================== -->
    <script>

    /* -------------------------------------------------------
       SIDEBAR
    ------------------------------------------------------- */
    function toggleSidebar() {
        var s = document.getElementById("sidebar");
        var b = document.body;
        if (s.classList.contains("open")) {
            s.classList.remove("open");
            b.classList.remove("sidebar-open");
            localStorage.setItem("centralSidebarOpen","N");
        } else {
            s.classList.add("open");
            b.classList.add("sidebar-open");
            localStorage.setItem("centralSidebarOpen","S");
        }
    }

    function toggleLinksUteis() {
        var s = document.getElementById("sidebar");
        var g = document.getElementById("linksUteisGroup");
        if (!s.classList.contains("open")) {
            s.classList.add("open");
            document.body.classList.add("sidebar-open");
            localStorage.setItem("centralSidebarOpen","S");
        }
        g.classList.toggle("open");
        localStorage.setItem("centralLinksUteisOpen", g.classList.contains("open") ? "S" : "N");
    }

    window.addEventListener("DOMContentLoaded", function() {
        if (localStorage.getItem("centralSidebarOpen") === "S") {
            document.getElementById("sidebar").classList.add("open");
            document.body.classList.add("sidebar-open");
        }
        if (localStorage.getItem("centralLinksUteisOpen") === "S") {
            document.getElementById("linksUteisGroup").classList.add("open");
        }
        inicializar();
    });

    /* -------------------------------------------------------
       ESTADO GLOBAL
    ------------------------------------------------------- */
    var FILTROS = { data_ini:"", data_fim:"", tipo:"", cod:"" };
    var PG_ATUAL = 1;
    var HISTORICO_CACHE = null;    // cache para CSV

    var chartInstancias = {};
    var CORES_TEAL = [
        "#00857A","#00A898","#006B63","#33ADA4","#005C55",
        "#44C2B8","#007A71","#00C4B8","#004F4A","#66D4CC"
    ];

    /* -------------------------------------------------------
       FILTROS
    ------------------------------------------------------- */
    function aplicarFiltros() {
        FILTROS.data_ini = document.getElementById("f_data_ini").value.trim();
        FILTROS.data_fim = document.getElementById("f_data_fim").value.trim();
        FILTROS.tipo     = document.getElementById("f_tipo").value.trim();
        FILTROS.cod      = document.getElementById("f_cod").value.trim();
        PG_ATUAL = 1;
        carregarTudo();
    }

    function limparFiltros() {
        ["f_data_ini","f_data_fim","f_tipo","f_cod"].forEach(function(id){
            document.getElementById(id).value = "";
        });
        FILTROS = { data_ini:"", data_fim:"", tipo:"", cod:"" };
        PG_ATUAL = 1;
        carregarTudo();
    }

    function buildQS(extra) {
        var p = [];
        if (FILTROS.data_ini) p.push("data_ini=" + encodeURIComponent(FILTROS.data_ini));
        if (FILTROS.data_fim) p.push("data_fim=" + encodeURIComponent(FILTROS.data_fim));
        if (FILTROS.tipo)     p.push("tipo="     + encodeURIComponent(FILTROS.tipo));
        if (FILTROS.cod)      p.push("cod="      + encodeURIComponent(FILTROS.cod));
        if (extra) p.push(extra);
        return p.length ? "&" + p.join("&") : "";
    }

    /* -------------------------------------------------------
       FETCH HELPER
    ------------------------------------------------------- */
    function apiGet(acao, extra, cb) {
        var url = "api_relatorios.asp?acao=" + acao + buildQS(extra) + "&_t=" + Date.now();
        fetch(url)
            .then(function(r){ return r.json(); })
            .then(cb)
            .catch(function(e){ console.error("Erro API [" + acao + "]:", e); });
    }

    /* -------------------------------------------------------
       INICIALIZAR TUDO
    ------------------------------------------------------- */
    function inicializar() {
        carregarTudo();
    }

    function carregarTudo() {
        carregarKPIs();
        carregarLeiturasdia();
        carregarPorTipo();
        carregarPendencias();
        carregarTempoMedio();
        carregarHistorico(PG_ATUAL);
    }

    /* -------------------------------------------------------
       KPIs
    ------------------------------------------------------- */
    function carregarKPIs() {
        var g = document.getElementById("kpiGrid");
        g.innerHTML = '<div class="loading-center"><div class="spinner"></div></div>';
        apiGet("resumo", "", function(d) {
            if (d.erro) { g.innerHTML = "<p style='color:red'>"+d.erro+"</p>"; return; }
            g.innerHTML =
                kpiCard("Total de Registros", d.total_registros, "", "kpi-default") +
                kpiCard("Leituras Confirmadas", d.total_leituras_sim, "", "kpi-success") +
                kpiCard("Pendências Abertas", d.total_pendencias, "", "kpi-danger") +
                kpiCard("Usuários Alcançados", d.total_usuarios, "", "kpi-yellow") +
                kpiCard("Taxa de Leitura", d.taxa_leitura_pct + "%", "", "kpi-default");
        });
    }

    function kpiCard(label, valor, sub, cls) {
        return '<div class="kpi-card ' + cls + '">' +
               '<div class="kpi-value">' + valor + '</div>' +
               '<div class="kpi-label">' + label + '</div>' +
               '</div>';
    }

    /* -------------------------------------------------------
       GRÁFICO: LEITURAS POR DIA (linha)
    ------------------------------------------------------- */
    function carregarLeiturasdia() {
        apiGet("leituras_dia", "", function(d) {
            if (d.erro || !d.labels) return;
            // Inverter para ordem cronológica
            var lbs  = d.labels.reverse();
            var vals = d.valores.reverse();
            renderChart("chartLeiturasdia", "line", lbs, vals, "Leituras", {
                tension: 0.4,
                fill: true,
                backgroundColor: "rgba(0,133,122,0.08)",
                borderColor: "#00857A",
                pointBackgroundColor: "#00857A"
            });
        });
    }

    /* -------------------------------------------------------
       GRÁFICO: POR TIPO (doughnut)
    ------------------------------------------------------- */
    function carregarPorTipo() {
        apiGet("por_tipo", "", function(d) {
            if (d.erro || !d.labels) return;
            renderChartDoughnut("chartPorTipo", d.labels, d.valores);
        });
    }

    /* -------------------------------------------------------
       GRÁFICO: PENDÊNCIAS POR USUÁRIO (barras)
    ------------------------------------------------------- */
    function carregarPendencias() {
        apiGet("pendencias_user", "", function(d) {
            if (d.erro || !d.labels) return;
            renderChart("chartPendencias", "bar", d.labels, d.valores, "Pendências", {
                backgroundColor: "rgba(198,40,40,0.72)",
                borderColor: "#C62828",
                borderRadius: 8,
                indexAxis: "y"
            });
        });
    }

    /* -------------------------------------------------------
       GRÁFICO: TEMPO MÉDIO (barras horizontais)
    ------------------------------------------------------- */
    function carregarTempoMedio() {
        apiGet("tempo_medio", "", function(d) {
            if (d.erro || !d.labels) return;
            renderChart("chartTempoMedio", "bar", d.labels, d.valores, "Dias (média)", {
                backgroundColor: "rgba(0,133,122,0.72)",
                borderColor: "#00857A",
                borderRadius: 8,
                indexAxis: "y"
            });
        });
    }

    /* -------------------------------------------------------
       HISTÓRICO – TABELA PAGINADA
    ------------------------------------------------------- */
    function carregarHistorico(pg) {
        PG_ATUAL = pg;
        var wrap = document.getElementById("tableWrap");
        wrap.innerHTML = '<div class="loading-center"><div class="spinner"></div></div>';
        apiGet("historico", "pg=" + pg, function(d) {
            if (d.erro) { wrap.innerHTML = "<p style='color:red'>"+d.erro+"</p>"; return; }
            HISTORICO_CACHE = d;
            document.getElementById("tableInfo").textContent =
                d.total_registros + " registro(s) encontrado(s) · Página " + d.pagina + " de " + d.total_paginas;
            renderTabela(d);
            renderPaginacao(d.pagina, d.total_paginas);
        });
    }

    function renderTabela(d) {
        var html = '<table class="hist-table">' +
            '<thead><tr>' +
            '<th>Código</th><th>Data</th><th>Título</th><th>Tipo</th><th>Autor</th>' +
            '<th>✅ Lidas</th><th>⏳ Pendentes</th><th></th>' +
            '</tr></thead><tbody>';

        if (d.itens.length === 0) {
            html += '<tr><td colspan="8" style="text-align:center;padding:28px;color:var(--text-muted)">Nenhum registro encontrado.</td></tr>';
        } else {
            d.itens.forEach(function(item) {
                var total = item.sim + item.nao;
                var taxa  = total > 0 ? Math.round((item.sim/total)*100) + "%" : "–";
                html += '<tr>' +
                    '<td><span class="hist-cod">#' + item.cod + '</span></td>' +
                    '<td>' + item.data + '</td>' +
                    '<td style="max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="'+escHtml(item.titulo)+'">' + escHtml(item.titulo) + '</td>' +
                    '<td>' + escHtml(item.tipo) + '</td>' +
                    '<td>' + escHtml(item.autor) + '</td>' +
                    '<td><span class="badge-sim">' + item.sim + ' (' + taxa + ')</span></td>' +
                    '<td><span class="badge-nao">' + item.nao + '</span></td>' +
                    '<td><button class="btn-detalhe" onclick="abrirDetalhe(' + item.cod + ')">Ver detalhe</button></td>' +
                    '</tr>';
            });
        }
        html += '</tbody></table>';
        document.getElementById("tableWrap").innerHTML = html;
    }

    function renderPaginacao(pg, total) {
        var html = "";
        if (total <= 1) { html = ""; }
        else {
            html += '<button class="page-btn" onclick="carregarHistorico(' + (pg-1) + ')" ' + (pg<=1?"disabled":"") + '>‹</button>';
            var inicio = Math.max(1, pg-3);
            var fim    = Math.min(total, pg+3);
            if (inicio > 1) html += '<button class="page-btn" onclick="carregarHistorico(1)">1</button>' + (inicio>2?'<span style="padding:0 4px">…</span>':"");
            for (var i=inicio; i<=fim; i++) {
                html += '<button class="page-btn' + (i===pg?" active":"") + '" onclick="carregarHistorico('+i+')">'+i+'</button>';
            }
            if (fim < total) html += (fim<total-1?'<span style="padding:0 4px">…</span>':"") + '<button class="page-btn" onclick="carregarHistorico('+total+')">'+total+'</button>';
            html += '<button class="page-btn" onclick="carregarHistorico(' + (pg+1) + ')" ' + (pg>=total?"disabled":"") + '>›</button>';
        }
        document.getElementById("paginationTop").innerHTML    = html;
        document.getElementById("paginationBottom").innerHTML = html;
    }

    /* -------------------------------------------------------
       MODAL DETALHE
    ------------------------------------------------------- */
    function abrirDetalhe(cod) {
        document.getElementById("modalDetalheTitle").textContent = "Registro #" + cod;
        document.getElementById("modalDetalheBody").innerHTML =
            '<div class="loading-center"><div class="spinner"></div></div>';
        document.getElementById("modalDetalhe").classList.add("open");

        apiGet("por_registro", "cod=" + cod, function(d) {
            if (d.erro) {
                document.getElementById("modalDetalheBody").innerHTML =
                    "<p style='color:red'>" + d.erro + "</p>";
                return;
            }
            document.getElementById("modalDetalheTitle").textContent =
                "Registro #" + d.cod + " – " + d.titulo;
            renderModalDetalhe(d);
        });
    }

    function renderModalDetalhe(d) {
        var body = document.getElementById("modalDetalheBody");
        var total = d.sim + d.nao;
        var taxaPct = total > 0 ? Math.round((d.sim/total)*100) : 0;

        var html =
            '<p style="font-size:13px;color:var(--text-muted);margin:0 0 14px 0">' +
            '<strong>Tipo:</strong> ' + escHtml(d.tipo) +
            ' &nbsp;·&nbsp; <strong>Data:</strong> ' + escHtml(d.data) + '</p>' +

            '<div class="modal-chart-wrap"><canvas id="chartModalPie"></canvas></div>' +

            '<p class="modal-sub">✅ Confirmaram leitura (' + d.sim + ')</p>' +
            tabelaModal(["Matrícula","Nome","Data Leitura","Dias Corridos"], d.leitores, function(item){
                return '<td>' + escHtml(item.matricula) + '</td>' +
                       '<td>' + escHtml(item.nome) + '</td>' +
                       '<td>' + escHtml(item.data_visualizacao) + '</td>' +
                       '<td>' + item.dias_corridos + '</td>';
            }) +

            '<p class="modal-sub">⏳ Pendentes (' + d.nao + ')</p>' +
            tabelaModal(["Matrícula","Nome","Dias Pendente"], d.pendentes, function(item){
                return '<td>' + escHtml(item.matricula) + '</td>' +
                       '<td>' + escHtml(item.nome) + '</td>' +
                       '<td style="color:var(--danger);font-weight:700">' + item.dias_pendente + '</td>';
            });

        body.innerHTML = html;

        // Renderiza o pie no modal
        if (chartInstancias["chartModalPie"]) { chartInstancias["chartModalPie"].destroy(); }
        var ctx = document.getElementById("chartModalPie").getContext("2d");
        chartInstancias["chartModalPie"] = new Chart(ctx, {
            type: "doughnut",
            data: {
                labels: ["Leram (" + taxaPct + "%)", "Pendentes (" + (100-taxaPct) + "%)"],
                datasets: [{
                    data: [d.sim, d.nao],
                    backgroundColor: ["#00857A","#C62828"],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: true,
                plugins: {
                    legend: { position: "bottom", labels: { font: { size: 12 }, boxWidth: 14 } },
                    tooltip: { callbacks: { label: function(ctx){
                        return " " + ctx.label + ": " + ctx.parsed;
                    }}}
                }
            }
        });
    }

    function tabelaModal(cabecalhos, itens, rowFn) {
        var html = '<div style="overflow-x:auto"><table class="modal-table"><thead><tr>';
        cabecalhos.forEach(function(c){ html += "<th>" + c + "</th>"; });
        html += "</tr></thead><tbody>";
        if (!itens || itens.length === 0) {
            html += '<tr><td colspan="' + cabecalhos.length + '" style="text-align:center;color:var(--text-muted);padding:14px">Nenhum item.</td></tr>';
        } else {
            itens.forEach(function(item){ html += "<tr>" + rowFn(item) + "</tr>"; });
        }
        html += "</tbody></table></div>";
        return html;
    }

    function fecharModal() {
        document.getElementById("modalDetalhe").classList.remove("open");
    }
    document.getElementById("modalDetalhe").addEventListener("click", function(e){
        if (e.target === this) fecharModal();
    });

    /* -------------------------------------------------------
       EXPORTAR CSV
    ------------------------------------------------------- */
    function exportarCSV() {
        if (!HISTORICO_CACHE || !HISTORICO_CACHE.itens) {
            alert("Carregue o histórico antes de exportar.");
            return;
        }
        var linhas = [["Código","Data","Título","Tipo","Autor","Lidas","Pendentes"]];
        HISTORICO_CACHE.itens.forEach(function(item){
            linhas.push([
                item.cod,
                item.data,
                '"' + item.titulo.replace(/"/g,'""') + '"',
                item.tipo,
                '"' + item.autor.replace(/"/g,'""') + '"',
                item.sim,
                item.nao
            ]);
        });
        var csv = linhas.map(function(l){ return l.join(";"); }).join("\r\n");
        var blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
        var url  = URL.createObjectURL(blob);
        var a    = document.createElement("a");
        a.href = url;
        a.download = "relatorio_central_pg" + HISTORICO_CACHE.pagina + ".csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    /* -------------------------------------------------------
       HELPER: renderizar charts
    ------------------------------------------------------- */
    function renderChart(id, tipo, labels, valores, datasetLabel, extraOpts) {
        if (chartInstancias[id]) { chartInstancias[id].destroy(); }
        var ctx = document.getElementById(id);
        if (!ctx) return;

        var isHorizontal = extraOpts.indexAxis === "y";
        var datasetBase = {
            label: datasetLabel,
            data: valores,
            backgroundColor: extraOpts.backgroundColor || "rgba(0,133,122,0.72)",
            borderColor:     extraOpts.borderColor     || "#00857A",
            borderWidth: 2,
            borderRadius: extraOpts.borderRadius || 0,
            fill:     extraOpts.fill     || false,
            tension:  extraOpts.tension  || 0,
            pointBackgroundColor: extraOpts.pointBackgroundColor || "#00857A",
            indexAxis: extraOpts.indexAxis
        };
        if (extraOpts.indexAxis) delete datasetBase.indexAxis;

        chartInstancias[id] = new Chart(ctx, {
            type: tipo,
            data: { labels: labels, datasets: [datasetBase] },
            options: {
                indexAxis: extraOpts.indexAxis || "x",
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: { label: function(ctx){
                        return " " + ctx.dataset.label + ": " + ctx.parsed[isHorizontal?"x":"y"];
                    }}}
                },
                scales: {
                    x: { grid: { color: "rgba(0,0,0,0.05)" }, ticks: { font: { size: 11 } } },
                    y: { grid: { color: "rgba(0,0,0,0.05)" }, ticks: { font: { size: 11 } } }
                }
            }
        });
    }

    function renderChartDoughnut(id, labels, valores) {
        if (chartInstancias[id]) { chartInstancias[id].destroy(); }
        var ctx = document.getElementById(id);
        if (!ctx) return;
        chartInstancias[id] = new Chart(ctx, {
            type: "doughnut",
            data: {
                labels: labels,
                datasets: [{
                    data: valores,
                    backgroundColor: CORES_TEAL,
                    borderWidth: 2,
                    borderColor: "#fff"
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { position: "right", labels: { font: { size: 11 }, boxWidth: 14 } },
                    tooltip: { callbacks: { label: function(ctx){
                        return " " + ctx.label + ": " + ctx.parsed;
                    }}}
                }
            }
        });
    }

    /* -------------------------------------------------------
       UTIL
    ------------------------------------------------------- */
    function escHtml(str) {
        return (""+str)
            .replace(/&/g,"&amp;")
            .replace(/</g,"&lt;")
            .replace(/>/g,"&gt;")
            .replace(/"/g,"&quot;");
    }

    </script>

</body>
</html>

<%
' Sem necessidade de redirect final – a tela já valida session no topo
%>

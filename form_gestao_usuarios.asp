<!-- #include file="conexao.asp" -->
<!-- #include file="escopo_usuarios.asp" -->
<%
' ── Proteção de acesso ──────────────────────────────────────
If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

If Not ESCOPO_ACESSO Then
    Response.Redirect("form_home.asp")
    Response.End
End If

Function HtmlSafe(v) : HtmlSafe = Server.HTMLEncode("" & v) : End Function

' ── Paginação ───────────────────────────────────────────────
Dim PG : PG = Trim("" & Request("pg"))
If PG = "" Or Not IsNumeric(PG) Then PG = 1
PG = CInt(PG)
If PG < 1 Then PG = 1
Const LIMITE = 20

' ── Filtros via GET ─────────────────────────────────────────
Dim fBusca   : fBusca   = Trim("" & Request("busca"))
Dim fNivel   : fNivel   = Trim("" & Request("nivel"))
Dim fCoor    : fCoor    = Trim("" & Request("coor"))
Dim fFuncao  : fFuncao  = Trim("" & Request("funcao"))

' ── WHERE dinâmico ──────────────────────────────────────────
Dim wExtra : wExtra = GetEscopoWhere()

If fBusca <> "" Then
    wExtra = wExtra & " AND (TBL_USUARIO.NOME LIKE '%" & _
             Replace(fBusca,"'","''") & "%' OR TBL_USUARIO.MATRICULA LIKE '%" & _
             Replace(fBusca,"'","''") & "%')"
End If
If fNivel <> "" Then
    wExtra = wExtra & " AND TBL_USUARIO.NIVEL = '" & Replace(fNivel,"'","''") & "'"
End If
If fCoor <> "" And IsNumeric(fCoor) Then
    wExtra = wExtra & " AND TBL_USUARIO.LOCALIDADE = " & CLng(fCoor)
End If
If fFuncao <> "" Then
    wExtra = wExtra & " AND TBL_USUARIO.FUNCAO LIKE '%" & Replace(fFuncao,"'","''") & "%'"
End If

' ── Total para paginação ────────────────────────────────────
Dim sqlTotal : sqlTotal = "SELECT Count(*) AS T FROM TBL_USUARIO WHERE 1=1 " & wExtra
Dim rsTot    : Set rsTot = conexao_.Execute(sqlTotal)
Dim TOTAL    : TOTAL = CLng(rsTot("T"))
rsTot.Close

Dim TOTAL_PG : TOTAL_PG = 1
If TOTAL > 0 Then TOTAL_PG = Int((TOTAL + LIMITE - 1) / LIMITE)
If PG > TOTAL_PG Then PG = TOTAL_PG
Dim INICIO : INICIO = (PG - 1) * LIMITE

' ── Query principal (Access não tem OFFSET, usamos subquery) ─
Dim sqlBase : sqlBase = _
    "SELECT TBL_USUARIO.COD, TBL_USUARIO.MATRICULA, TBL_USUARIO.NOME, " & _
    "TBL_USUARIO.FUNCAO, TBL_USUARIO.TITULO, TBL_USUARIO.NIVEL, " & _
    "TBL_COOR.COOR AS LOCALIDADE_NOME " & _
    "FROM TBL_USUARIO " & _
    "LEFT JOIN TBL_COOR ON TBL_USUARIO.LOCALIDADE = TBL_COOR.ID " & _
    "WHERE 1=1 " & wExtra

Dim sqlLista
If INICIO = 0 Then
    sqlLista = "SELECT TOP " & LIMITE & " " & _
               Mid(sqlBase, InStr(sqlBase, "TBL_USUARIO.COD")) & _
               " ORDER BY TBL_USUARIO.NOME ASC"
    sqlLista = "SELECT TOP " & LIMITE & " " & _
               "TBL_USUARIO.COD, TBL_USUARIO.MATRICULA, TBL_USUARIO.NOME, " & _
               "TBL_USUARIO.FUNCAO, TBL_USUARIO.TITULO, TBL_USUARIO.NIVEL, " & _
               "TBL_COOR.COOR AS LOCALIDADE_NOME " & _
               "FROM TBL_USUARIO " & _
               "LEFT JOIN TBL_COOR ON TBL_USUARIO.LOCALIDADE = TBL_COOR.ID " & _
               "WHERE 1=1 " & wExtra & " ORDER BY TBL_USUARIO.NOME ASC"
Else
    Dim sqlInner : sqlInner = _
        "SELECT TOP " & INICIO & " TBL_USUARIO.COD " & _
        "FROM TBL_USUARIO " & _
        "LEFT JOIN TBL_COOR ON TBL_USUARIO.LOCALIDADE = TBL_COOR.ID " & _
        "WHERE 1=1 " & wExtra & " ORDER BY TBL_USUARIO.NOME ASC"
    sqlLista = _
        "SELECT TOP " & LIMITE & " " & _
        "TBL_USUARIO.COD, TBL_USUARIO.MATRICULA, TBL_USUARIO.NOME, " & _
        "TBL_USUARIO.FUNCAO, TBL_USUARIO.TITULO, TBL_USUARIO.NIVEL, " & _
        "TBL_COOR.COOR AS LOCALIDADE_NOME " & _
        "FROM TBL_USUARIO " & _
        "LEFT JOIN TBL_COOR ON TBL_USUARIO.LOCALIDADE = TBL_COOR.ID " & _
        "WHERE 1=1 " & wExtra & _
        " AND TBL_USUARIO.COD NOT IN (" & sqlInner & ")" & _
        " ORDER BY TBL_USUARIO.NOME ASC"
End If

Dim rsLista : Set rsLista = conexao_.Execute(sqlLista)

' ── COOR disponíveis para filtro (dentro do escopo) ─────────
Dim sqlCoor : sqlCoor = _
    "SELECT DISTINCT TBL_COOR.ID, TBL_COOR.COOR " & _
    "FROM TBL_USUARIO " & _
    "INNER JOIN TBL_COOR ON TBL_USUARIO.LOCALIDADE = TBL_COOR.ID " & _
    "WHERE 1=1 " & GetEscopoWhere() & _
    " ORDER BY TBL_COOR.COOR ASC"
Dim rsCoor : Set rsCoor = conexao_.Execute(sqlCoor)

' ── Mensagem de feedback (set por valida_nivel.asp) ─────────
Dim msgFeedback : msgFeedback = ""
If Session("msg_gestao") <> "" Then
    msgFeedback = "" & Session("msg_gestao")
    Session("msg_gestao") = ""
End If

' ── Monta QS atual para paginação manter filtros ────────────
Dim qsFiltros : qsFiltros = ""
If fBusca  <> "" Then qsFiltros = qsFiltros & "&busca="  & Server.URLEncode(fBusca)
If fNivel  <> "" Then qsFiltros = qsFiltros & "&nivel="  & Server.URLEncode(fNivel)
If fCoor   <> "" Then qsFiltros = qsFiltros & "&coor="   & Server.URLEncode(fCoor)
If fFuncao <> "" Then qsFiltros = qsFiltros & "&funcao=" & Server.URLEncode(fFuncao)
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Gestão de Empregados – Central de Informações</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="libs/js/jquery.js"></script>
<style>
:root {
    --vale-teal:      #00857A;
    --vale-teal-dark: #006B63;
    --vale-yellow:    #F6B800;
    --text-dark:      #263238;
    --text-muted:     #607D8B;
    --card-bg:        rgba(255,255,255,0.92);
    --sidebar-bg:     rgba(255,255,255,0.86);
    --sidebar-collapsed: 88px;
    --sidebar-expanded:  292px;
    --danger:   #C62828;
    --success:  #2E7D32;
    --amber:    #F59E0B;
}
* { box-sizing: border-box; }
html, body { width:100%; min-height:100vh; margin:0; overflow-x:hidden; }
body {
    font-family: Arial, Helvetica, sans-serif;
    background:
        linear-gradient(rgba(0,60,55,0.16), rgba(0,60,55,0.16)),
        url("libs/img/login_bg.jpg") center center / cover no-repeat fixed;
    color: var(--text-dark);
}

/* ── SIDEBAR ── */
.sidebar {
    position:fixed; left:0; top:0;
    width:var(--sidebar-collapsed); height:100vh;
    background:var(--sidebar-bg);
    backdrop-filter:blur(10px); -webkit-backdrop-filter:blur(10px);
    box-shadow:5px 0 24px rgba(0,0,0,0.18);
    z-index:100; transition:width 0.25s ease-in-out;
    display:flex; flex-direction:column;
    overflow:hidden; padding:14px 12px;
}
.sidebar.open { width:var(--sidebar-expanded); }
.sidebar-header {
    flex-shrink:0; display:flex; flex-direction:column;
    align-items:center; gap:14px; padding-bottom:14px;
    border-bottom:1px solid rgba(0,133,122,0.12);
}
.menu-toggle {
    width:56px; height:56px; border:none; border-radius:18px;
    background:var(--vale-teal); color:#fff;
    font-size:26px; font-weight:800; cursor:pointer; transition:all 0.2s;
}
.menu-toggle:hover { background:var(--vale-teal-dark); transform:translateY(-1px); }
.sidebar-logo-area { width:100%; display:flex; justify-content:center; min-height:54px; }
.sidebar-logo-area img { width:58px; height:auto; transition:all 0.2s; }
.sidebar.open .sidebar-logo-area img { width:118px; }
.sidebar-nav { flex:1 1 auto; overflow-y:auto; display:flex; flex-direction:column; gap:11px; padding:14px 0; }
.sidebar-footer { flex-shrink:0; padding-top:12px; border-top:1px solid rgba(0,133,122,0.12); }
.sidebar-link {
    width:100%; min-height:58px; border:none; border-radius:18px;
    background:rgba(255,255,255,0.78); color:var(--text-dark); text-decoration:none;
    display:flex; align-items:center; gap:14px; padding:0 12px; cursor:pointer;
    box-shadow:0 7px 16px rgba(0,0,0,0.10); transition:all 0.2s;
    font-family:Arial,Helvetica,sans-serif;
}
.sidebar-link:hover { background:#fff; transform:translateX(4px); box-shadow:0 12px 25px rgba(0,0,0,0.18); }
.sidebar-link.active { background:var(--vale-teal); color:#fff; }
.sidebar-link.active .sidebar-title { color:#fff; }
.sidebar:not(.open) .sidebar-link { justify-content:center; padding:0; }
.sidebar-icon {
    width:38px; height:38px; border-radius:14px;
    background:var(--vale-teal); color:#fff;
    display:flex; align-items:center; justify-content:center;
    flex-shrink:0; overflow:hidden; font-weight:800; font-size:17px;
}
.sidebar-icon img { width:100%; height:100%; object-fit:cover; }
.sidebar-text { display:none; flex-direction:column; text-align:left; line-height:1.15; white-space:nowrap; }
.sidebar.open .sidebar-text { display:flex; }
.sidebar-title    { font-size:13px; font-weight:800; color:var(--vale-teal-dark); }
.sidebar-subtitle { font-size:11px; color:var(--text-muted); margin-top:2px; }

/* ── CONTEÚDO ── */
.main-content {
    min-height:100vh;
    margin-left:var(--sidebar-collapsed);
    width:calc(100vw - var(--sidebar-collapsed));
    padding:clamp(20px,2.8vw,40px);
    transition:margin-left 0.25s,width 0.25s;
}
body.sidebar-open .main-content {
    margin-left:var(--sidebar-expanded);
    width:calc(100vw - var(--sidebar-expanded));
}

.page-header { width:min(100%,1200px); margin:0 auto 18px; text-align:center; }
.page-title { font-size:clamp(18px,2.2vw,28px); font-weight:800; color:#fff; text-shadow:0 2px 10px rgba(0,0,0,0.35); margin:0 0 6px; }
.page-sub   { font-size:13px; color:rgba(255,255,255,0.82); }

/* ── ALERTAS ── */
.alert {
    width:min(100%,1200px); margin:0 auto 16px;
    padding:12px 18px; border-radius:14px;
    font-size:13px; font-weight:700;
}
.alert-ok  { background:rgba(46,125,50,0.15); color:#1B5E20; border:1.5px solid rgba(46,125,50,0.3); }
.alert-err { background:rgba(198,40,40,0.12); color:#7f0000; border:1.5px solid rgba(198,40,40,0.25); }

/* ── FILTROS ── */
.filtros-card {
    width:min(100%,1200px); margin:0 auto 18px;
    background:var(--card-bg); border-radius:20px;
    padding:16px 20px; box-shadow:0 8px 24px rgba(0,0,0,0.12);
    display:flex; flex-wrap:wrap; gap:10px; align-items:flex-end;
}
.filtros-card label { font-size:12px; font-weight:700; color:var(--text-muted); display:flex; flex-direction:column; gap:4px; min-width:140px; }
.filtros-card input,
.filtros-card select { padding:8px 12px; border:1.5px solid rgba(0,133,122,0.22); border-radius:10px; font-size:13px; color:var(--text-dark); background:#fff; outline:none; transition:border-color 0.2s; }
.filtros-card input:focus,
.filtros-card select:focus { border-color:var(--vale-teal); }
.btn-filtrar { padding:10px 20px; border:none; border-radius:12px; background:var(--vale-teal); color:#fff; font-size:13px; font-weight:700; cursor:pointer; transition:all 0.2s; }
.btn-filtrar:hover { background:var(--vale-teal-dark); transform:translateY(-1px); }
.btn-limpar  { padding:10px 18px; border:none; border-radius:12px; background:rgba(0,133,122,0.08); color:var(--vale-teal-dark); font-size:13px; font-weight:700; cursor:pointer; transition:all 0.2s; }
.btn-limpar:hover { background:rgba(0,133,122,0.16); }

/* ── TABELA ── */
.table-card {
    width:min(100%,1200px); margin:0 auto 28px;
    background:var(--card-bg); border-radius:20px;
    box-shadow:0 8px 24px rgba(0,0,0,0.12); overflow:hidden;
}
.table-info-bar {
    padding:14px 20px; border-bottom:1px solid rgba(0,133,122,0.08);
    display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;
}
.table-info-bar span { font-size:13px; color:var(--text-muted); }
.scope-badge {
    font-size:12px; font-weight:700; padding:4px 12px; border-radius:8px;
    background:rgba(0,133,122,0.10); color:var(--vale-teal-dark);
}

.gest-table { width:100%; border-collapse:collapse; }
.gest-table thead tr { background:linear-gradient(135deg,var(--vale-teal),var(--vale-teal-dark)); color:#fff; }
.gest-table th { padding:12px 14px; font-size:12px; font-weight:700; text-align:left; white-space:nowrap; letter-spacing:0.04em; }
.gest-table tbody tr { border-bottom:1px solid rgba(0,133,122,0.07); transition:background 0.15s; }
.gest-table tbody tr:hover { background:rgba(0,133,122,0.04); }
.gest-table td { padding:10px 14px; font-size:13px; vertical-align:middle; }
.mat-badge {
    font-size:11px; font-weight:800; padding:3px 8px; border-radius:7px;
    background:rgba(0,133,122,0.10); color:var(--vale-teal-dark);
}

/* Badges de nível */
.nivel-badge { display:inline-block; font-size:11px; font-weight:700; padding:3px 9px; border-radius:8px; white-space:nowrap; }
.nivel-ADM_DEV { background:rgba(198,40,40,0.12);  color:#7f0000; }
.nivel-ADM_GG  { background:rgba(0,133,122,0.14);  color:var(--vale-teal-dark); }
.nivel-ADM_GA  { background:rgba(0,133,122,0.10);  color:#005C55; }
.nivel-ADM_DEL { background:rgba(245,158,11,0.15); color:#92400E; }
.nivel-USER    { background:rgba(96,125,139,0.12); color:#37474F; }

/* Botão editar */
.btn-editar {
    padding:6px 14px; border:1.5px solid var(--vale-teal);
    background:transparent; color:var(--vale-teal-dark);
    border-radius:10px; font-size:12px; font-weight:700;
    cursor:pointer; transition:all 0.18s;
}
.btn-editar:hover { background:var(--vale-teal); color:#fff; }
.btn-editar:disabled { opacity:0.35; cursor:default; border-color:var(--text-muted); color:var(--text-muted); }

/* ── PAGINAÇÃO ── */
.pagination { display:flex; justify-content:center; gap:6px; padding:16px; }
.page-btn {
    min-width:36px; height:36px; border:1.5px solid rgba(0,133,122,0.22);
    background:#fff; color:var(--vale-teal-dark);
    border-radius:10px; font-size:13px; font-weight:700;
    cursor:pointer; transition:all 0.18s;
}
.page-btn:hover, .page-btn.active { background:var(--vale-teal); color:#fff; border-color:var(--vale-teal); }
.page-btn:disabled { opacity:0.38; cursor:default; }

/* ── MODAL NÍVEL ── */
.modal-overlay {
    display:none; position:fixed; inset:0;
    background:rgba(0,0,0,0.50); backdrop-filter:blur(4px);
    z-index:200; align-items:center; justify-content:center;
}
.modal-overlay.open { display:flex; }
.modal-box {
    background:#fff; border-radius:22px;
    width:min(92vw,460px); box-shadow:0 24px 56px rgba(0,0,0,0.28);
}
.modal-header {
    background:linear-gradient(135deg,var(--vale-teal),var(--vale-teal-dark));
    color:#fff; padding:16px 20px; border-radius:22px 22px 0 0;
    display:flex; justify-content:space-between; align-items:center;
}
.modal-title { font-size:15px; font-weight:800; }
.modal-close { background:rgba(255,255,255,0.18); border:none; color:#fff; width:32px; height:32px; border-radius:9px; font-size:18px; cursor:pointer; }
.modal-body  { padding:22px; }
.modal-body p { font-size:13px; color:var(--text-muted); margin:0 0 16px; }
.modal-body strong { color:var(--text-dark); }
.modal-body label { display:block; font-size:12px; font-weight:700; color:var(--text-muted); margin-bottom:6px; }
.modal-body select {
    width:100%; padding:10px 12px;
    border:1.5px solid rgba(0,133,122,0.25); border-radius:12px;
    font-size:14px; color:var(--text-dark); background:#fff; outline:none;
    transition:border-color 0.2s;
}
.modal-body select:focus { border-color:var(--vale-teal); }
.modal-aviso {
    margin-top:14px; padding:10px 14px; border-radius:10px;
    background:rgba(245,158,11,0.12); border:1.5px solid rgba(245,158,11,0.3);
    font-size:12px; color:#92400E;
}
.modal-footer { padding:14px 20px; text-align:right; border-top:1px solid rgba(0,0,0,0.07); display:flex; gap:10px; justify-content:flex-end; }
.btn-cancelar { padding:10px 20px; border:1.5px solid rgba(0,0,0,0.12); background:transparent; border-radius:12px; font-size:13px; font-weight:700; cursor:pointer; color:var(--text-muted); }
.btn-salvar   { padding:10px 22px; border:none; border-radius:12px; background:var(--vale-teal); color:#fff; font-size:13px; font-weight:700; cursor:pointer; transition:background 0.2s; }
.btn-salvar:hover { background:var(--vale-teal-dark); }
</style>
</head>
<body>

<!-- ── SIDEBAR ── -->
<aside id="sidebar" class="sidebar">
    <div class="sidebar-header">
        <button type="button" class="menu-toggle" onclick="toggleSidebar()">☰</button>
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
        <a href="form_gestao_usuarios.asp" class="sidebar-link active">
            <span class="sidebar-icon">👥</span>
            <span class="sidebar-text">
                <span class="sidebar-title">Gestão de empregados</span>
                <span class="sidebar-subtitle"><%=HtmlSafe(GetEscopoLabel())%></span>
            </span>
        </a>
        <a href="form_relatorios.asp" class="sidebar-link">
            <span class="sidebar-icon">📊</span>
            <span class="sidebar-text"><span class="sidebar-title">Dashboard</span></span>
        </a>
    </nav>
    <div class="sidebar-footer">
        <a class="sidebar-link" href="fecha_session.asp">
            <span class="sidebar-icon">S</span>
            <span class="sidebar-text"><span class="sidebar-title">Sair</span></span>
        </a>
    </div>
</aside>

<!-- ── CONTEÚDO ── -->
<main class="main-content">

    <header class="page-header">
        <h1 class="page-title">👥 Gestão de Empregados</h1>
        <p class="page-sub"><%=HtmlSafe(GetEscopoLabel())%> · logado como <strong><%=HtmlSafe(Session("name"))%></strong> (<%=HtmlSafe(ESCOPO_NIVEL)%>)</p>
    </header>

    <% If msgFeedback <> "" Then %>
    <div class="alert <%=IIf(InStr(msgFeedback,"sucesso")>0,"alert-ok","alert-err")%>">
        <%=HtmlSafe(msgFeedback)%>
    </div>
    <% End If %>

    <!-- FILTROS -->
    <form method="GET" action="form_gestao_usuarios.asp">
    <div class="filtros-card">
        <label>
            Busca (nome ou matrícula)
            <input type="text" name="busca" value="<%=HtmlSafe(fBusca)%>" placeholder="ex: João Silva">
        </label>
        <label>
            Nível
            <select name="nivel">
                <option value="">Todos</option>
                <option value="USER"    <%=IIf(fNivel="USER",    "selected","")%>>USER</option>
                <option value="ADM_DEL" <%=IIf(fNivel="ADM_DEL","selected","")%>>ADM_DEL</option>
                <% If ESCOPO_NIVEL = "ADM_DEV" Then %>
                <option value="ADM_GA"  <%=IIf(fNivel="ADM_GA", "selected","")%>>ADM_GA</option>
                <option value="ADM_GG"  <%=IIf(fNivel="ADM_GG", "selected","")%>>ADM_GG</option>
                <% End If %>
            </select>
        </label>
        <label>
            Localidade / COOR
            <select name="coor">
                <option value="">Todas</option>
                <%
                Do Until rsCoor.EOF
                    Dim coorId   : coorId   = "" & rsCoor("ID")
                    Dim coorNome : coorNome = "" & rsCoor("LOCALIDADE_NOME")
                    Dim sel      : sel      = IIf(fCoor = coorId, " selected", "")
                %>
                <option value="<%=HtmlSafe(coorId)%>"<%=sel%>><%=HtmlSafe(coorNome)%></option>
                <%
                    rsCoor.MoveNext
                Loop
                rsCoor.Close
                %>
            </select>
        </label>
        <button type="submit" class="btn-filtrar">Filtrar</button>
        <button type="button" class="btn-limpar" onclick="window.location='form_gestao_usuarios.asp'">Limpar</button>
    </div>
    </form>

    <!-- TABELA -->
    <div class="table-card">
        <div class="table-info-bar">
            <span><%=TOTAL%> empregado(s) encontrado(s) · Página <%=PG%> de <%=TOTAL_PG%></span>
            <span class="scope-badge">Escopo: <%=HtmlSafe(GetEscopoLabel())%></span>
        </div>

        <table class="gest-table">
            <thead>
                <tr>
                    <th>Matrícula</th>
                    <th>Nome</th>
                    <th>Função</th>
                    <th>Cargo</th>
                    <th>Localidade</th>
                    <th>Nível atual</th>
                    <th></th>
                </tr>
            </thead>
            <tbody>
            <%
            If rsLista.EOF Then
            %>
                <tr>
                    <td colspan="7" style="text-align:center;padding:28px;color:var(--text-muted)">
                        Nenhum empregado encontrado neste escopo.
                    </td>
                </tr>
            <%
            Else
                Do Until rsLista.EOF
                    Dim uCod     : uCod     = "" & rsLista("COD")
                    Dim uMat     : uMat     = "" & rsLista("MATRICULA")
                    Dim uNome    : uNome    = "" & rsLista("NOME")
                    Dim uFuncao  : uFuncao  = "" & rsLista("FUNCAO")
                    Dim uTitulo  : uTitulo  = "" & rsLista("TITULO")
                    Dim uNivel   : uNivel   = "" & rsLista("NIVEL")
                    Dim uLoc     : uLoc     = "" & rsLista("LOCALIDADE_NOME")

                    ' Botão desabilitado se o nível do alvo não pode ser alterado pelo logado
                    Dim podeDel  : podeDel  = PodeEditarUsuario(uMat)
                    Dim btnDis   : btnDis   = IIf(podeDel, "", " disabled")
            %>
                <tr>
                    <td><span class="mat-badge"><%=HtmlSafe(uMat)%></span></td>
                    <td><%=HtmlSafe(uNome)%></td>
                    <td><%=HtmlSafe(uFuncao)%></td>
                    <td style="font-size:12px;color:var(--text-muted)"><%=HtmlSafe(uTitulo)%></td>
                    <td style="font-size:12px"><%=HtmlSafe(uLoc)%></td>
                    <td><span class="nivel-badge nivel-<%=HtmlSafe(uNivel)%>"><%=HtmlSafe(uNivel)%></span></td>
                    <td>
                        <button class="btn-editar" <%=btnDis%>
                            onclick="abrirModalNivel('<%=HtmlSafe(uMat)%>','<%=HtmlSafe(Replace(uNome,"'","\\'"))%>','<%=HtmlSafe(uNivel)%>')">
                            Alterar nível
                        </button>
                    </td>
                </tr>
            <%
                    rsLista.MoveNext
                Loop
            End If
            rsLista.Close
            %>
            </tbody>
        </table>

        <!-- Paginação -->
        <div class="pagination">
            <%
            Dim pgUrl : pgUrl = "form_gestao_usuarios.asp?pg="
            Dim btnPrev : btnPrev = IIf(PG<=1," disabled","")
            Dim btnNext : btnNext = IIf(PG>=TOTAL_PG," disabled","")
            %>
            <button class="page-btn" <%=btnPrev%>
                onclick="window.location='<%=pgUrl%><%=PG-1%><%=qsFiltros%>'">‹</button>
            <%
            Dim pi, pIni, pFim
            pIni = WorksheetFunction.Max(1, PG-3)
            pFim = WorksheetFunction.Min(TOTAL_PG, PG+3)
            If pIni > 1 Then
                Response.Write "<button class='page-btn' onclick=""window.location='" & pgUrl & "1" & qsFiltros & "'"">1</button>"
                If pIni > 2 Then Response.Write "<span style='padding:0 4px;line-height:36px'>…</span>"
            End If
            For pi = pIni To pFim
                Dim isActive : isActive = IIf(pi=PG," active","")
                Response.Write "<button class='page-btn" & isActive & "' onclick=""window.location='" & pgUrl & pi & qsFiltros & "'"">"+CStr(pi)+"</button>"
            Next
            If pFim < TOTAL_PG Then
                If pFim < TOTAL_PG - 1 Then Response.Write "<span style='padding:0 4px;line-height:36px'>…</span>"
                Response.Write "<button class='page-btn' onclick=""window.location='" & pgUrl & TOTAL_PG & qsFiltros & "'"">"+CStr(TOTAL_PG)+"</button>"
            End If
            %>
            <button class="page-btn" <%=btnNext%>
                onclick="window.location='<%=pgUrl%><%=PG+1%><%=qsFiltros%>'">›</button>
        </div>
    </div>

</main>

<!-- ── MODAL ALTERAR NÍVEL ── -->
<div id="modalNivel" class="modal-overlay">
    <div class="modal-box">
        <div class="modal-header">
            <div class="modal-title">Alterar nível de acesso</div>
            <button type="button" class="modal-close" onclick="fecharModal()">✕</button>
        </div>
        <div class="modal-body">
            <p>
                Empregado: <strong id="mdNome"></strong><br>
                Matrícula: <strong id="mdMat"></strong><br>
                Nível atual: <strong id="mdNivelAtual"></strong>
            </p>
            <label>Novo nível</label>
            <select id="mdNovoNivel">
                <option value="USER">USER – Somente leitura</option>
                <option value="ADM_DEL">ADM_DEL – Delegado (pode publicar avisos)</option>
                <% If ESCOPO_NIVEL = "ADM_DEV" Then %>
                <option value="ADM_GA">ADM_GA – Gerente de Área</option>
                <option value="ADM_GG">ADM_GG – Gerente Geral</option>
                <% End If %>
            </select>
            <div class="modal-aviso">
                ⚠️ Esta ação será registrada. O empregado passará a ter as permissões do novo nível imediatamente após salvar.
            </div>
        </div>
        <div class="modal-footer">
            <button type="button" class="btn-cancelar" onclick="fecharModal()">Cancelar</button>
            <button type="button" class="btn-salvar"   onclick="confirmarAlteracao()">Salvar</button>
        </div>
    </div>
</div>

<!-- Form oculto de submissão -->
<form id="formNivel" method="POST" action="valida_nivel.asp" style="display:none">
    <input type="hidden" name="matricula"  id="hMat">
    <input type="hidden" name="novo_nivel" id="hNivel">
    <input type="hidden" name="pg"         value="<%=PG%>">
    <input type="hidden" name="filtros"    value="<%=HtmlSafe(qsFiltros)%>">
</form>

<script>
function toggleSidebar() {
    var s = document.getElementById("sidebar");
    var b = document.body;
    if (s.classList.contains("open")) {
        s.classList.remove("open"); b.classList.remove("sidebar-open");
        localStorage.setItem("centralSidebarOpen","N");
    } else {
        s.classList.add("open"); b.classList.add("sidebar-open");
        localStorage.setItem("centralSidebarOpen","S");
    }
}
window.addEventListener("DOMContentLoaded", function() {
    if (localStorage.getItem("centralSidebarOpen") === "S") {
        document.getElementById("sidebar").classList.add("open");
        document.body.classList.add("sidebar-open");
    }
});

function abrirModalNivel(mat, nome, nivelAtual) {
    document.getElementById("mdNome").textContent       = nome;
    document.getElementById("mdMat").textContent        = mat;
    document.getElementById("mdNivelAtual").textContent = nivelAtual;
    document.getElementById("hMat").value               = mat;
    // Pré-seleciona o nível atual
    var sel = document.getElementById("mdNovoNivel");
    for (var i = 0; i < sel.options.length; i++) {
        sel.options[i].selected = (sel.options[i].value === nivelAtual);
    }
    document.getElementById("modalNivel").classList.add("open");
}

function fecharModal() {
    document.getElementById("modalNivel").classList.remove("open");
}

function confirmarAlteracao() {
    var novoNivel = document.getElementById("mdNovoNivel").value;
    var nivelAtual = document.getElementById("mdNivelAtual").textContent;
    if (novoNivel === nivelAtual) {
        alert("O nível selecionado é igual ao nível atual. Nenhuma alteração necessária.");
        return;
    }
    document.getElementById("hNivel").value = novoNivel;
    document.getElementById("formNivel").submit();
}

document.getElementById("modalNivel").addEventListener("click", function(e) {
    if (e.target === this) fecharModal();
});
</script>

</body>
</html>

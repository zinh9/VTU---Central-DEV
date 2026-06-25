<!-- #include file="conexao.asp" -->
<%
codigo = Request("codigo")

If codigo <> "" Then
    whereSQL = whereSQL & " AND TBL_INFORMACAO.COD=" & codigo
End If
' =========================================================
' SEGURANÇA BÁSICA
' =========================================================
If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
End If

' =========================================================
' FUNÇÕES AUXILIARES
' =========================================================
Function SqlSafe(v)
    SqlSafe = Replace(Trim("" & v), "'", "''")
End Function

Function HtmlSafe(v)
    HtmlSafe = Server.HTMLEncode("" & v)
End Function

Function TextoCurto(v, limite)
    Dim t
    t = "" & v
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    If Len(t) > limite Then
        t = Left(t, limite) & "..."
    End If
    TextoCurto = Server.HTMLEncode(t)
End Function

Function TratarImagem(v)
    Dim arq
    arq = Trim("" & v)
    If arq = "" Or UCase(arq) = "NOIMG" Then
        TratarImagem = "NOIMG.jpg"
    Else
        TratarImagem = Replace(arq, "\", "/")
    End If
End Function

Function EhImagem(v)
    Dim arq
    arq = LCase("" & v)
    If InStr(arq, ".jpg") > 0 Or InStr(arq, ".jpeg") > 0 Or InStr(arq, ".png") > 0 Or InStr(arq, ".gif") > 0 Or InStr(arq, ".jfif") > 0 Then
        EhImagem = True
    Else
        EhImagem = False
    End If
End Function

Function TipoDocumento(v)
    Dim arq
    arq = LCase("" & v)
    If InStr(arq, ".pdf") > 0 Then
        TipoDocumento = "PDF"
    ElseIf InStr(arq, ".pptx") > 0 Or InStr(arq, ".ppt") > 0 Then
        TipoDocumento = "PPT"
    ElseIf InStr(arq, ".png") > 0 Or InStr(arq, ".jpg") > 0 Or InStr(arq, ".jpeg") > 0 Then
        TipoDocumento = "IMG"
    Else
        TipoDocumento = "ARQ"
    End If
End Function

Function BuildPageUrl(pg, busca, nome, tipo, dataTxt, status)
    Dim u
    u = "form_visualizar_registro.asp?pg=" & pg
    If busca <> "" Then u = u & "&busca=" & Server.URLEncode(busca)
    If nome <> "" Then u = u & "&nome=" & Server.URLEncode(nome)
    If tipo <> "" Then u = u & "&tipo=" & Server.URLEncode(tipo)
    If dataTxt <> "" Then u = u & "&data_txt=" & Server.URLEncode(dataTxt)
    If status <> "" Then u = u & "&status=" & Server.URLEncode(status)
    BuildPageUrl = u
End Function

' =========================================================
' DADOS DE SESSÃO
' =========================================================
usuarioNome = "" & Session("name")
usuarioFuncao = "" & Session("funcao")
usuarioNivel = "" & Session("nivel")
usuarioMatricula = "" & Session("matricula")

' =========================================================
' FILTROS
' =========================================================
busca = Trim("" & Request("busca"))      ' título / conteúdo
nome  = Trim("" & Request("nome"))       ' autor
tipo  = Trim("" & Request("tipo"))       ' tipo
dataTxt = Trim("" & Request("data_txt")) ' data (texto parcial para lidar com formatos mistos)
statusFiltro = Trim("" & Request("status")) ' PENDENTE / VISUALIZADO / vazio

' =========================================================
' PAGINAÇÃO
' =========================================================
pagina = Trim("" & Request("pg"))
If pagina = "" Then pagina = "1"
If Not IsNumeric(pagina) Then pagina = "1"

pagina = CInt(pagina)
If pagina < 1 Then pagina = 1

limite = 8
inicio = (pagina - 1) * limite

' =========================================================
' MONTA WHERE DE FORMA SEGURA
' =========================================================
whereSQL = " WHERE TBL_CHECK_INFORMACAO.MATRICULA='" & SqlSafe(usuarioMatricula) & "' "

If busca <> "" Then
    whereSQL = whereSQL & " AND (TBL_INFORMACAO.TITULO LIKE '%" & SqlSafe(busca) & "%' " & _
                          " OR TBL_INFORMACAO.INFORMACAO LIKE '%" & SqlSafe(busca) & "%') "
End If

If nome <> "" Then
    whereSQL = whereSQL & " AND TBL_USUARIO.NOME LIKE '%" & SqlSafe(nome) & "%' "
End If

If tipo <> "" Then
    whereSQL = whereSQL & " AND TBL_INFORMACAO.TIPO='" & SqlSafe(tipo) & "' "
End If

If dataTxt <> "" Then
    ' Usa LIKE porque a coluna DATA no Access pode estar armazenada em formatos mistos
    whereSQL = whereSQL & " AND TBL_INFORMACAO.DATA LIKE '%" & SqlSafe(dataTxt) & "%' "
End If

If statusFiltro = "PENDENTE" Then
    whereSQL = whereSQL & " AND TBL_CHECK_INFORMACAO.[CHECK]='NÃO' "
ElseIf statusFiltro = "VISUALIZADO" Then
    whereSQL = whereSQL & " AND TBL_CHECK_INFORMACAO.[CHECK]='SIM' "
End If

' =========================================================
' BASE DA CONSULTA
' =========================================================
fromSQL = " FROM (TBL_CHECK_INFORMACAO " & _
          " INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD) " & _
          " INNER JOIN TBL_USUARIO ON TBL_INFORMACAO.FK_MATRICULA = TBL_USUARIO.MATRICULA "

orderSQL = " ORDER BY TBL_INFORMACAO.COD DESC "

' =========================================================
' TOTAL DE REGISTROS PARA PAGINAÇÃO
' =========================================================
sqlCount = "SELECT Count(*) AS TOTAL " & fromSQL & whereSQL
Set rsCount = conexao_.Execute(sqlCount)

totalRegistros = 0
If Not rsCount.EOF Then
    If Not IsNull(rsCount("TOTAL")) Then
        totalRegistros = CLng(rsCount("TOTAL"))
    End If
End If

If limite < 1 Then limite = 8
totalPaginas = 1
If totalRegistros > 0 Then
    totalPaginas = Int((totalRegistros + limite - 1) / limite)
End If
If pagina > totalPaginas Then pagina = totalPaginas
If pagina < 1 Then pagina = 1
inicio = (pagina - 1) * limite

' =========================================================
' CONSULTA PAGINADA (ACCESS LEGADO)
' Estratégia:
' - Página 1: TOP N DESC
' - Página > 1: remove TOP início e pega TOP limite dos restantes
' =========================================================
If inicio = 0 Then
    sqlPage = "SELECT TOP " & limite & _
              " TBL_CHECK_INFORMACAO.[CHECK] AS STATUS_LEITURA, " & _
              " TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO, " & _
              " TBL_INFORMACAO.COD, TBL_INFORMACAO.DATA, TBL_INFORMACAO.TITULO, " & _
              " TBL_INFORMACAO.TIPO, TBL_INFORMACAO.INFORMACAO, TBL_INFORMACAO.IMG_ID, " & _
              " TBL_USUARIO.NOME AS AUTOR " & _
              fromSQL & whereSQL & orderSQL
Else
    sqlPage = "SELECT TOP " & limite & _
              " TBL_CHECK_INFORMACAO.[CHECK] AS STATUS_LEITURA, " & _
              " TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO, " & _
              " TBL_INFORMACAO.COD, TBL_INFORMACAO.DATA, TBL_INFORMACAO.TITULO, " & _
              " TBL_INFORMACAO.TIPO, TBL_INFORMACAO.INFORMACAO, TBL_INFORMACAO.IMG_ID, " & _
              " TBL_USUARIO.NOME AS AUTOR " & _
              fromSQL & whereSQL & _
              " AND TBL_INFORMACAO.COD NOT IN (" & _
                    "SELECT TOP " & inicio & " TBL_INFORMACAO.COD " & fromSQL & whereSQL & orderSQL & _
              ") " & orderSQL
End If

Set RS = conexao_.Execute(sqlPage)
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Visualizar Registros</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>

    /* =====================================================
           MENU LATERAL
        ===================================================== */

        .sidebar {
            position: fixed;
            left: 0;
            top: 0;
            width: var(--sidebar-collapsed);
            height: 100vh;
            height: 100dvh;
            background: var(--sidebar-bg);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            box-shadow: 5px 0 24px rgba(0, 0, 0, 0.18);
            z-index: 100;
            transition: width 0.25s ease-in-out;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            padding: 14px 12px;
        }

        .sidebar.open {
            width: var(--sidebar-expanded);
        }

        .sidebar-header {
            flex-shrink: 0;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 14px;
            padding-bottom: 14px;
            border-bottom: 1px solid rgba(0, 133, 122, 0.12);
        }

        .menu-toggle {
            width: 56px;
            height: 56px;
            border: none;
            border-radius: 18px;
            background: var(--vale-teal);
            color: #ffffff;
            font-size: 26px;
            font-weight: 800;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
        }

        .menu-toggle:hover {
            background: var(--vale-teal-dark);
            transform: translateY(-1px);
        }

        .sidebar-logo-area {
            width: 100%;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 54px;
        }

        .sidebar-logo-area img {
            width: 58px;
            max-width: 58px;
            height: auto;
            transition: all 0.2s ease-in-out;
        }

        .sidebar.open .sidebar-logo-area img {
            width: 118px;
            max-width: 118px;
        }

        .sidebar-nav {
            flex: 1 1 auto;
            overflow-y: auto;
            overflow-x: hidden;
            display: flex;
            flex-direction: column;
            gap: 11px;
            padding-top: 14px;
            padding-bottom: 14px;
        }

        .sidebar-footer {
            flex-shrink: 0;
            padding-top: 12px;
            border-top: 1px solid rgba(0, 133, 122, 0.12);
        }

        .sidebar-nav::-webkit-scrollbar {
            width: 6px;
        }

        .sidebar-nav::-webkit-scrollbar-thumb {
            background: rgba(0, 133, 122, 0.35);
            border-radius: 10px;
        }

        .sidebar-nav::-webkit-scrollbar-track {
            background: transparent;
        }

        .sidebar-link,
        .external-toggle {
            width: 100%;
            min-height: 58px;
            border: none;
            border-radius: 18px;
            background: rgba(255, 255, 255, 0.78);
            color: var(--text-dark);
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 0 12px;
            cursor: pointer;
            box-shadow: 0 7px 16px rgba(0, 0, 0, 0.10);
            transition: all 0.2s ease-in-out;
            font-family: Arial, Helvetica, sans-serif;
        }

        .sidebar-link:hover,
        .external-toggle:hover {
            background: #ffffff;
            transform: translateX(4px);
            box-shadow: 0 12px 25px rgba(0, 0, 0, 0.18);
        }

        .sidebar:not(.open) .sidebar-link,
        .sidebar:not(.open) .external-toggle {
            justify-content: center;
            padding-left: 0;
            padding-right: 0;
        }

        .sidebar-icon {
            width: 38px;
            height: 38px;
            border-radius: 14px;
            background: var(--vale-teal);
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            overflow: hidden;
            font-weight: 800;
            font-size: 17px;
        }

        .sidebar-icon img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .sidebar-text {
            display: none;
            flex-direction: column;
            text-align: left;
            line-height: 1.15;
            white-space: nowrap;
            min-width: 0;
        }

        .sidebar.open .sidebar-text {
            display: flex;
        }

        .sidebar-title {
            font-size: 13px;
            font-weight: 800;
            color: var(--vale-teal-dark);
        }

        .sidebar-subtitle {
            font-size: 11px;
            color: var(--text-muted);
            margin-top: 2px;
        }

        .sidebar-footer .sidebar-icon {
            background: #C62828;
        }

        /* =====================================================
           LINKS ÚTEIS
        ===================================================== */

        .external-group {
            width: 100%;
        }

        .external-toggle {
            position: relative;
        }

        .external-arrow {
            display: none;
            margin-left: auto;
            color: var(--vale-teal-dark);
            font-size: 15px;
            transition: transform 0.2s ease-in-out;
        }

        .sidebar.open .external-arrow {
            display: inline-block;
        }

        .external-group.open .external-arrow {
            transform: rotate(180deg);
        }

        .external-icon-stack {
            position: relative;
            background: #ffffff;
        }

        .external-icon-stack img {
            width: 25px;
            height: 25px;
            border-radius: 9px;
            object-fit: cover;
            position: absolute;
            border: 2px solid #ffffff;
            box-shadow: 0 2px 5px rgba(0,0,0,0.18);
        }

        .external-icon-stack img:nth-child(1) {
            left: 3px;
            top: 6px;
        }

        .external-icon-stack img:nth-child(2) {
            right: 3px;
            top: 6px;
        }

        .external-icon-stack img:nth-child(3) {
            left: 7px;
            bottom: 3px;
        }

        .external-list {
            display: none;
            flex-direction: column;
            gap: 8px;
            margin-top: 8px;
            padding-left: 4px;
            padding-right: 4px;
        }

        .sidebar.open .external-group.open .external-list {
            display: flex;
        }

        .sidebar:not(.open) .external-list {
            display: none !important;
        }

        .external-link {
            min-height: 48px;
            border-radius: 15px;
            background: rgba(255, 255, 255, 0.68);
            color: var(--text-dark);
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 11px;
            padding: 0 12px;
            transition: all 0.2s ease-in-out;
        }

        .external-link:hover {
            background: #ffffff;
            transform: translateX(4px);
        }

        .external-link img,
        .external-logo {
            width: 31px;
            height: 31px;
            border-radius: 11px;
            object-fit: cover;
            flex-shrink: 0;
        }

        .external-logo-fallback {
            background: var(--vale-teal);
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 800;
        }

        .external-link span {
            font-size: 13px;
            font-weight: 800;
            color: var(--vale-teal-dark);
        }
        :root {
            --vale-teal: #00857A;
            --vale-teal-dark: #006B63;
            --vale-yellow: #F6B800;
            --text-dark: #263238;
            --text-muted: #607D8B;
            --card-bg: rgba(255,255,255,0.92);
            --sidebar-collapsed: 88px;
            --sidebar-expanded: 292px;
        }

        * {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            width: 100%;
            min-height: 100vh;
            min-height: 100dvh;
            overflow-x: hidden;
        }

        body {
            font-family: Arial, Helvetica, sans-serif;
            background:
                linear-gradient(rgba(0,60,55,0.15), rgba(0,60,55,0.15)),
                url("libs/img/login_bg.jpg") center center / cover no-repeat fixed;
            color: var(--text-dark);
        }

        .sidebar {
    position: fixed;
    left: 0;
    top: 0;
    width: 88px;
    height: 100vh;
    background: rgba(255,255,255,0.85);
    backdrop-filter: blur(10px);
    display: flex;
    flex-direction: column;
    padding: 12px;
    z-index: 100;
    transition: width 0.3s;
}

.sidebar.open {
    width: 260px;
}

.sidebar-header {
    text-align: center;
}

.sidebar-logo-area img {
    width: 50px;
}

.sidebar-nav {
    display: flex;
    flex-direction: column;
    gap: 10px;
}

.sidebar-link {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px;
    border-radius: 10px;
    text-decoration: none;
    color: #333;
}

.sidebar-link:hover {
    background: #fff;
}

.sidebar-icon {
    width: 35px;
    height: 35px;
    background: #00857A;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 10px;
}

.sidebar-text {
    display: none;
}

.sidebar.open .sidebar-text {
    display: block;
}

        .main-content {
            min-height: 100vh;
            margin-left: var(--sidebar-collapsed);
            padding: 26px 42px 44px 42px;
            transition: margin-left 0.25s ease-in-out;
        }

        .main-content {
    margin-left: 88px;
}

body.sidebar-open .main-content {
    margin-left: 260px;
}

        .page-header {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto 16px auto;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 12px;
            text-align: center;
        }

        .page-title {
            margin: 0;
            width: 100%;
            text-align: center;
            color: #ffffff;
            font-size: clamp(34px, 4.2vw, 58px);
            line-height: 1.05;
            font-weight: 900;
            letter-spacing: 3px;
            text-transform: uppercase;
            text-shadow: 0 4px 12px rgba(0, 0, 0, 0.48);
        }

        .user-info {
            display: inline-block;
            padding: 7px 18px;
            border-radius: 999px;
            background: rgba(0, 107, 99, 0.62);
            color: #ffffff;
            font-size: 13px;
            line-height: 1.35;
            text-shadow: 0 2px 7px rgba(0, 0, 0, 0.35);
            backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
        }

        .summary-bar {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto 18px auto;
            display: flex;
            justify-content: center;
            gap: 14px;
            flex-wrap: wrap;
        }

        .summary-pill {
            padding: 10px 16px;
            border-radius: 999px;
            background: rgba(255,255,255,0.82);
            color: var(--vale-teal-dark);
            font-size: 13px;
            font-weight: 800;
            box-shadow: 0 8px 18px rgba(0,0,0,0.12);
        }

        /* Busca com lupa interativa no canto superior direito */
        .search-holder {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto 18px auto;
            display: flex;
            justify-content: flex-end;
            position: relative;
        }

        .search-shell {
            width: 58px;
            max-width: 100%;
            overflow: hidden;
            transition: width 0.28s ease-in-out;
            border-radius: 999px;
            background: rgba(255,255,255,0.92);
            box-shadow: 0 12px 24px rgba(0,0,0,0.16);
        }

        .search-shell:hover,
        .search-shell:focus-within,
        .search-shell.open {
            width: min(96vw, 980px);
        }

        .search-form {
            display: flex;
            align-items: center;
            gap: 10px;
            min-height: 58px;
            padding: 8px 10px;
            white-space: nowrap;
        }

        .search-button-icon {
            width: 42px;
            height: 42px;
            min-width: 42px;
            border: none;
            border-radius: 50%;
            background: var(--vale-teal);
            color: #ffffff;
            font-size: 18px;
            font-weight: 800;
            cursor: pointer;
        }

        .search-form input,
        .search-form select {
            height: 40px;
            padding: 0 12px;
            border: none;
            border-radius: 10px;
            background: #f4f7f8;
            color: var(--text-dark);
            font-size: 14px;
            outline: none;
        }

        .search-form input[name="busca"] {
            flex: 1.3;
            min-width: 160px;
        }

        .search-form input[name="nome"] {
            flex: 1;
            min-width: 140px;
        }

        .search-form input[name="data_txt"] {
            width: 170px;
        }

        .search-form select {
            width: 170px;
        }

        .search-submit {
            height: 40px;
            padding: 0 16px;
            border: none;
            border-radius: 10px;
            background: var(--vale-teal);
            color: #ffffff;
            font-weight: 800;
            cursor: pointer;
        }

        .records-list {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 18px;
        }

        .record-card {
            display: grid;
            grid-template-columns: 190px 1fr;
            gap: 24px;
            padding: 22px;
            border-radius: 24px;
            background: var(--card-bg);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            box-shadow: 0 14px 30px rgba(0,0,0,0.18);
            transition: transform .2s ease, box-shadow .2s ease;
        }

        .record-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 20px 38px rgba(0,0,0,0.22);
        }

        .record-card.alerta {
            border-left: 6px solid #E53935;
        }

        .record-card.pendente {
            border-top: 4px solid var(--vale-yellow);
        }

        .record-media {
            width: 190px;
            height: 140px;
            border-radius: 16px;
            overflow: hidden;
            background: #E4F2F0;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: inset 0 0 0 1px rgba(0,133,122,.18);
        }

        .record-media img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .doc-thumb {
            width: 100%;
            height: 100%;
            background: linear-gradient(135deg, var(--vale-teal), var(--vale-teal-dark));
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-direction: column;
            font-weight: 800;
            gap: 4px;
        }

        .doc-thumb .doc-type {
            font-size: 16px;
        }

        .record-content {
            min-width: 0;
        }

        .record-top {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 18px;
            margin-bottom: 8px;
        }

        .record-title {
            margin: 0;
            font-size: clamp(20px, 1.8vw, 28px);
            line-height: 1.18;
            color: var(--vale-teal-dark);
            font-weight: 900;
        }

        .record-badges {
            display: flex;
            gap: 8px;
            align-items: center;
            flex-wrap: wrap;
            justify-content: flex-end;
        }

        .badge {
            padding: 7px 12px;
            border-radius: 999px;
            color: #ffffff;
            font-size: 13px;
            font-weight: 900;
            white-space: nowrap;
        }

        .badge-cod {
            background: var(--vale-teal);
        }

        .badge-status-pendente {
            background: #F6B800;
            color: #263238;
        }

        .badge-status-visualizado {
            background: #43A047;
        }

        .record-meta {
            margin-bottom: 10px;
            font-size: 13px;
            color: var(--text-muted);
            font-weight: 700;
        }

        .record-text {
            margin: 0 0 14px 0;
            font-size: 15px;
            line-height: 1.5;
            color: #37474F;
        }

        details.record-details {
            margin-top: 8px;
            border-top: 1px solid rgba(0,0,0,.08);
            padding-top: 10px;
        }

        details.record-details summary {
            cursor: pointer;
            font-weight: 800;
            color: var(--vale-teal-dark);
            list-style: none;
            outline: none;
        }

        details.record-details summary::-webkit-details-marker {
            display: none;
        }

        .details-body {
            margin-top: 12px;
            display: grid;
            grid-template-columns: 1fr;
            gap: 12px;
        }

        .full-text {
            padding: 14px;
            border-radius: 14px;
            background: rgba(244,247,248,0.92);
            line-height: 1.55;
            white-space: pre-wrap;
        }

        .annex-actions {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .btn,
        .btn-secondary {
            display: inline-block;
            padding: 10px 16px;
            border-radius: 10px;
            text-decoration: none;
            font-size: 14px;
            font-weight: 800;
            cursor: pointer;
            border: none;
        }

        .btn {
            background: var(--vale-teal);
            color: #ffffff;
        }

        .btn-secondary {
            background: #ECEFF1;
            color: #263238;
        }

        .confirm-form {
            margin-top: 4px;
        }

        .empty-state {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto;
            padding: 28px;
            border-radius: 20px;
            background: var(--card-bg);
            text-align: center;
            font-weight: 800;
        }

        .pagination {
            width: 100%;
            max-width: 1240px;
            margin: 26px auto 0 auto;
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
        }

        .page-link,
        .page-current,
        .page-dots {
            min-width: 40px;
            height: 40px;
            border-radius: 999px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 0 12px;
            font-size: 14px;
            font-weight: 800;
            text-decoration: none;
        }

        .page-link {
            background: rgba(255,255,255,0.88);
            color: var(--vale-teal-dark);
            box-shadow: 0 8px 18px rgba(0,0,0,0.12);
        }

        .page-link:hover {
            background: #ffffff;
        }

        .page-current {
            background: var(--vale-teal);
            color: #ffffff;
        }

        .page-dots {
            background: transparent;
            color: #ffffff;
            min-width: auto;
            padding: 0 4px;
        }

        @media (max-width: 1180px) {
            .main-content {
                padding: 22px 24px 34px 24px;
            }

            .record-card {
                grid-template-columns: 1fr;
            }

            .record-media {
                width: 100%;
                height: 220px;
            }

            .search-shell:hover,
            .search-shell:focus-within,
            .search-shell.open {
                width: min(96vw, 940px);
            }

            .search-form {
                flex-wrap: wrap;
                min-height: auto;
            }

            .search-form input[name="busca"],
            .search-form input[name="nome"],
            .search-form input[name="data_txt"],
            .search-form select {
                width: 100%;
                flex: 1 1 220px;
            }
        }

        @media (max-width: 760px) {
            .main-content {
                padding: 18px 16px 28px 16px;
            }

            .page-title {
                font-size: clamp(26px, 8vw, 38px);
                letter-spacing: 1.5px;
            }

            .search-shell {
                width: 58px;
            }

            .search-shell:hover,
            .search-shell:focus-within,
            .search-shell.open {
                width: min(96vw, 96vw);
            }

            .record-media {
                height: 190px;
            }
        }

    /* ==============================================
   BARRA SUPERIOR (IGUAL PROTÓTIPO)
============================================== */

.top-controls {
    max-width: 1240px;
    margin: 20px auto;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 20px;
}

/* FILTROS */
.filters-box {
    display: flex;
    align-items: center;
    gap: 8px;
    background: rgba(255,255,255,0.7);
    padding: 8px 12px;
    border-radius: 20px;
}

.filters-title {
    font-weight: 800;
    color: #444;
}

.filters-box input {
    padding: 6px 8px;
    border-radius: 10px;
    border: none;
    background: #fff;
}

.filters-box button {
    background: #00857A;
    color: white;
    border: none;
    padding: 6px 10px;
    border-radius: 10px;
}

/* PAGINAÇÃO CENTRAL */
.top-pagination {
    color: white;
    font-weight: 800;
}

/* BUSCA DIREITA */
.search-box input {
    padding: 8px 14px;
    border-radius: 20px;
    border: none;
    background: rgba(255,255,255,0.8);
    width: 180px;
}
.sidebar {
    position: fixed;
    left: 0;
    top: 0;
    width: var(--sidebar-collapsed);
    height: 100vh;
    background: rgba(255,255,255,0.85);
    z-index: 100;
}
.sidebar-text {
    display: none;
}

.sidebar.open .sidebar-text {
    display: flex;
}
.sidebar:not(.open) .sidebar-link {
    justify-content: center;
}
.sidebar-text {
    display: none !important;
}

.sidebar.open .sidebar-text {
    display: flex !important;
}

.sidebar:not(.open) .sidebar-link {
    justify-content: center !important;
}

/* =====================================================
   ÍCONE LINKS ÚTEIS - MODELO PASTINHAS SOBREPOSTAS
   Visual: cards/logos uma na frente da outra
===================================================== */

/* Remove interferências do modelo antigo */
.external-icon-stack,
.external-icon-stack img,
.external-icon-stack::before,
.external-icon-stack::after {
    background-image: none !important;
}

/* Container principal do ícone de Links úteis */
.external-icon-stack-folder {
    width: 42px !important;
    height: 42px !important;
    min-width: 42px !important;
    min-height: 42px !important;
    max-width: 42px !important;
    max-height: 42px !important;

    position: relative !important;
    display: block !important;

    background: transparent !important;
    border-radius: 14px !important;
    overflow: visible !important;

    padding: 0 !important;
    flex-shrink: 0 !important;
}

/* Cada “pastinha/card” */
.external-icon-stack-folder .folder-card {
    position: absolute !important;

    width: 29px !important;
    height: 29px !important;

    border-radius: 10px !important;
    background: #ffffff !important;

    display: flex !important;
    align-items: center !important;
    justify-content: center !important;

    box-shadow: 0 3px 8px rgba(0, 0, 0, 0.22) !important;
    border: 2px solid rgba(255, 255, 255, 0.95) !important;

    overflow: hidden !important;
}

/* Imagem dentro de cada card */
.external-icon-stack-folder .folder-card img {
    width: 100% !important;
    height: 100% !important;
    object-fit: cover !important;

    position: static !important;
    display: block !important;

    border: none !important;
    border-radius: 8px !important;
    box-shadow: none !important;

    margin: 0 !important;
    padding: 0 !important;
}

/* Card de trás */
.external-icon-stack-folder .folder-card-back {
    left: 1px !important;
    top: 5px !important;
    z-index: 1 !important;
    opacity: 0.92 !important;
    transform: rotate(-8deg) scale(0.92) !important;
}

/* Card do meio */
.external-icon-stack-folder .folder-card-middle {
    left: 7px !important;
    top: 2px !important;
    z-index: 2 !important;
    opacity: 0.96 !important;
    transform: rotate(5deg) scale(0.96) !important;
}

/* Card da frente */
.external-icon-stack-folder .folder-card-front {
    left: 11px !important;
    top: 10px !important;
    z-index: 3 !important;
    transform: rotate(0deg) scale(1) !important;
}

/* Efeito suave ao passar o mouse no botão Links úteis */
.external-toggle:hover .external-icon-stack-folder .folder-card-back {
    transform: rotate(-10deg) translateX(-2px) scale(0.92) !important;
}

.external-toggle:hover .external-icon-stack-folder .folder-card-middle {
    transform: rotate(6deg) translateY(-2px) scale(0.96) !important;
}

.external-toggle:hover .external-icon-stack-folder .folder-card-front {
    transform: translateX(2px) translateY(1px) scale(1.02) !important;
}

/* =====================================================
   AJUSTE RESPONSIVO FINAL - VISUALIZAR REGISTROS
   Corrige adaptação para diferentes tamanhos de tela
===================================================== */

/* Estrutura principal da página */
.main-content {
    margin-left: var(--sidebar-collapsed) !important;
    width: calc(100vw - var(--sidebar-collapsed)) !important;
    max-width: none !important;
    min-height: 100vh !important;
    padding: clamp(20px, 2.8vw, 40px) !important;
    transition: margin-left 0.25s ease-in-out, width 0.25s ease-in-out;
}

body.sidebar-open .main-content {
    margin-left: var(--sidebar-expanded) !important;
    width: calc(100vw - var(--sidebar-expanded)) !important;
}

/* Conteúdo central com largura fluida */
.page-header,
.summary-bar,
.top-controls,
.records-list,
.pagination {
    width: min(100%, 1400px) !important;
    max-width: none !important;
    margin-left: auto !important;
    margin-right: auto !important;
}

/* Título */
.page-title {
    font-size: clamp(30px, 4vw, 72px) !important;
    line-height: 1.05 !important;
    letter-spacing: clamp(1px, 0.35vw, 3px) !important;
    text-align: center !important;
    word-break: break-word !important;
}

/* Barra superior com comportamento flexível */
.top-controls {
    display: flex !important;
    flex-wrap: wrap !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 14px !important;
}

/* Filtros */
.filters-box {
    display: flex !important;
    flex-wrap: wrap !important;
    align-items: center !important;
    gap: 8px !important;
    flex: 1 1 560px !important;
    min-width: 280px !important;
}

.filters-box input {
    flex: 1 1 140px !important;
    min-width: 120px !important;
}

.filters-box button {
    flex: 0 0 auto !important;
}

/* Paginação central */
.top-pagination {
    flex: 0 0 auto !important;
    white-space: nowrap !important;
}

/* Busca */
.search-box {
    flex: 0 1 220px !important;
    display: flex !important;
    justify-content: flex-end !important;
}

.search-box input {
    width: 100% !important;
    min-width: 150px !important;
    max-width: 220px !important;
}

/* Cards */
.record-card {
    display: grid !important;
    grid-template-columns: 190px 1fr !important;
    gap: 22px !important;
    width: 100% !important;
}

.record-media {
    width: 190px !important;
    height: 140px !important;
}

.record-content {
    min-width: 0 !important;
}

.record-title {
    font-size: clamp(20px, 2vw, 30px) !important;
    line-height: 1.2 !important;
}

.record-text {
    font-size: clamp(14px, 1.2vw, 16px) !important;
    line-height: 1.55 !important;
}

/* Ajuste da paginação */
.pagination {
    display: flex !important;
    justify-content: center !important;
    flex-wrap: wrap !important;
    gap: 8px !important;
}

/* =========================
   NOTEBOOK / TELAS MÉDIAS
========================= */
@media (max-width: 1180px) {

    /* Menu aberto não deve empurrar demais o conteúdo */
    body.sidebar-open .main-content {
        margin-left: var(--sidebar-collapsed) !important;
        width: calc(100vw - var(--sidebar-collapsed)) !important;
    }

    .page-title {
        font-size: clamp(28px, 5vw, 52px) !important;
    }

    .top-controls {
        flex-direction: column !important;
        align-items: stretch !important;
    }

    .filters-box,
    .search-box,
    .top-pagination {
        width: 100% !important;
        max-width: 100% !important;
    }

    .search-box {
        justify-content: flex-start !important;
    }

    .search-box input {
        max-width: none !important;
    }

    .record-card {
        grid-template-columns: 160px 1fr !important;
    }

    .record-media {
        width: 160px !important;
        height: 125px !important;
    }
}

/* =========================
   TABLET / TELAS PEQUENAS
========================= */
@media (max-width: 900px) {

    .main-content {
        padding: 18px 18px 28px 18px !important;
    }

    .record-card {
        grid-template-columns: 1fr !important;
    }

    .record-media {
        width: 100% !important;
        height: 220px !important;
    }

    .record-top {
        flex-direction: column !important;
        align-items: flex-start !important;
    }

    .record-badges {
        justify-content: flex-start !important;
    }
}

/* =========================
   TELAS MENORES / NOTEBOOK COM JANELA REDUZIDA
========================= */
@media (max-width: 760px) {

    :root {
        --sidebar-collapsed: 76px;
        --sidebar-expanded: 240px;
    }

    .main-content {
        margin-left: var(--sidebar-collapsed) !important;
        width: calc(100vw - var(--sidebar-collapsed)) !important;
        padding: 16px !important;
    }

    body.sidebar-open .main-content {
        margin-left: var(--sidebar-collapsed) !important;
        width: calc(100vw - var(--sidebar-collapsed)) !important;
    }

    .page-title {
        font-size: clamp(24px, 8vw, 40px) !important;
        letter-spacing: 1px !important;
    }

    .summary-bar {
        gap: 10px !important;
    }

    .summary-pill {
        width: 100% !important;
        text-align: center !important;
    }

    .filters-box {
        padding: 10px !important;
    }

    .filters-box input,
    .filters-box button {
        width: 100% !important;
    }

    .search-box input {
        width: 100% !important;
    }
}

/* =====================================================
   CORREÇÃO DO BLOCO DE FILTROS
   Impede o botão "Aplicar" de descer em telas médias
===================================================== */

/* Barra superior */
.top-controls {
    max-width: 1240px;
    margin: 20px auto;
    display: grid !important;
    grid-template-columns: minmax(680px, 1fr) auto 220px;
    align-items: center;
    gap: 18px;
}

/* Filtros */
.filters-box {
    display: grid !important;
    grid-template-columns: auto minmax(140px, 1fr) minmax(140px, 1fr) minmax(120px, 1fr) auto;
    align-items: center;
    gap: 10px;
    background: rgba(255,255,255,0.72);
    padding: 10px 14px;
    border-radius: 22px;
    min-width: 0;
    width: 100%;
}

/* título "Filtros" */
.filters-title {
    font-weight: 800;
    color: #444;
    white-space: nowrap;
}

/* campos */
.filters-box input {
    width: 100%;
    min-width: 0;
    height: 42px;
    padding: 0 12px;
    border-radius: 12px;
    border: none;
    background: #fff;
    font-size: 14px;
}

/* botão */
.filters-box button {
    height: 42px;
    padding: 0 18px;
    border: none;
    border-radius: 12px;
    background: #00857A;
    color: #ffffff;
    font-weight: 800;
    white-space: nowrap;
    cursor: pointer;
}

/* paginação central */
.top-pagination {
    color: #ffffff;
    font-weight: 800;
    white-space: nowrap;
    justify-self: center;
}

/* busca à direita */
.search-box {
    justify-self: end;
    width: 100%;
    max-width: 220px;
}

.search-box input {
    width: 100%;
    height: 42px;
    padding: 0 14px;
    border-radius: 20px;
    border: none;
    background: rgba(255,255,255,0.82);
    font-size: 14px;
}

/* =========================
   Notebook / telas médias
========================= */
@media (max-width: 1280px) {
    .top-controls {
        grid-template-columns: 1fr auto 220px;
    }

    .filters-box {
        grid-template-columns: auto 1fr 1fr 1fr auto;
    }
}

/* =========================
   Quando a tela ficar menor,
   a quebra passa a ser controlada
========================= */
@media (max-width: 1080px) {
    .top-controls {
        grid-template-columns: 1fr;
        gap: 14px;
    }

    .top-pagination {
        justify-self: start;
    }

    .search-box {
        justify-self: start;
        max-width: 260px;
    }
}

/* =========================
   Tablet e telas menores
========================= */
@media (max-width: 820px) {
    .filters-box {
        grid-template-columns: 1fr 1fr;
    }

    .filters-title {
        grid-column: 1 / -1;
    }

    .filters-box input[name="nome"] {
        grid-column: 1 / 2;
    }

    .filters-box input[name="data_txt"] {
        grid-column: 2 / 3;
    }

    .filters-box input[name="codigo"] {
        grid-column: 1 / 2;
    }

    .filters-box button {
        grid-column: 2 / 3;
        width: 100%;
    }
}

/* =========================
   Celulares / largura bem pequena
========================= */
@media (max-width: 560px) {
    .filters-box {
        grid-template-columns: 1fr;
    }

    .filters-title,
    .filters-box input,
    .filters-box button {
        grid-column: 1 / -1;
        width: 100%;
    }

    .search-box {
        max-width: 100%;
    }
}
    </style>
</head>
<body>
<aside id="sidebar" class="sidebar">

    <div class="sidebar-header">
        <button type="button" class="menu-toggle" onclick="toggleSidebar()">☰</button>

        <div class="sidebar-logo-area">
            <img src="libs/img/logo-vale.png">
        </div>
    </div>

    <nav class="sidebar-nav">

        <a href="form_home.asp" class="sidebar-link">
            <span class="sidebar-icon">
                <img src="libs/img/img_link/Home.jpg">
            </span>
            <span class="sidebar-text">
                <span class="sidebar-title">Home</span>
            </span>
        </a>

        <% If VAR_USUARIO_FUNCAO = "SUPERVISOR" Or VAR_USUARIO_NIVEL = "ADM" Then %>
        <a href="form_formulario.asp" class="sidebar-link">
            <span class="sidebar-icon">+</span>
            <span class="sidebar-text">
                <span class="sidebar-title">Novo registro</span>
            </span>
        </a>
        <% End If %>

        <% If VAR_USUARIO_FUNCAO = "INSPETOR" Or VAR_USUARIO_FUNCAO = "SUPERVISOR" Or VAR_USUARIO_NIVEL = "ADM" Then %>
        <a href="form_grafico.asp" class="sidebar-link">
            <span class="sidebar-icon">
                <img src="libs/img/img_link/relatorios.jpg">
            </span>
            <span class="sidebar-text">
                <span class="sidebar-title">Relatórios</span>
            </span>
        </a>
        <% End If %>

        <a href="https://efvmworkplace/central de kaizen/index.php" class="sidebar-link" target="_blank">
            <span class="sidebar-icon">
                <img src="libs/img/img_link/kaizen.jpg">
            </span>
            <span class="sidebar-text">
                <span class="sidebar-title">Kaizen</span>
            </span>
        </a>

        <a href="https://linktr.ee/FaleComVale" class="sidebar-link" target="_blank">
            <span class="sidebar-icon">#</span>
            <span class="sidebar-text">
                <span class="sidebar-title">#FaleCOM</span>
            </span>
        </a>

        <a href="dss_online.asp" class="sidebar-link">
            <span class="sidebar-icon">D</span>
            <span class="sidebar-text">
                <span class="sidebar-title">DSS Online</span>
            </span>
        </a>
 <div id="linksUteisGroup" class="external-group">

    <button type="button" class="external-toggle" onclick="toggleLinksUteis()">
    <span class="sidebar-icon external-icon-stack-folder">
        <span class="folder-card folder-card-back">
            <img src="libs/img/img_link/gdb.jpg" alt="GDB">
        </span>

        <span class="folder-card folder-card-middle">
            <img src="libs/img/img_link/edados.jpg" alt="E-Dados">
        </span>

        <span class="folder-card folder-card-front">
            <img src="libs/img/img_link/irislogo.png" alt="Iris">
        </span>
    </span>

    <span class="sidebar-text">
        <span class="sidebar-title">Links úteis</span>
        <span class="sidebar-subtitle">Sistemas externos</span>
    </span>

    <span class="external-arrow">▾</span>
</button>


    <div class="external-list">

        <a href="https://equipfer.valenet.valeglobal.net/equipfer/" class="external-link" target="_blank">
            <img src="libs/img/img_link/tooplate_image_01.jpg" alt="Equipfer">
            <span>Equipfer</span>
        </a>

        <a href="https://gdb.valeglobal.net/gdb/view/login/login.faces" class="external-link" target="_blank">
            <img src="libs/img/img_link/gdb.jpg" alt="GDB">
            <span>GDB</span>
        </a>

        <a href="https://performancemanager4.successfactors.com/sf/home" class="external-link" target="_blank">
            <img src="libs/img/img_link/prontidao1.jpg" alt="VES">
            <span>VES</span>
        </a>

        <a href="https://lbrportalfolha.valenet.valeglobal.net/portalrh/Produtos/SAAA/Principal2.aspx?amb_selecionado=0&abrir_nova_janela=N&eh_mdesigner=N&nome_portal=616653596455764672655738516965596E57664E38413D3D" class="external-link" target="_blank">
            <img src="libs/img/img_link/edados.jpg" alt="E-Dados">
            <span>E-Dados</span>
        </a>

        <a href="https://iris.valeglobal.net/login" class="external-link" target="_blank">
            <img src="libs/img/img_link/irislogo.png" alt="Iris">
            <span>Iris</span>
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
         COLE AQUI O MESMO BLOCO <aside>...</aside> DA HOME
         (sem alterar nada, para manter o mesmo menu lateral)
         ===================================================== -->

    <main class="main-content">

        <header class="page-header">
            <h1 class="page-title">Visualizar Registros</h1>
            <div class="user-info">
                Bem-vindo(a), <strong><%=HtmlSafe(usuarioNome)%></strong>
                &nbsp;·&nbsp;
                <%=HtmlSafe(usuarioFuncao)%>
            </div>
        </header>

        <div class="summary-bar">
            <div class="summary-pill">Total encontrado: <%=totalRegistros%></div>
            <div class="summary-pill">Página <%=pagina%> de <%=totalPaginas%></div>
            <div class="summary-pill">Filtros ativos: 
                <% 
                    filtrosAtivos = 0
                    If busca <> "" Then filtrosAtivos = filtrosAtivos + 1
                    If nome <> "" Then filtrosAtivos = filtrosAtivos + 1
                    If tipo <> "" Then filtrosAtivos = filtrosAtivos + 1
                    If dataTxt <> "" Then filtrosAtivos = filtrosAtivos + 1
                    If statusFiltro <> "" Then filtrosAtivos = filtrosAtivos + 1
                    Response.Write filtrosAtivos
                %>
            </div>
        </div>

        <!-- Busca com lupa interativa no canto superior direito -->
        <div class="top-controls">

    <!-- FILTROS (ESQUERDA) -->
    <form class="filters-box" method="get">

        <span class="filters-title">Filtros</span>

        <input type="text" name="nome" placeholder="Nome" value="<%=HtmlSafe(nome)%>">
        <input type="text" name="data_txt" placeholder="Data" value="<%=HtmlSafe(dataTxt)%>">
        <input type="text" name="codigo" placeholder="Código">

        <button type="submit">Aplicar</button>
    </form>

    <!-- PAGINAÇÃO CENTRAL -->
    <div class="top-pagination">
        Página <strong><%=pagina%></strong>
    </div>

    <!-- BUSCA DIREITA -->
    <form class="search-box" method="get">
        <input type="text" name="busca" placeholder="Pesquisar..." value="<%=HtmlSafe(busca)%>">
    </form>

</div>

        <div class="records-list">
            <%
            If Not RS.EOF Then
                Do Until RS.EOF

                    cod = RS("COD")
                    titulo = "" & RS("TITULO")
                    tipoR = "" & RS("TIPO")
                    dataR = "" & RS("DATA")
                    info = "" & RS("INFORMACAO")
                    img = TratarImagem(RS("IMG_ID"))
                    autor = "" & RS("AUTOR")
                    statusL = "" & RS("STATUS_LEITURA")
                    dataVisual = "" & RS("DATA_VISUALIZACAO")

                    classeTipo = ""
                    If UCase(tipoR) = "ALERTA" Then classeTipo = "alerta"

                    classeStatus = ""
                    If UCase(statusL) = "NÃO" Then
                        classeStatus = "pendente"
                    End If
            %>

            <article class="record-card <%=classeTipo%> <%=classeStatus%>">
                <div class="record-media">
                    <% If EhImagem(img) Then %>
                        <img src="<%=img%>" alt="Imagem do registro <%=cod%>" onerror="this.onerror=null;this.src='NOIMG.jpg';">
                    <% Else %>
                        <div class="doc-thumb">
                            <span>📄</span>
                            <span class="doc-type"><%=TipoDocumento(img)%></span>
                        </div>
                    <% End If %>
                </div>

                <div class="record-content">
                    <div class="record-top">
                        <h3 class="record-title"><%=TextoCurto(titulo, 130)%></h3>

                        <div class="record-badges">
                            <span class="badge badge-cod">#<%=cod%></span>

                            <% If UCase(statusL) = "NÃO" Then %>
                                <span class="badge badge-status-pendente">Pendente</span>
                            <% Else %>
                                <span class="badge badge-status-visualizado">Visualizado</span>
                            <% End If %>
                        </div>
                    </div>

                    <div class="record-meta">
                        Autor: <%=HtmlSafe(autor)%>
                        &nbsp;·&nbsp;
                        Tipo: <%=HtmlSafe(tipoR)%>
                        &nbsp;·&nbsp;
                        Data: <%=HtmlSafe(dataR)%>
                        <% If UCase(statusL) = "SIM" And dataVisual <> "" Then %>
                            &nbsp;·&nbsp;
                            Visualizado em: <%=HtmlSafe(dataVisual)%>
                        <% End If %>
                    </div>

                    <p class="record-text"><%=TextoCurto(info, 320)%></p>

                    <details class="record-details">
                        <summary>Visualizar em tela cheia</summary>

                        <div class="details-body">
                            <div class="full-text"><%=HtmlSafe(info)%></div>

                            <div class="annex-actions">
                                <% If EhImagem(img) Then %>
                                    <a class="btn-secondary" href="<%=img%>" target="_blank">Abrir imagem</a>
                                <% ElseIf UCase(img) <> "NOIMG.JPG" And UCase(img) <> "NOIMG" Then %>
                                    <a class="btn-secondary" href="<%=img%>" target="_blank">Abrir anexo</a>
                                <% End If %>

                                <% If UCase(statusL) = "NÃO" Then %>
                                    <form class="confirm-form" method="post" action="valida_visualizar_registro.asp" style="display:inline;">
                                        <input type="hidden" name="inputCOD" value="<%=cod%>">
                                        <button type="submit" class="btn">Confirmar visualização</button>
                                    </form>
                                <% End If %>
                            </div>
                        </div>
                    </details>
                </div>
            </article>

            <%
                    RS.MoveNext
                Loop
            Else
            %>
                <div class="empty-state">
                    Nenhum registro encontrado para os filtros aplicados.
                </div>
            <%
            End If
            %>
        </div>

        <!-- Paginação estilo 1 ... 2 ... 3 ... (...) ... 10 -->
        <div class="pagination">
            <%
            If totalPaginas <= 1 Then
                Response.Write "<span class='page-current'>1</span>"
            Else
                ' Link anterior
                If pagina > 1 Then
                    Response.Write "<a class='page-link' href='" & BuildPageUrl(pagina-1, busca, nome, tipo, dataTxt, statusFiltro) & "'>&laquo;</a>"
                End If

                ' Sempre mostra página 1
                If pagina = 1 Then
                    Response.Write "<span class='page-current'>1</span>"
                Else
                    Response.Write "<a class='page-link' href='" & BuildPageUrl(1, busca, nome, tipo, dataTxt, statusFiltro) & "'>1</a>"
                End If

                startPage = pagina - 2
                endPage = pagina + 2

                If startPage < 2 Then startPage = 2
                If endPage > totalPaginas - 1 Then endPage = totalPaginas - 1

                If startPage > 2 Then
                    Response.Write "<span class='page-dots'>...</span>"
                End If

                For i = startPage To endPage
                    If i = pagina Then
                        Response.Write "<span class='page-current'>" & i & "</span>"
                    Else
                        Response.Write "<a class='page-link' href='" & BuildPageUrl(i, busca, nome, tipo, dataTxt, statusFiltro) & "'>" & i & "</a>"
                    End If
                Next

                If endPage < totalPaginas - 1 Then
                    Response.Write "<span class='page-dots'>...</span>"
                End If

                ' Sempre mostra última página se houver mais de 1
                If totalPaginas > 1 Then
                    If pagina = totalPaginas Then
                        Response.Write "<span class='page-current'>" & totalPaginas & "</span>"
                    Else
                        Response.Write "<a class='page-link' href='" & BuildPageUrl(totalPaginas, busca, nome, tipo, dataTxt, statusFiltro) & "'>" & totalPaginas & "</a>"
                    End If
                End If

                ' Link próximo
                If pagina < totalPaginas Then
                    Response.Write "<a class='page-link' href='" & BuildPageUrl(pagina+1, busca, nome, tipo, dataTxt, statusFiltro) & "'>&raquo;</a>"
                End If
            End If
            %>
        </div>

    </main>

    <script>
        function toggleBusca() {
            document.getElementById("searchShell").classList.toggle("open");
        }

function toggleSidebar() {
    var sidebar = document.getElementById("sidebar");
    var body = document.body;

    if (sidebar.classList.contains("open")) {
        sidebar.classList.remove("open");
        body.classList.remove("sidebar-open");
    } else {
        sidebar.classList.add("open");
        body.classList.add("sidebar-open");
    }
}
function toggleLinksUteis() {
            var sidebar = document.getElementById("sidebar");
            var group = document.getElementById("linksUteisGroup");

            if (!sidebar.classList.contains("open")) {
                sidebar.classList.add("open");
                document.body.classList.add("sidebar-open");
                localStorage.setItem("centralSidebarOpen", "S");
            }

            if (group.classList.contains("open")) {
                group.classList.remove("open");
                localStorage.setItem("centralLinksUteisOpen", "N");
            } else {
                group.classList.add("open");
                localStorage.setItem("centralLinksUteisOpen", "S");
            }
        }
    </script>

 <script>
        function toggleSidebar() {
           var sidebar = document.getElementById("sidebar");
           var body = document.body;

        if (sidebar.classList.contains("open")) {
          sidebar.classList.remove("open");
          body.classList.remove("sidebar-open");
     } else {
        sidebar.classList.add("open");
        body.classList.add("sidebar-open");
    }
}
</script>

<Script>
function toggleSidebar() {
    var sidebar = document.getElementById("sidebar");
    var body = document.body;

    if (sidebar.classList.contains("open")) {
        sidebar.classList.remove("open");
        body.classList.remove("sidebar-open");
        localStorage.setItem("centralSidebarOpen", "N");
    } else {
        sidebar.classList.add("open");
        body.classList.add("sidebar-open");
        localStorage.setItem("centralSidebarOpen", "S");
    }
}

function toggleLinksUteis() {
    var sidebar = document.getElementById("sidebar");
    var group = document.getElementById("linksUteisGroup");

    if (!sidebar.classList.contains("open")) {
        sidebar.classList.add("open");
        document.body.classList.add("sidebar-open");
    }

    group.classList.toggle("open");
}

window.onload = function () {
    var sidebarStatus = localStorage.getItem("centralSidebarOpen");

    if (sidebarStatus === "S") {
        document.getElementById("sidebar").classList.add("open");
        document.body.classList.add("sidebar-open");
    }
}
    </script>
</body>
</html>

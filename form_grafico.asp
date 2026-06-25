<!-- #include file="conexao.asp" -->

<%
If Session("matricula") <> "" Then

    If Session("funcao") = "INSPETOR" Or Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM" Then

        VAR_USUARIO_NOME = "" & Session("name")
        VAR_USUARIO_FUNCAO = "" & Session("funcao")
        VAR_USUARIO_NIVEL = "" & Session("nivel")
        VAR_USUARIO_MATRICULA = "" & Session("matricula")

        Function SqlSafe(valor)
            SqlSafe = Replace(Trim("" & valor), "'", "''")
        End Function

        Function HtmlSafe(valor)
            HtmlSafe = Server.HTMLEncode("" & valor)
        End Function

        Function BuildPageUrl(pg, busca, autor, dataTxt, codigo)
            Dim u
            u = "form_grafico.asp?pg=" & pg

            If busca <> "" Then u = u & "&busca=" & Server.URLEncode(busca)
            If autor <> "" Then u = u & "&autor=" & Server.URLEncode(autor)
            If dataTxt <> "" Then u = u & "&data_txt=" & Server.URLEncode(dataTxt)
            If codigo <> "" Then u = u & "&codigo=" & Server.URLEncode(codigo)

            BuildPageUrl = u
        End Function

        busca = Trim("" & Request("busca"))
        autor = Trim("" & Request("autor"))
        dataTxt = Trim("" & Request("data_txt"))
        codigo = Trim("" & Request("codigo"))

        pagina = Trim("" & Request("pg"))
        If pagina = "" Then pagina = "1"
        If Not IsNumeric(pagina) Then pagina = "1"

        pagina = CInt(pagina)
        If pagina < 1 Then pagina = 1

        limite = 10
        inicio = (pagina - 1) * limite

        whereSQL = " WHERE 1=1 "

        If busca <> "" Then
            whereSQL = whereSQL & " AND (TBL_INFORMACAO.TITULO LIKE '%" & SqlSafe(busca) & "%' " & _
                                  " OR TBL_INFORMACAO.INFORMACAO LIKE '%" & SqlSafe(busca) & "%') "
        End If

        If autor <> "" Then
            whereSQL = whereSQL & " AND TBL_USUARIO.NOME LIKE '%" & SqlSafe(autor) & "%' "
        End If

        If dataTxt <> "" Then
            whereSQL = whereSQL & " AND TBL_INFORMACAO.DATA LIKE '%" & SqlSafe(dataTxt) & "%' "
        End If

        If codigo <> "" And IsNumeric(codigo) Then
            whereSQL = whereSQL & " AND TBL_INFORMACAO.COD=" & CLng(codigo) & " "
        End If

        fromSQL = " FROM TBL_INFORMACAO " & _
                  " INNER JOIN TBL_USUARIO ON TBL_INFORMACAO.FK_MATRICULA = TBL_USUARIO.MATRICULA "

        orderSQL = " ORDER BY TBL_INFORMACAO.COD DESC "

        sqlCount = "SELECT Count(*) AS TOTAL " & fromSQL & whereSQL
        Set rsCount = conexao_.Execute(sqlCount)

        totalRegistros = 0

        If Not rsCount.EOF Then
            If Not IsNull(rsCount("TOTAL")) Then
                totalRegistros = CLng(rsCount("TOTAL"))
            End If
        End If

        totalPaginas = 1

        If totalRegistros > 0 Then
            totalPaginas = Int((totalRegistros + limite - 1) / limite)
        End If

        If pagina > totalPaginas Then pagina = totalPaginas
        If pagina < 1 Then pagina = 1

        inicio = (pagina - 1) * limite

        If inicio = 0 Then
            sqlPage = "SELECT TOP " & limite & _
                      " TBL_INFORMACAO.COD, " & _
                      " TBL_INFORMACAO.DATA, " & _
                      " TBL_INFORMACAO.TITULO, " & _
                      " TBL_INFORMACAO.TIPO, " & _
                      " TBL_USUARIO.NOME AS AUTOR " & _
                      fromSQL & whereSQL & orderSQL
        Else
            sqlPage = "SELECT TOP " & limite & _
                      " TBL_INFORMACAO.COD, " & _
                      " TBL_INFORMACAO.DATA, " & _
                      " TBL_INFORMACAO.TITULO, " & _
                      " TBL_INFORMACAO.TIPO, " & _
                      " TBL_USUARIO.NOME AS AUTOR " & _
                      fromSQL & whereSQL & _
                      " AND TBL_INFORMACAO.COD NOT IN (" & _
                            "SELECT TOP " & inicio & " TBL_INFORMACAO.COD " & fromSQL & whereSQL & orderSQL & _
                      ") " & orderSQL
        End If

        Set RESULTADO_SQL_1 = conexao_.Execute(sqlPage)
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Relatórios</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <script src="libs/js/jquery.js"></script>

    <style>
        :root {
            --vale-teal: #00857A;
            --vale-teal-dark: #006B63;
            --vale-yellow: #F6B800;
            --text-dark: #263238;
            --text-muted: #607D8B;
            --card-bg: rgba(255,255,255,0.92);
            --sidebar-bg: rgba(255,255,255,0.86);
            --sidebar-collapsed: 88px;
            --sidebar-expanded: 292px;
        }

        * {
            box-sizing: border-box;
        }

        html,
        body {
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
            box-shadow: 5px 0 24px rgba(0,0,0,0.18);
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
            border-bottom: 1px solid rgba(0,133,122,0.12);
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
            border-top: 1px solid rgba(0,133,122,0.12);
        }

        .sidebar-link,
        .external-toggle {
            width: 100%;
            min-height: 58px;
            border: none;
            border-radius: 18px;
            background: rgba(255,255,255,0.78);
            color: var(--text-dark);
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 0 12px;
            cursor: pointer;
            box-shadow: 0 7px 16px rgba(0,0,0,0.10);
            transition: all 0.2s ease-in-out;
            font-family: Arial, Helvetica, sans-serif;
        }

        .sidebar-link:hover,
        .external-toggle:hover {
            background: #ffffff;
            transform: translateX(4px);
            box-shadow: 0 12px 25px rgba(0,0,0,0.18);
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

        /* Links úteis */

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
            background: rgba(255,255,255,0.68);
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

        .external-link span {
            font-size: 13px;
            font-weight: 800;
            color: var(--vale-teal-dark);
        }

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

        .external-icon-stack-folder .folder-card {
            position: absolute !important;
            width: 29px !important;
            height: 29px !important;
            border-radius: 10px !important;
            background: #ffffff !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            box-shadow: 0 3px 8px rgba(0,0,0,0.22) !important;
            border: 2px solid rgba(255,255,255,0.95) !important;
            overflow: hidden !important;
        }

        .external-icon-stack-folder .folder-card img {
            width: 100% !important;
            height: 100% !important;
            object-fit: cover !important;
            border: none !important;
            border-radius: 8px !important;
        }

        .external-icon-stack-folder .folder-card-back {
            left: 1px !important;
            top: 5px !important;
            z-index: 1 !important;
            opacity: 0.92 !important;
            transform: rotate(-8deg) scale(0.92) !important;
        }

        .external-icon-stack-folder .folder-card-middle {
            left: 7px !important;
            top: 2px !important;
            z-index: 2 !important;
            opacity: 0.96 !important;
            transform: rotate(5deg) scale(0.96) !important;
        }

        .external-icon-stack-folder .folder-card-front {
            left: 11px !important;
            top: 10px !important;
            z-index: 3 !important;
            transform: rotate(0deg) scale(1) !important;
        }

        /* =====================================================
           CONTEÚDO
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
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            gap: 10px;
        }

        .page-title {
            margin: 0;
            width: 100%;
            text-align: center;
            color: #ffffff;
            font-size: clamp(32px, 4vw, 64px);
            line-height: 1.05;
            font-weight: 900;
            letter-spacing: clamp(1px, 0.35vw, 3px);
            text-transform: uppercase;
            text-shadow: 0 4px 12px rgba(0,0,0,0.48);
        }

        .user-info {
            display: inline-block;
            padding: 7px 18px;
            border-radius: 999px;
            background: rgba(0,107,99,0.62);
            color: #ffffff;
            font-size: 13px;
            line-height: 1.35;
            text-shadow: 0 2px 7px rgba(0,0,0,0.35);
            backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
        }

        .summary-bar {
            width: min(100%, 1400px);
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

        .top-controls {
            width: min(100%, 1400px);
            margin: 20px auto;
            display: grid;
            grid-template-columns: minmax(680px, 1fr) auto 220px;
            align-items: center;
            gap: 18px;
        }

        .filters-box {
            display: grid;
            grid-template-columns: auto minmax(140px, 1fr) minmax(140px, 1fr) minmax(120px, 1fr) auto;
            align-items: center;
            gap: 10px;
            background: rgba(255,255,255,0.72);
            padding: 10px 14px;
            border-radius: 22px;
            min-width: 0;
            width: 100%;
        }

        .filters-title {
            font-weight: 800;
            color: #444;
            white-space: nowrap;
        }

        .filters-box input {
            width: 100%;
            min-width: 0;
            height: 42px;
            padding: 0 12px;
            border-radius: 12px;
            border: none;
            background: #ffffff;
            font-size: 14px;
            outline: none;
        }

        .filters-box button {
            height: 42px;
            padding: 0 18px;
            border: none;
            border-radius: 12px;
            background: var(--vale-teal);
            color: #ffffff;
            font-weight: 800;
            white-space: nowrap;
            cursor: pointer;
        }

        .top-pagination {
            color: #ffffff;
            font-weight: 800;
            white-space: nowrap;
            justify-self: center;
        }

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
            outline: none;
        }

        .report-list {
            width: min(100%, 1400px);
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 14px;
        }

        .report-card {
            display: grid;
            grid-template-columns: 90px 130px 1fr 170px 150px;
            align-items: center;
            gap: 14px;
            padding: 18px 20px;
            border-radius: 22px;
            background: var(--card-bg);
            box-shadow: 0 14px 30px rgba(0,0,0,0.18);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            transition: all 0.2s ease-in-out;
        }

        .report-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 20px 38px rgba(0,0,0,0.23);
        }

        .report-cod {
            min-height: 42px;
            border-radius: 14px;
            background: var(--vale-teal);
            color: #ffffff;
            font-weight: 900;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .report-date,
        .report-author,
        .report-title {
            color: var(--text-dark);
            font-size: 14px;
            font-weight: 700;
        }

        .report-author {
            color: var(--text-muted);
        }

        .report-title {
            font-size: 15px;
            font-weight: 900;
            color: var(--vale-teal-dark);
        }

        .btn-chart {
            min-height: 42px;
            border: none;
            border-radius: 12px;
            background: var(--vale-teal);
            color: #ffffff;
            font-size: 14px;
            font-weight: 900;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
        }

        .btn-chart:hover {
            background: var(--vale-teal-dark);
            transform: translateY(-1px);
        }

        .empty-state {
            width: min(100%, 1400px);
            margin: 0 auto;
            padding: 28px;
            border-radius: 20px;
            background: var(--card-bg);
            text-align: center;
            font-weight: 800;
        }

        .pagination {
            width: min(100%, 1400px);
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

        .page-current {
            background: var(--vale-teal);
            color: #ffffff;
        }

        .page-dots {
            color: #ffffff;
        }

        /* Modal */

        .modal-overlay {
            display: none;
            position: fixed;
            inset: 0;
            z-index: 999;
            background: rgba(0,0,0,0.62);
            align-items: center;
            justify-content: center;
            padding: 24px;
        }

        .modal-overlay.open {
            display: flex;
        }

        .modal-box {
            width: min(96vw, 920px);
            max-height: 88vh;
            overflow-y: auto;
            border-radius: 26px;
            background: #ffffff;
            box-shadow: 0 24px 70px rgba(0,0,0,0.35);
        }

        .modal-header {
            min-height: 64px;
            padding: 18px 24px;
            background: var(--vale-teal);
            color: #ffffff;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-radius: 26px 26px 0 0;
        }

        .modal-title {
            font-size: 18px;
            font-weight: 900;
        }

        .modal-close {
            width: 36px;
            height: 36px;
            border: none;
            border-radius: 12px;
            background: rgba(255,255,255,0.18);
            color: #ffffff;
            font-size: 20px;
            cursor: pointer;
        }

        .modal-content-area {
            padding: 24px;
        }

        @media (max-width: 1280px) {
            .top-controls {
                grid-template-columns: 1fr;
                gap: 14px;
            }

            .top-pagination {
                justify-self: start;
            }

            .search-box {
                justify-self: start;
                max-width: 280px;
            }

            .report-card {
                grid-template-columns: 80px 120px 1fr 150px;
            }

            .report-author {
                display: none;
            }
        }

        @media (max-width: 900px) {
            body.sidebar-open .main-content {
                margin-left: var(--sidebar-collapsed);
                width: calc(100vw - var(--sidebar-collapsed));
            }

            .report-card {
                grid-template-columns: 1fr;
                align-items: stretch;
            }

            .btn-chart {
                width: 100%;
            }

            .filters-box {
                grid-template-columns: 1fr 1fr;
            }

            .filters-title {
                grid-column: 1 / -1;
            }

            .filters-box button {
                width: 100%;
            }
        }

        @media (max-width: 620px) {
            :root {
                --sidebar-collapsed: 76px;
                --sidebar-expanded: 240px;
            }

            .main-content {
                margin-left: var(--sidebar-collapsed);
                width: calc(100vw - var(--sidebar-collapsed));
                padding: 16px;
            }

            body.sidebar-open .main-content {
                margin-left: var(--sidebar-collapsed);
                width: calc(100vw - var(--sidebar-collapsed));
            }

            .page-title {
                font-size: clamp(26px, 8vw, 40px);
            }

            .filters-box {
                grid-template-columns: 1fr;
            }

            .search-box {
                max-width: 100%;
            }

            .summary-pill {
                width: 100%;
                text-align: center;
            }
        }

		/* =====================================================
   MODAL DO GRÁFICO COM IFRAME
===================================================== */

.modal-overlay {
    display: none;
    position: fixed;
    inset: 0;
    z-index: 9999;
    background: rgba(0, 0, 0, 0.62);
    align-items: center;
    justify-content: center;
    padding: 24px;
}

.modal-overlay.open {
    display: flex;
}

.modal-box {
    width: min(96vw, 980px);
    height: min(90vh, 820px);
    border-radius: 26px;
    background: #ffffff;
    box-shadow: 0 24px 70px rgba(0, 0, 0, 0.35);
    overflow: hidden;
    display: flex;
    flex-direction: column;
}

.modal-header {
    min-height: 64px;
    padding: 18px 24px;
    background: #00857A;
    color: #ffffff;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.modal-title {
    font-size: 18px;
    font-weight: 900;
}

.modal-close {
    width: 36px;
    height: 36px;
    border: none;
    border-radius: 12px;
    background: rgba(255, 255, 255, 0.18);
    color: #ffffff;
    font-size: 20px;
    cursor: pointer;
}

.modal-close:hover {
    background: rgba(255, 255, 255, 0.28);
}

.modal-content-area {
    flex: 1;
    padding: 0;
    overflow: hidden;
}

.grafico-frame {
    width: 100%;
    height: 100%;
    border: none;
    background: #ffffff;
}

/* =====================================================
   MODAL DO GRÁFICO
===================================================== */

.modal-overlay {
    display: none;
    position: fixed;
    inset: 0;
    z-index: 9999;
    background: rgba(0, 0, 0, 0.62);
    align-items: center;
    justify-content: center;
    padding: 24px;
}

.modal-overlay.open {
    display: flex;
}

.modal-box {
    width: min(96vw, 980px);
    height: min(90vh, 820px);
    border-radius: 26px;
    background: #ffffff;
    box-shadow: 0 24px 70px rgba(0, 0, 0, 0.35);
    overflow: hidden;
    display: flex;
    flex-direction: column;
}

.modal-header {
    min-height: 64px;
    padding: 18px 24px;
    background: #00857A;
    color: #ffffff;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.modal-title {
    font-size: 18px;
    font-weight: 900;
}

.modal-close {
    width: 36px;
    height: 36px;
    border: none;
    border-radius: 12px;
    background: rgba(255, 255, 255, 0.18);
    color: #ffffff;
    font-size: 20px;
    cursor: pointer;
}

.modal-close:hover {
    background: rgba(255, 255, 255, 0.28);
}

.modal-content-area {
    flex: 1;
    padding: 0;
    overflow: hidden;
}

.grafico-frame {
    width: 100%;
    height: 100%;
    border: none;
    background: #ffffff;
}
    </style>
</head>

<body>

    <aside id="sidebar" class="sidebar">

        <div class="sidebar-header">
            <button type="button" class="menu-toggle" onclick="toggleSidebar()" title="Abrir ou fechar menu">
                ☰
            </button>

            <div class="sidebar-logo-area">
                <img src="libs/img/logo-vale.png" alt="Logo Vale" onerror="this.onerror=null;this.src='libs/img/Logotipo_Vale.png';">
            </div>
        </div>

        <nav class="sidebar-nav">

            <a href="form_home.asp" class="sidebar-link">
                <span class="sidebar-icon">
                    <img src="libs/img/img_link/Home.jpg" alt="Home">
                </span>
                <span class="sidebar-text">
                    <span class="sidebar-title">Home</span>
                </span>
            </a>

            <a href="form_visualizar_registro.asp" class="sidebar-link">
                <span class="sidebar-icon">
                    <img src="libs/img/img_link/visualizar.jpg" alt="Visualizar">
                </span>
                <span class="sidebar-text">
                    <span class="sidebar-title">Visualizar registro</span>
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

            <a href="form_grafico.asp" class="sidebar-link">
                <span class="sidebar-icon">
                    <img src="libs/img/img_link/relatorios.jpg" alt="Relatórios">
                </span>
                <span class="sidebar-text">
                    <span class="sidebar-title">Relatórios</span>
                </span>
            </a>

            <a href="https://efvmworkplace/central%20de%20kaizen/index.php" class="sidebar-link" target="_blank">
                <span class="sidebar-icon">
                    <img src="libs/img/img_link/kaizen.jpg" alt="Kaizen">
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

    <main class="main-content">

        <header class="page-header">
            <h1 class="page-title">Relatórios e Gráficos</h1>
            <div class="user-info">
                Bem-vindo(a), <strong><%=HtmlSafe(VAR_USUARIO_NOME)%></strong>
                &nbsp;·&nbsp;
                <%=HtmlSafe(VAR_USUARIO_FUNCAO)%>
            </div>
        </header>

        <div class="summary-bar">
            <div class="summary-pill">Total encontrado: <%=totalRegistros%></div>
            <div class="summary-pill">Página <%=pagina%> de <%=totalPaginas%></div>
            <div class="summary-pill">Registros por página: <%=limite%></div>
        </div>

        <div class="top-controls">

            <form class="filters-box" method="get">
                <span class="filters-title">Filtros</span>

                <input type="text" name="codigo" placeholder="Código" value="<%=HtmlSafe(codigo)%>">
                <input type="text" name="data_txt" placeholder="Data" value="<%=HtmlSafe(dataTxt)%>">
                <input type="text" name="autor" placeholder="Autor" value="<%=HtmlSafe(autor)%>">

                <button type="submit">Aplicar</button>
            </form>

            <div class="top-pagination">
                Página <strong><%=pagina%></strong>
            </div>

            <form class="search-box" method="get">
                <input type="text" name="busca" placeholder="Pesquisar título..." value="<%=HtmlSafe(busca)%>">
            </form>

        </div>

        <section class="report-list">
            <%
            If Not RESULTADO_SQL_1.EOF Then

                Do Until RESULTADO_SQL_1.EOF

                    VAR_SQL_1_COD = RESULTADO_SQL_1("COD")
                    VAR_SQL_1_DATA = RESULTADO_SQL_1("DATA")
                    VAR_SQL_1_TITULO = RESULTADO_SQL_1("TITULO")
                    VAR_SQL_1_TIPO = RESULTADO_SQL_1("TIPO")
                    VAR_SQL_2_NOME = RESULTADO_SQL_1("AUTOR")
            %>

                <article class="report-card">
                    <div class="report-cod">#<%=VAR_SQL_1_COD%></div>
                    <div class="report-date"><%=HtmlSafe(VAR_SQL_1_DATA)%></div>
                    <div class="report-title"><%=HtmlSafe(VAR_SQL_1_TITULO)%></div>
                    <div class="report-author"><%=HtmlSafe(VAR_SQL_2_NOME)%></div>

                    <button type="button" class="btn-chart" onclick="exibe_grafico('<%=VAR_SQL_1_COD%>');">
    Gráfico
</button>
                </article>

            <%
                    RESULTADO_SQL_1.MoveNext
                Loop

            Else
            %>
                <div class="empty-state">
                    Nenhum registro encontrado para os filtros aplicados.
                </div>
            <%
            End If
            %>
        </section>

        <div class="pagination">
            <%
            If totalPaginas <= 1 Then

                Response.Write "<span class='page-current'>1</span>"

            Else

                If pagina > 1 Then
                    Response.Write "<a class='page-link' href='" & BuildPageUrl(pagina-1, busca, autor, dataTxt, codigo) & "'>&laquo;</a>"
                End If

                If pagina = 1 Then
                    Response.Write "<span class='page-current'>1</span>"
                Else
                    Response.Write "<a class='page-link' href='" & BuildPageUrl(1, busca, autor, dataTxt, codigo) & "'>1</a>"
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
                        Response.Write "<a class='page-link' href='" & BuildPageUrl(i, busca, autor, dataTxt, codigo) & "'>" & i & "</a>"
                    End If

                Next

                If endPage < totalPaginas - 1 Then
                    Response.Write "<span class='page-dots'>...</span>"
                End If

                If totalPaginas > 1 Then
                    If pagina = totalPaginas Then
                        Response.Write "<span class='page-current'>" & totalPaginas & "</span>"
                    Else
                        Response.Write "<a class='page-link' href='" & BuildPageUrl(totalPaginas, busca, autor, dataTxt, codigo) & "'>" & totalPaginas & "</a>"
                    End If
                End If

                If pagina < totalPaginas Then
                    Response.Write "<a class='page-link' href='" & BuildPageUrl(pagina+1, busca, autor, dataTxt, codigo) & "'>&raquo;</a>"
                End If

            End If
            %>
        </div>

    </main>

    <div id="ModalGrafico" class="modal-overlay">
    <div class="modal-box">
        <div class="modal-header">
            <div class="modal-title">Gráfico de visualização</div>

            <button type="button" class="modal-close" onclick="fecha_grafico()">
                ×
            </button>
        </div>

        <div class="modal-content-area">
            <iframe
                id="IframeGrafico"
                class="grafico-frame"
                src=""
                title="Gráfico de visualização">
            </iframe>
        </div>
    </div>
</div>
</div>

    <script>
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

        window.onload = function () {
            var sidebarStatus = localStorage.getItem("centralSidebarOpen");
            var linksStatus = localStorage.getItem("centralLinksUteisOpen");

            if (sidebarStatus === "S") {
                document.getElementById("sidebar").classList.add("open");
                document.body.classList.add("sidebar-open");
            }

            if (linksStatus === "S") {
                document.getElementById("linksUteisGroup").classList.add("open");
            }
        }

</script>
        }

        function fecha_grafico() {
            document.getElementById("ModalGrafico").classList.remove("open");
            document.getElementById("ContentModalGrafico").innerHTML = "";
        }
    </script>

	<script>
function exibe_grafico(cod) {
    var modal = document.getElementById("ModalGrafico");
    var iframe = document.getElementById("IframeGrafico");

    if (!modal || !iframe) {
        alert("Não foi possível abrir o gráfico. Verifique se o modal do gráfico está na página.");
        return;
    }

    iframe.src = "aux_grafico.asp?COD=" + encodeURIComponent(cod) + "&t=" + new Date().getTime();

    modal.classList.add("open");
}

function fecha_grafico() {
    var modal = document.getElementById("ModalGrafico");
    var iframe = document.getElementById("IframeGrafico");

    if (modal) {
        modal.classList.remove("open");
    }

    if (iframe) {
        setTimeout(function () {
            iframe.src = "";
        }, 200);
    }
}
</script>

</body>
</html>

<%
    Else
        Response.Redirect("form_home.asp")
    End If

Else
    Response.Redirect("form_login.asp")
End If
%>
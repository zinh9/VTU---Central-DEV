<!-- #include file="conexao.asp" -->

<%
If session("matricula") = "" Then
    response.redirect("form_login.asp")
End If

VAR_USUARIO_NOME = Session("name")
VAR_USUARIO_FUNCAO = Session("funcao")
VAR_USUARIO_NIVEL = Session("nivel")
VAR_USUARIO_MATRICULA = Session("matricula")

Function TextoCurto(valor, limite)
    Dim texto
    texto = "" & valor

    texto = Replace(texto, vbCrLf, " ")
    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")

    If Len(texto) > limite Then
        texto = Left(texto, limite) & "..."
    End If

    TextoCurto = Server.HTMLEncode(texto)
End Function

Function TratarImagem(valor)
    Dim imagem
    imagem = Trim("" & valor)

    If imagem = "" Or UCase(imagem) = "NOIMG" Then
        TratarImagem = "NOIMG.jpg"
    Else
        imagem = Replace(imagem, "\", "/")
        TratarImagem = imagem
    End If
End Function

Function EhImagem(valor)
    Dim arquivo
    arquivo = LCase("" & valor)

    If InStr(arquivo, ".jpg") > 0 Or InStr(arquivo, ".jpeg") > 0 Or InStr(arquivo, ".png") > 0 Or InStr(arquivo, ".gif") > 0 Or InStr(arquivo, ".jfif") > 0 Then
        EhImagem = True
    Else
        EhImagem = False
    End If
End Function

Function TipoDocumento(valor)
    Dim arquivo
    arquivo = LCase("" & valor)

    If InStr(arquivo, ".pdf") > 0 Then
        TipoDocumento = "PDF"
    ElseIf InStr(arquivo, ".ppt") > 0 Or InStr(arquivo, ".pptx") > 0 Then
        TipoDocumento = "PPT"
    ElseIf InStr(arquivo, ".bin") > 0 Then
        TipoDocumento = "ARQ"
    Else
        TipoDocumento = "DOC"
    End If
End Function

VAR_QTD_PENDENCIAS = 0

SQL_PENDENCIAS = "SELECT Count(*) AS TOTAL_PENDENCIAS " & _
                 "FROM (" & _
                 " SELECT DISTINCT TBL_CHECK_INFORMACAO.COD " & _
                 " FROM TBL_CHECK_INFORMACAO " & _
                 " INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
                 " WHERE TBL_CHECK_INFORMACAO.MATRICULA='" & Replace(VAR_USUARIO_MATRICULA, "'", "''") & "' " & _
                 " AND TBL_CHECK_INFORMACAO.COD > 0 " & _
                 " AND UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='NÃO'" & _
                 ") AS Q"

Set RS_PENDENCIAS = conexao_.Execute(SQL_PENDENCIAS)

If Not RS_PENDENCIAS.EOF Then
    If Not IsNull(RS_PENDENCIAS("TOTAL_PENDENCIAS")) Then
        VAR_QTD_PENDENCIAS = CLng(RS_PENDENCIAS("TOTAL_PENDENCIAS"))
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Início</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        :root {
            --vale-teal: #00857A;
            --vale-teal-dark: #006B63;
            --vale-yellow: #F6B800;
            --text-dark: #263238;
            --text-muted: #607D8B;
            --card-bg: rgba(255, 255, 255, 0.88);
            --sidebar-bg: rgba(255, 255, 255, 0.86);
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
                linear-gradient(rgba(0, 60, 55, 0.16), rgba(0, 60, 55, 0.16)),
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

        /* =====================================================
           CONTEÚDO PRINCIPAL
        ===================================================== */

        .main-content {
            min-height: 100vh;
            margin-left: var(--sidebar-collapsed);
            padding: 26px 42px 44px 42px;
            transition: margin-left 0.25s ease-in-out;
        }

        body.sidebar-open .main-content {
            margin-left: var(--sidebar-expanded);
        }

        .page-header {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto 24px auto;
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

        .user-info strong {
            font-weight: 800;
        }

        .summary-row {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto 24px auto;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 18px;
        }

        .summary-card {
            min-height: 86px;
            border-radius: 22px;
            padding: 18px 24px;
            background: var(--card-bg);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            box-shadow: 0 14px 28px rgba(0, 0, 0, 0.18);
        }

        .summary-label {
            display: block;
            font-size: 12px;
            font-weight: 800;
            color: var(--vale-teal-dark);
            text-transform: uppercase;
            margin-bottom: 7px;
        }

        .summary-value {
            display: block;
            font-size: 26px;
            font-weight: 900;
            color: var(--text-dark);
        }

        .section-title {
            width: 100%;
            max-width: 1240px;
            margin: 28px auto 18px auto;
            color: #ffffff;
            font-size: 24px;
            font-weight: 900;
            text-shadow: 0 3px 9px rgba(0, 0, 0, 0.48);
        }

        /* =====================================================
           CARDS DOS ÚLTIMOS REGISTROS
        ===================================================== */

        .records-list {
            width: 100%;
            max-width: 1240px;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 22px;
        }

        .record-card {
            display: grid;
            grid-template-columns: 200px 1fr;
            gap: 26px;
            min-height: 190px;
            padding: 26px;
            border-radius: 26px;
            background: var(--card-bg);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            box-shadow: 0 16px 34px rgba(0, 0, 0, 0.22);
            transition: all 0.2s ease-in-out;
        }

        .record-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 22px 42px rgba(0, 0, 0, 0.26);
        }

        .record-media {
            width: 200px;
            height: 150px;
            border-radius: 18px;
            overflow: hidden;
            background: #E4F2F0;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: inset 0 0 0 1px rgba(0, 133, 122, 0.20);
        }

        .record-media img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .document-thumb {
            width: 100%;
            height: 100%;
            background: linear-gradient(135deg, var(--vale-teal), var(--vale-teal-dark));
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-direction: column;
            font-weight: 800;
        }

        .document-thumb span:first-child {
            font-size: 30px;
            margin-bottom: 4px;
        }

        .document-thumb span:last-child {
            font-size: 15px;
        }

        .record-content {
            min-width: 0;
        }

        .record-top {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 18px;
            margin-bottom: 10px;
        }

        .record-title {
            margin: 0;
            font-size: clamp(20px, 1.7vw, 28px);
            line-height: 1.22;
            color: var(--vale-teal-dark);
            font-weight: 900;
        }

        .record-cod {
            min-width: 78px;
            padding: 8px 13px;
            border-radius: 24px;
            background: var(--vale-teal);
            color: #ffffff;
            text-align: center;
            font-size: 14px;
            font-weight: 900;
            flex-shrink: 0;
        }

        .record-meta {
            margin-bottom: 12px;
            font-size: 14px;
            color: var(--text-muted);
            font-weight: 800;
        }

        .record-text {
            margin: 0;
            font-size: 16px;
            line-height: 1.48;
            color: #37474F;
        }

        .record-actions {
            margin-top: 18px;
        }

        .record-link {
            display: inline-block;
            padding: 11px 18px;
            border-radius: 11px;
            background: var(--vale-teal-dark);
            color: #ffffff;
            text-decoration: none;
            font-size: 14px;
            font-weight: 800;
            transition: all 0.2s ease-in-out;
        }

        .record-link:hover {
            background: #004D47;
            transform: translateY(-1px);
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

        .quick-links,
        .bottom-links {
            display: none !important;
        }

        /* =====================================================
           RESPONSIVIDADE
        ===================================================== */

        @media (max-width: 1180px) {
            .summary-row {
                grid-template-columns: 1fr;
            }

            .record-card {
                grid-template-columns: 1fr;
            }

            .record-media {
                width: 100%;
                height: 230px;
            }
        }

        @media (max-width: 760px) {
            :root {
                --sidebar-collapsed: 76px;
                --sidebar-expanded: 246px;
            }

            .sidebar {
                padding-left: 10px;
                padding-right: 10px;
            }

            .menu-toggle {
                width: 52px;
                height: 52px;
            }

            .sidebar-icon {
                width: 36px;
                height: 36px;
            }

            .main-content {
                padding: 22px 18px 34px 18px;
            }

            .page-title {
                font-size: clamp(26px, 8vw, 38px);
                letter-spacing: 1.5px;
            }

            .record-card {
                padding: 20px;
            }

            .record-media {
                height: 200px;
            }
        }
        .main-content {
           max-width: 1600px;
           margin: 0 auto;
        } 
        /* =====================================================
   HOME - CORREÇÃO VISUAL DAS LOGOS DOS LINKS ÚTEIS
   Sem alterar links/endpoints.
===================================================== */

/* Ícone do botão "Links úteis" com logos agrupadas */
.external-icon-stack {
    position: relative !important;
    background: #ffffff !important;
    overflow: visible !important;
}

/* Caso existam imagens dentro do botão agrupado */
.external-icon-stack img {
    display: block !important;
    position: absolute !important;
    width: 24px !important;
    height: 24px !important;
    border-radius: 8px !important;
    object-fit: cover !important;
    border: 2px solid #ffffff !important;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.22) !important;
}

/* Distribuição das imagens agrupadas */
.external-icon-stack img:nth-child(1) {
    left: 2px !important;
    top: 5px !important;
}

.external-icon-stack img:nth-child(2) {
    right: 2px !important;
    top: 5px !important;
}

.external-icon-stack img:nth-child(3) {
    left: 7px !important;
    bottom: 2px !important;
}

/* Caso o botão tenha apenas o ícone de globo, cria visual agrupado via CSS */
.external-icon-stack::before,
.external-icon-stack::after {
    content: "";
    position: absolute;
    width: 22px;
    height: 22px;
    border-radius: 8px;
    background-size: cover;
    background-position: center;
    border: 2px solid #ffffff;
    box-shadow: 0 2px 5px rgba(0,0,0,0.20);
}

.external-icon-stack::before {
    left: 3px;
    top: 6px;
    background-image: url("libs/img/img_link/gdb.jpg");
}

.external-icon-stack::after {
    right: 3px;
    top: 6px;
    background-image: url("libs/img/img_link/irislogo.png");
}

/* Terceira logo agrupada */
.external-icon-stack {
    background-image: url("libs/img/img_link/edados.jpg") !important;
    background-size: 22px 22px !important;
    background-repeat: no-repeat !important;
    background-position: center bottom 3px !important;
}

/* =====================================================
   Logos individuais dos Links Úteis por ordem
   Funciona mesmo quando aparecem letras E/G/V/E/I.
===================================================== */

.external-list .external-link {
    position: relative !important;
}

/* Remove aparência de letra fallback */
.external-list .external-link .external-logo-fallback,
.external-list .external-link .external-logo {
    font-size: 0 !important;
    color: transparent !important;
    background-color: transparent !important;
    background-size: cover !important;
    background-position: center !important;
    background-repeat: no-repeat !important;
}

/* 1 - Equipfer */
.external-list .external-link:nth-child(1) .external-logo-fallback,
.external-list .external-link:nth-child(1) .external-logo {
    background-image: url("libs/img/img_link/tooplate_image_01.jpg") !important;
}

/* 2 - GDB */
.external-list .external-link:nth-child(2) .external-logo-fallback,
.external-list .external-link:nth-child(2) .external-logo {
    background-image: url("libs/img/img_link/gdb.jpg") !important;
}

/* 3 - VES */
.external-list .external-link:nth-child(3) .external-logo-fallback,
.external-list .external-link:nth-child(3) .external-logo {
    background-image: url("libs/img/img_link/prontidao1.jpg") !important;
}

/* 4 - E-Dados */
.external-list .external-link:nth-child(4) .external-logo-fallback,
.external-list .external-link:nth-child(4) .external-logo {
    background-image: url("libs/img/img_link/edados.jpg") !important;
}

/* 5 - Iris */
.external-list .external-link:nth-child(5) .external-logo-fallback,
.external-list .external-link:nth-child(5) .external-logo {
    background-image: url("libs/img/img_link/irislogo.png") !important;
}

/* Se o link já tiver img, força exibição da imagem */
.external-list .external-link img {
    display: block !important;
    width: 31px !important;
    height: 31px !important;
    border-radius: 11px !important;
    object-fit: cover !important;
    flex-shrink: 0 !important;
}

/* Garante tamanho visual correto do bloco da logo */
.external-logo,
.external-logo-fallback,
.external-list .external-link img {
    width: 31px !important;
    height: 31px !important;
    min-width: 31px !important;
    min-height: 31px !important;
}
     /* =====================================================
   LINKS ÚTEIS - CORREÇÃO FINAL DAS LOGOS
===================================================== */

.external-group {
    width: 100%;
}

/* Botão principal Links úteis */
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

.external-toggle:hover {
    background: #ffffff;
    transform: translateX(4px);
    box-shadow: 0 12px 25px rgba(0, 0, 0, 0.18);
}

/* Ícone agrupado do botão Links úteis */
.external-icon-stack {
    position: relative !important;
    background: #ffffff !important;
    overflow: visible !important;
}

.external-icon-stack img {
    display: block !important;
    position: absolute !important;
    width: 24px !important;
    height: 24px !important;
    border-radius: 9px !important;
    object-fit: cover !important;
    border: 2px solid #ffffff !important;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.22) !important;
}

/* Distribuição das três logos agrupadas */
.external-icon-stack img:nth-child(1) {
    left: 2px !important;
    top: 5px !important;
}

.external-icon-stack img:nth-child(2) {
    right: 2px !important;
    top: 5px !important;
}

.external-icon-stack img:nth-child(3) {
    left: 7px !important;
    bottom: 2px !important;
}

/* Seta do Links úteis */
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

/* Lista dos links úteis */
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

/* Cada link externo */
.external-link {
    min-height: 48px;
    border-radius: 15px;
    background: rgba(255, 255, 255, 0.70);
    color: var(--text-dark);
    text-decoration: none !important;
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

/* Imagens reais dos links */
.external-link img,
.external-logo {
    width: 31px !important;
    height: 31px !important;
    min-width: 31px !important;
    min-height: 31px !important;
    border-radius: 11px !important;
    object-fit: cover !important;
    flex-shrink: 0 !important;
    display: block !important;
}

/* Texto dos links */
.external-link span {
    font-size: 13px;
    font-weight: 800;
    color: var(--vale-teal-dark);
}

/* Fallback visual para Equipfer */
.external-logo-equipfer {
    background-image: url("libs/img/img_link/tooplate_image_01.jpg");
    background-size: cover;
    background-position: center;
    background-repeat: no-repeat;
}

/* Fallback visual para VES */
.external-logo-ves {
    background-image: url("libs/img/img_link/prontidao1.jpg");
    background-size: cover;
    background-position: center;
    background-repeat: no-repeat;
}
/* =====================================================
   CORREÇÃO DO ÍCONE AGRUPADO - LINKS ÚTEIS
   Remove sobreposição e organiza as logos em mini grade
===================================================== */

.external-icon-stack {
    background: #ffffff !important;
    display: grid !important;
    grid-template-columns: repeat(2, 15px) !important;
    grid-template-rows: repeat(2, 15px) !important;
    gap: 2px !important;
    align-content: center !important;
    justify-content: center !important;
    padding: 3px !important;
    overflow: hidden !important;
    position: relative !important;
}

.external-icon-stack img {
    position: static !important;
    width: 15px !important;
    height: 15px !important;
    min-width: 15px !important;
    min-height: 15px !important;
    max-width: 15px !important;
    max-height: 15px !important;
    border-radius: 5px !important;
    object-fit: cover !important;
    border: none !important;
    box-shadow: none !important;
}

/* terceira logo centralizada embaixo */
.external-icon-stack img:nth-child(3) {
    grid-column: 1 / span 2 !important;
    justify-self: center !important;
}   
/* =====================================================
   CORREÇÃO RESPONSIVA DA HOME
   Impede que o conteúdo fique atrás do menu lateral
===================================================== */

/* Estado padrão: conteúdo respeita menu fechado */
.main-content {
    margin-left: var(--sidebar-collapsed) !important;
    width: calc(100vw - var(--sidebar-collapsed)) !important;
    max-width: none !important;
    padding-left: clamp(22px, 3vw, 42px) !important;
    padding-right: clamp(22px, 3vw, 42px) !important;
}

/* Em telas grandes, quando menu abre, conteúdo acompanha */
body.sidebar-open .main-content {
    margin-left: var(--sidebar-expanded) !important;
    width: calc(100vw - var(--sidebar-expanded)) !important;
}

/* Garante que os blocos internos não passem por baixo do menu */
.page-header,
.summary-row,
.section-title,
.records-list {
    width: 100% !important;
    max-width: 1240px !important;
    margin-left: auto !important;
    margin-right: auto !important;
}

/* Título mais adaptável para notebook e monitor grande */
.page-title {
    font-size: clamp(30px, 3.6vw, 58px) !important;
    line-height: 1.08 !important;
    letter-spacing: clamp(1px, 0.28vw, 3px) !important;
    text-align: center !important;
    max-width: 100% !important;
}

/* Notebook e telas menores: menu aberto vira sobreposição, conteúdo não anda para trás */
@media (max-width: 1180px) {
    body.sidebar-open .main-content {
        margin-left: var(--sidebar-collapsed) !important;
        width: calc(100vw - var(--sidebar-collapsed)) !important;
    }

    .main-content {
        padding-left: 28px !important;
        padding-right: 24px !important;
    }

    .page-title {
        font-size: clamp(28px, 5vw, 44px) !important;
        letter-spacing: 1.5px !important;
    }

    .summary-row {
        grid-template-columns: 1fr !important;
    }

    .record-card {
        grid-template-columns: 180px 1fr !important;
        gap: 22px !important;
    }

    .record-media {
        width: 180px !important;
        height: 135px !important;
    }
}

/* Telas bem menores */
@media (max-width: 820px) {
    :root {
        --sidebar-collapsed: 76px;
        --sidebar-expanded: 252px;
    }

    .main-content {
        margin-left: var(--sidebar-collapsed) !important;
        width: calc(100vw - var(--sidebar-collapsed)) !important;
        padding-left: 18px !important;
        padding-right: 16px !important;
    }

    body.sidebar-open .main-content {
        margin-left: var(--sidebar-collapsed) !important;
        width: calc(100vw - var(--sidebar-collapsed)) !important;
    }

    .page-title {
        font-size: clamp(24px, 7vw, 36px) !important;
    }

    .record-card {
        grid-template-columns: 1fr !important;
    }

    .record-media {
        width: 100% !important;
        height: 210px !important;
    }
}
/* =====================================================
   MENU ABERTO EM TELAS MENORES COMO SOBREPOSIÇÃO
===================================================== */

@media (max-width: 1180px) {
    .sidebar.open {
        box-shadow: 10px 0 36px rgba(0, 0, 0, 0.30) !important;
    }

    .sidebar.open::after {
        content: "";
        position: fixed;
        left: var(--sidebar-expanded);
        top: 0;
        width: 0;
        height: 100vh;
    }
}
/* =====================================================
   CORREÇÃO FINAL - ÍCONE LINKS ÚTEIS
   Remove sobreposição/distorção das logos agrupadas
===================================================== */

/* Limpa qualquer efeito antigo aplicado ao botão Links úteis */
.external-icon-stack {
    width: 38px !important;
    height: 38px !important;
    min-width: 38px !important;
    min-height: 38px !important;
    max-width: 38px !important;
    max-height: 38px !important;

    padding: 4px !important;
    border-radius: 14px !important;

    background: #ffffff !important;
    background-image: none !important;

    display: grid !important;
    grid-template-columns: repeat(2, 1fr) !important;
    grid-template-rows: repeat(2, 1fr) !important;
    gap: 2px !important;

    align-items: center !important;
    justify-items: center !important;

    overflow: hidden !important;
    position: relative !important;
}

/* Remove pseudo-elementos antigos que estavam sobrepondo as imagens */
.external-icon-stack::before,
.external-icon-stack::after {
    content: none !important;
    display: none !important;
    background-image: none !important;
}

/* Corrige as imagens internas do botão agrupado */
.external-icon-stack img {
    position: static !important;

    width: 14px !important;
    height: 14px !important;
    min-width: 14px !important;
    min-height: 14px !important;
    max-width: 14px !important;
    max-height: 14px !important;

    border: none !important;
    border-radius: 4px !important;
    box-shadow: none !important;

    object-fit: contain !important;
    display: block !important;

    margin: 0 !important;
    padding: 0 !important;
}

/* Se houver 3 imagens, centraliza a terceira embaixo */
.external-icon-stack img:nth-child(3) {
    grid-column: 1 / span 2 !important;
    justify-self: center !important;
}

/* Garante que o ícone não estoure dentro do botão */
.external-toggle .sidebar-icon.external-icon-stack {
    flex-shrink: 0 !important;
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

    <!-- VISUALIZAR (todos) -->
    <a href="form_visualizar_registro.asp" class="sidebar-link">
        <span class="sidebar-icon">
            <img src="libs/img/img_link/visualizar.jpg">
        </span>
        <span class="sidebar-text">
            <span class="sidebar-title">Visualizar registro</span>
        </span>
    </a>

    <!-- INSERIR (somente SUPERVISOR / ADM) -->
    <% If VAR_USUARIO_FUNCAO = "SUPERVISOR" Or VAR_USUARIO_NIVEL = "ADM" Then %>
    <a href="form_formulario.asp" class="sidebar-link">
        <span class="sidebar-icon">
            <span>+</span>
        </span>
        <span class="sidebar-text">
            <span class="sidebar-title">Novo registro</span>
        </span>
    </a>
    <% End If %>

    <!-- RELATÓRIOS -->
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

    <!-- KAIZEN -->
    <a href="https://efvmworkplace/central%20de%20kaizen/index.php" class="sidebar-link" target="_blank">
        <span class="sidebar-icon">
            <img src="libs/img/img_link/kaizen.jpg">
        </span>
        <span class="sidebar-text">
            <span class="sidebar-title">Kaizen</span>
        </span>
    </a>

    <!-- FALECOM -->
    <a href="https://linktr.ee/FaleComVale" class="sidebar-link" target="_blank">
        <span class="sidebar-icon">#</span>
        <span class="sidebar-text">
            <span class="sidebar-title">#FaleCOM</span>
        </span>
    </a>

    <!-- DSS -->
    <a href="dss_online.asp" class="sidebar-link">
        <span class="sidebar-icon">
            <img src="libs/img/img_link/DSSOnline.jpg" onerror="this.style.display='none'">
            <span>D</span>
        </span>
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
            <span class="external-logo external-logo-equipfer"></span>
            <span>Equipfer</span>
        </a>

        <a href="https://gdb.valeglobal.net/gdb/view/login/login.faces" class="external-link" target="_blank">
            <img src="libs/img/img_link/gdb.jpg" alt="GDB">
            <span>GDB</span>
        </a>

        <a href="https://performancemanager4.successfactors.com/sf/home" class="external-link" target="_blank">
            <span class="external-logo external-logo-ves"></span>
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
            <h1 class="page-title">Central de Informações</h1>

            <div class="user-info">
                Bem-vindo(a), <strong><%=Server.HTMLEncode(VAR_USUARIO_NOME)%></strong>
                &nbsp;·&nbsp;
                <%=Server.HTMLEncode(VAR_USUARIO_FUNCAO)%>
            </div>
        </header>

        <section class="summary-row">
            <div class="summary-card">
                <span class="summary-label">Pendências</span>
                <span class="summary-value"><%=VAR_QTD_PENDENCIAS%> pendência(s)</span>
            </div>

            <div class="summary-card">
                <span class="summary-label">Perfil</span>
                <span class="summary-value"><%=Server.HTMLEncode(VAR_USUARIO_NIVEL)%></span>
            </div>

            <div class="summary-card">
                <span class="summary-label">Matrícula</span>
                <span class="summary-value"><%=Server.HTMLEncode(VAR_USUARIO_MATRICULA)%></span>
            </div>
        </section>

        <h2 class="section-title">Últimos registros publicados</h2>

        <section class="records-list">

            <%
                SQL_ULTIMAS = "SELECT TOP 3 " & _
              "TBL_INFORMACAO.COD, " & _
              "TBL_INFORMACAO.DATA, " & _
              "TBL_INFORMACAO.TIPO, " & _
              "TBL_INFORMACAO.TITULO, " & _
              "TBL_INFORMACAO.INFORMACAO, " & _
              "TBL_INFORMACAO.IMG_ID, " & _
              "TBL_USUARIO.NOME, " & _
              "TBL_CHECK_INFORMACAO.[CHECK] AS STATUS_LEITURA " & _
              "FROM (TBL_CHECK_INFORMACAO " & _
              "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD) " & _
              "INNER JOIN TBL_USUARIO ON TBL_INFORMACAO.FK_MATRICULA = TBL_USUARIO.MATRICULA " & _
              "WHERE TBL_CHECK_INFORMACAO.MATRICULA='" & VAR_USUARIO_MATRICULA & "' " & _
              "ORDER BY TBL_INFORMACAO.COD DESC"

                Set RS_ULTIMAS = conexao_.execute(SQL_ULTIMAS)

                If Not RS_ULTIMAS.EOF Then
                    Do Until RS_ULTIMAS.EOF

                        VAR_COD = RS_ULTIMAS("COD")
                        VAR_DATA = RS_ULTIMAS("DATA")
                        VAR_TIPO = RS_ULTIMAS("TIPO")
                        VAR_TITULO = RS_ULTIMAS("TITULO")
                        VAR_INFO = RS_ULTIMAS("INFORMACAO")
                        VAR_IMG = TratarImagem(RS_ULTIMAS("IMG_ID"))
                        VAR_AUTOR = RS_ULTIMAS("NOME")
            %>

            <article class="record-card">
                <div class="record-media">
                    <% If EhImagem(VAR_IMG) Then %>
                        <img src="<%=VAR_IMG%>" alt="Imagem do registro <%=VAR_COD%>" onerror="this.onerror=null;this.src='NOIMG.jpg';">
                    <% Else %>
                        <div class="document-thumb">
                            <span>DOC</span>
                            <span><%=TipoDocumento(VAR_IMG)%></span>
                        </div>
                    <% End If %>
                </div>

                <div class="record-content">
                    <div class="record-top">
                        <h3 class="record-title"><%=TextoCurto(VAR_TITULO, 95)%></h3>
                        <div class="record-cod">#<%=VAR_COD%></div>
                    </div>

                    <div class="record-meta">
                        Tipo: <%=Server.HTMLEncode(VAR_TIPO)%>
                        &nbsp;·&nbsp;
                        Data: <%=VAR_DATA%>
                        &nbsp;·&nbsp;
                        Autor: <%=Server.HTMLEncode(VAR_AUTOR)%>
                    </div>

                    <p class="record-text">
                        <%=TextoCurto(VAR_INFO, 370)%>
                    </p>

                    <div class="record-actions">
                        <a class="record-link" href="form_visualizar_registro.asp">Visualizar registros</a>
                    </div>
                </div>
            </article>

            <%
                        RS_ULTIMAS.movenext
                    Loop
                Else
            %>

                <div class="empty-state">
                    Nenhuma informação cadastrada até o momento.
                </div>

            <%
                End If
            %>

        </section>

    </main>

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
        };
    </script>
</body>
</html>
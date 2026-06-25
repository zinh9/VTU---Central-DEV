<!-- #include file="conexao.asp" -->

<%
If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

matriculaSessao = "" & Session("matricula")

Function HtmlSafe(valor)
    HtmlSafe = Server.HTMLEncode("" & valor)
End Function

Function SqlSafe(valor)
    SqlSafe = Replace(Trim("" & valor), "'", "''")
End Function

Function JsSafe(valor)
    Dim t
    t = "" & valor
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, "'", "\'")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    JsSafe = t
End Function

Function Nz(valor, padrao)
    If IsNull(valor) Then
        Nz = padrao
    Else
        Nz = valor
    End If
End Function

perfilNome = ""
perfilCargo = ""
perfilMatricula = matriculaSessao
perfilCoordId = ""
perfilCoordenacao = ""
perfilGA = ""
perfilGG = ""
perfilDiretoria = "DIRETORIA EFVM"
perfilNivel = "" & Session("nivel")

SQL_PERFIL = "SELECT " & _
             "U.MATRICULA, " & _
             "U.NOME AS PERFIL_NOME, " & _
             "U.CARGO, " & _
             "U.ID_TBL_COORD, " & _
             "L.NOME AS LOGIN_NOME, " & _
             "L.FUNCAO AS LOGIN_FUNCAO, " & _
             "L.NIVEL, " & _
             "C.COORDENACAO, " & _
             "C.[COORDENACAO NOME] AS COORDENACAO_NOME, " & _
             "GA.GA, " & _
             "GA.[GA NOME] AS GA_NOME, " & _
             "GG.GG, " & _
             "GG.[GG NOME] AS GG_NOME " & _
             "FROM (((USUARIO AS U " & _
             "LEFT JOIN TBL_USUARIO AS L ON U.MATRICULA = L.MATRICULA) " & _
             "LEFT JOIN TBL_COORD AS C ON U.ID_TBL_COORD = C.ID) " & _
             "LEFT JOIN TBL_GA AS GA ON C.ID_TBL_GA = GA.ID) " & _
             "LEFT JOIN TBL_GG AS GG ON GA.ID_TBL_GG = GG.ID " & _
             "WHERE U.MATRICULA='" & SqlSafe(matriculaSessao) & "'"

Set RS_PERFIL = conexao_.Execute(SQL_PERFIL)

If Not RS_PERFIL.EOF Then
    perfilNome = "" & Nz(RS_PERFIL("PERFIL_NOME"), "")
    perfilCargo = "" & Nz(RS_PERFIL("CARGO"), "")
    perfilMatricula = "" & Nz(RS_PERFIL("MATRICULA"), matriculaSessao)
    perfilCoordId = "" & Nz(RS_PERFIL("ID_TBL_COORD"), "")
    perfilCoordenacao = "" & Nz(RS_PERFIL("COORDENACAO"), "")
    perfilGA = "" & Nz(RS_PERFIL("GA"), "")
    perfilGG = "" & Nz(RS_PERFIL("GG"), "")
    perfilNivel = "" & Nz(RS_PERFIL("NIVEL"), Session("nivel"))
Else
    SQL_LOGIN = "SELECT * FROM TBL_USUARIO WHERE MATRICULA='" & SqlSafe(matriculaSessao) & "'"
    Set RS_LOGIN = conexao_.Execute(SQL_LOGIN)

    If Not RS_LOGIN.EOF Then
        perfilNome = "" & Nz(RS_LOGIN("NOME"), "")
        perfilCargo = "" & Nz(RS_LOGIN("FUNCAO"), "")
        perfilMatricula = "" & Nz(RS_LOGIN("MATRICULA"), matriculaSessao)
        perfilNivel = "" & Nz(RS_LOGIN("NIVEL"), Session("nivel"))
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Perfil</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        :root {
            --vale-teal: #00857A;
            --vale-teal-dark: #006B63;
            --vale-green: #29D99D;
            --vale-red: #E1261C;
            --text-white: #ffffff;
            --text-dark: #263238;
            --sidebar-width: 132px;
            --glass-bg: rgba(230, 232, 214, 0.48);
            --field-bg: rgba(245, 247, 238, 0.62);
        }

        * {
            box-sizing: border-box;
        }

        html,
        body {
            width: 100%;
            min-height: 100vh;
            margin: 0;
            overflow-x: hidden;
        }

        body {
            font-family: Arial, Helvetica, sans-serif;
            background:
                linear-gradient(rgba(0, 60, 55, 0.08), rgba(0, 60, 55, 0.08)),
                url("libs/img/login_bg.jpg") center center / cover no-repeat fixed;
            color: var(--text-white);
        }

        .sidebar-profile {
            position: fixed;
            left: 0;
            top: 0;
            width: var(--sidebar-width);
            height: 100vh;
            padding: 28px 12px 22px 12px;
            background: rgba(255,255,255,0.28);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            border-radius: 0 0 24px 0;
            box-shadow: 8px 0 24px rgba(0,0,0,0.18);
            display: flex;
            flex-direction: column;
            align-items: center;
            z-index: 20;
        }

        .sidebar-top-links {
            display: flex;
            flex-direction: column;
            gap: 24px;
            align-items: center;
            width: 100%;
        }

        .sidebar-bottom-links {
            margin-top: auto;
            width: 100%;
            padding-top: 18px;
            border-top: 2px solid rgba(255,255,255,0.72);
            display: flex;
            flex-direction: column;
            gap: 16px;
            align-items: center;
        }

        .side-circle,
        .side-action {
            width: 56px;
            height: 56px;
            border-radius: 50%;
            border: none;
            background: rgba(255,255,255,0.96);
            box-shadow: 0 5px 12px rgba(0,0,0,0.24);
            display: flex;
            align-items: center;
            justify-content: center;
            text-decoration: none;
            color: #6E6E6E;
            font-size: 25px;
            font-weight: 900;
            transition: all 0.2s ease-in-out;
        }

        .side-circle:hover,
        .side-action:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 18px rgba(0,0,0,0.28);
        }

        .side-action.profile {
            border-radius: 12px;
            background: rgba(255,255,255,0.55);
            color: #707070;
        }

        .side-action.exit {
            border-radius: 12px;
            background: rgba(255,255,255,0.70);
            color: #B71C1C;
            font-size: 30px;
        }

        .main-profile {
            min-height: 100vh;
            margin-left: var(--sidebar-width);
            padding: 22px 28px 32px 28px;
            position: relative;
        }

        .page-title {
            margin: 0 0 22px 0;
            text-align: center;
            font-size: clamp(42px, 5vw, 68px);
            line-height: 1;
            font-weight: 900;
            letter-spacing: 4px;
            text-transform: uppercase;
            color: #ffffff;
            text-shadow: 0 4px 10px rgba(0,0,0,0.45);
        }

        .profile-form {
            width: min(100%, 1106px);
            margin: 0 auto;
            display: grid;
            grid-template-columns: repeat(12, 1fr);
            gap: 18px 30px;
        }

        .field {
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        .field label {
            font-size: 20px;
            font-weight: 900;
            color: #ffffff;
            text-shadow: 0 3px 8px rgba(0,0,0,0.45);
        }

        .field input,
        .field select {
            width: 100%;
            height: 58px;
            border: none;
            outline: none;
            border-radius: 9px;
            background: var(--field-bg);
            color: #ffffff;
            padding: 0 18px;
            font-size: 24px;
            font-weight: 700;
            text-shadow: 0 2px 5px rgba(0,0,0,0.35);
            box-shadow: inset 0 0 0 1px rgba(255,255,255,0.22);
            backdrop-filter: blur(6px);
            -webkit-backdrop-filter: blur(6px);
        }

        .field select {
            appearance: none;
            -webkit-appearance: none;
            cursor: pointer;
        }

        .field input:focus,
        .field select:focus {
            box-shadow:
                inset 0 0 0 2px rgba(255,255,255,0.65),
                0 0 0 3px rgba(246,184,0,0.30);
            background: rgba(245,247,238,0.72);
        }

        .field input[readonly] {
            opacity: 0.95;
            cursor: not-allowed;
        }

        .field-name {
            grid-column: span 9;
        }

        .field-matricula {
            grid-column: span 3;
        }

        .field-full {
            grid-column: span 12;
        }

        .field-half {
            grid-column: span 6;
        }

        .actions-profile {
            grid-column: span 12;
            margin-top: 76px;
            display: flex;
            justify-content: flex-end;
            align-items: center;
            gap: 18px;
        }

        .btn-password,
        .btn-confirm,
        .btn-cancel,
        .btn-save {
            min-height: 72px;
            border: none;
            border-radius: 9px;
            color: #ffffff;
            font-size: clamp(24px, 2.2vw, 34px);
            font-weight: 900;
            text-shadow: 0 3px 8px rgba(0,0,0,0.38);
            cursor: pointer;
            transition: all 0.2s ease-in-out;
            box-shadow: 0 10px 20px rgba(0,0,0,0.22);
        }

        .btn-password {
            min-width: 230px;
            padding: 0 22px;
            background: rgba(230,232,214,0.66);
        }

        .btn-confirm,
        .btn-save {
            min-width: 230px;
            padding: 0 30px;
            background: var(--vale-green);
        }

        .btn-cancel {
            min-width: 230px;
            padding: 0 30px;
            background: var(--vale-red);
        }

        .btn-password:hover,
        .btn-confirm:hover,
        .btn-cancel:hover,
        .btn-save:hover {
            transform: translateY(-2px);
            box-shadow: 0 14px 28px rgba(0,0,0,0.28);
        }

        .alert-box {
            width: min(100%, 1106px);
            margin: 0 auto 18px auto;
            padding: 14px 18px;
            border-radius: 14px;
            background: rgba(0, 133, 122, 0.82);
            color: #ffffff;
            font-weight: 900;
            text-align: center;
            box-shadow: 0 10px 22px rgba(0,0,0,0.20);
        }

        .modal-overlay {
            display: none;
            position: fixed;
            inset: 0;
            z-index: 1000;
            background: rgba(0,0,0,0.10);
            align-items: center;
            justify-content: center;
            padding: 26px;
        }

        .modal-overlay.open {
            display: flex;
        }

        .modal-card {
            width: min(100%, 1084px);
            border-radius: 22px;
            padding: 24px 22px 36px 22px;
            background: var(--glass-bg);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            box-shadow:
                inset 0 0 0 1px rgba(255,255,255,0.38),
                0 18px 42px rgba(0,0,0,0.30);
        }

        .modal-title {
            margin: 0 0 34px 0;
            text-align: center;
            color: #ffffff;
            font-size: clamp(30px, 3.2vw, 42px);
            font-weight: 900;
            text-transform: uppercase;
            text-shadow: 0 4px 9px rgba(0,0,0,0.45);
        }

        .password-grid {
            display: grid;
            grid-template-columns: 160px 1fr;
            gap: 20px 12px;
            align-items: center;
        }

        .password-grid label {
            color: #ffffff;
            font-size: 24px;
            font-weight: 900;
            text-shadow: 0 3px 8px rgba(0,0,0,0.48);
        }

        .password-grid input {
            height: 58px;
            border: none;
            outline: none;
            border-radius: 9px;
            background: var(--field-bg);
            color: #ffffff;
            padding: 0 18px;
            font-size: 24px;
            font-weight: 700;
        }

        .password-grid input:focus {
            box-shadow:
                inset 0 0 0 2px rgba(255,255,255,0.65),
                0 0 0 3px rgba(246,184,0,0.30);
        }

        .modal-actions {
            margin-top: 62px;
            display: flex;
            justify-content: center;
            gap: 70px;
        }

        .confirmation-text {
            width: min(100%, 860px);
            min-height: 166px;
            margin: 0 auto;
            border-radius: 8px;
            background: rgba(245,247,238,0.62);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px 54px;
            color: #ffffff;
            font-size: clamp(26px, 3vw, 38px);
            line-height: 1.05;
            text-align: left;
            text-shadow: 0 3px 8px rgba(0,0,0,0.45);
        }

        @media (max-width: 920px) {
            :root {
                --sidebar-width: 88px;
            }

            .sidebar-profile {
                width: var(--sidebar-width);
            }

            .main-profile {
                margin-left: var(--sidebar-width);
                padding: 18px;
            }

            .field-name,
            .field-matricula,
            .field-half,
            .field-full {
                grid-column: span 12;
            }

            .actions-profile {
                margin-top: 32px;
                flex-direction: column;
                align-items: stretch;
            }

            .btn-password,
            .btn-confirm,
            .btn-cancel,
            .btn-save {
                width: 100%;
            }

            .password-grid {
                grid-template-columns: 1fr;
            }

            .modal-actions {
                gap: 16px;
                flex-direction: column;
            }

            .confirmation-text {
                padding: 22px;
                font-size: 24px;
            }
        }

        @media (max-width: 560px) {
            .side-circle,
            .side-action {
                width: 48px;
                height: 48px;
            }

            .field input,
            .field select {
                font-size: 18px;
            }

            .field label {
                font-size: 17px;
            }

            .page-title {
                font-size: 38px;
            }
        }
    </style>
</head>

<body>

    <aside class="sidebar-profile">

        <div class="sidebar-top-links">
            <a href="form_home.asp" class="side-circle" title="Home">⌂</a>

            <a href="form_visualizar_registro.asp" class="side-circle" title="Visualizar registros">#</a>

            <% If Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM" Then %>
                <a href="form_formulario.asp" class="side-circle" title="Novo registro">+</a>
            <% End If %>

            <% If Session("funcao") = "INSPETOR" Or Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM" Then %>
                <a href="form_grafico.asp" class="side-circle" title="Relatórios">▥</a>
            <% End If %>

            <a href="dss_online.asp" class="side-circle" title="DSS Online">D</a>
        </div>

        <div class="sidebar-bottom-links">
            <a href="form_perfil.asp" class="side-action profile" title="Perfil">◎</a>
            <a href="fecha_session.asp" class="side-action exit" title="Sair">↪</a>
        </div>

    </aside>

    <main class="main-profile">

        <h1 class="page-title">Perfil de Usuário</h1>

        <% If Session("msg_perfil") <> "" Then %>
            <div class="alert-box">
                <%=HtmlSafe(Session("msg_perfil"))%>
            </div>
            <%
                Session("msg_perfil") = ""
            %>
        <% End If %>

        <form id="formPerfil" class="profile-form" method="post" action="valida_perfil.asp">

            <input type="hidden" name="inputMatricula" value="<%=HtmlSafe(perfilMatricula)%>">

            <div class="field field-name">
                <label for="inputNome">Nome</label>
                <input
                    type="text"
                    id="inputNome"
                    name="inputNome"
                    value="<%=HtmlSafe(perfilNome)%>"
                    required>
            </div>

            <div class="field field-matricula">
                <label for="inputMatriculaVisual">Matrícula</label>
                <input
                    type="text"
                    id="inputMatriculaVisual"
                    value="<%=HtmlSafe(perfilMatricula)%>"
                    readonly>
            </div>

            <div class="field field-full">
                <label for="inputFuncao">Função</label>
                <input
                    type="text"
                    id="inputFuncao"
                    name="inputFuncao"
                    value="<%=HtmlSafe(perfilCargo)%>"
                    required>
            </div>

            <div class="field field-half">
                <label for="inputCoord">Coordenação</label>
                <select id="inputCoord" name="inputCoord" onchange="atualizarHierarquia()" required>
                    <option value=""></option>

                    <%
                    SQL_COORD = "SELECT " & _
                                "C.ID, " & _
                                "C.COORDENACAO, " & _
                                "GA.GA, " & _
                                "GG.GG " & _
                                "FROM (TBL_COORD AS C " & _
                                "LEFT JOIN TBL_GA AS GA ON C.ID_TBL_GA = GA.ID) " & _
                                "LEFT JOIN TBL_GG AS GG ON GA.ID_TBL_GG = GG.ID " & _
                                "ORDER BY C.COORDENACAO"

                    Set RS_COORD = conexao_.Execute(SQL_COORD)

                    Do Until RS_COORD.EOF
                        coordIdOption = "" & RS_COORD("ID")
                        selectedCoord = ""

                        If CStr(coordIdOption) = CStr(perfilCoordId) Then
                            selectedCoord = "selected"
                        End If
                    %>
                        <option value="<%=HtmlSafe(coordIdOption)%>" <%=selectedCoord%>>
                            <%=HtmlSafe(RS_COORD("COORDENACAO"))%>
                        </option>
                    <%
                        RS_COORD.MoveNext
                    Loop
                    %>
                </select>
            </div>

            <div class="field field-half">
                <label for="inputGA">Gerência de área</label>
                <input
                    type="text"
                    id="inputGA"
                    value="<%=HtmlSafe(perfilGA)%>"
                    readonly>
            </div>

            <div class="field field-half">
                <label for="inputGG">Gerência Geral</label>
                <input
                    type="text"
                    id="inputGG"
                    value="<%=HtmlSafe(perfilGG)%>"
                    readonly>
            </div>

            <div class="field field-half">
                <label for="inputDiretoria">Diretoria</label>
                <input
                    type="text"
                    id="inputDiretoria"
                    value="<%=HtmlSafe(perfilDiretoria)%>"
                    readonly>
            </div>

            <div class="actions-profile">
                <button type="button" class="btn-password" onclick="abrirModalSenha()">
                    Alterar Senha
                </button>

                <button type="button" class="btn-confirm" onclick="abrirModalConfirmacao()">
                    Confirmar
                </button>
            </div>

        </form>

    </main>

    <!-- MODAL TROCA DE SENHA -->
    <div id="modalSenha" class="modal-overlay">
        <div class="modal-card">
            <h2 class="modal-title">Troca de Senha</h2>

            <form id="formSenha" method="post" action="valida_trocar_senha.asp">

                <div class="password-grid">
                    <label for="senhaAtual">Senha atual:</label>
                    <input type="password" id="senhaAtual" name="senhaAtual" required>

                    <label for="novaSenha">Nova senha:</label>
                    <input type="password" id="novaSenha" name="novaSenha" required>

                    <label for="confirmaSenha">Confirmação:</label>
                    <input type="password" id="confirmaSenha" name="confirmaSenha" required>
                </div>

                <div class="modal-actions">
                    <button type="button" class="btn-cancel" onclick="fecharModalSenha()">
                        Cancelar
                    </button>

                    <button type="submit" class="btn-save">
                        Salvar
                    </button>
                </div>

            </form>
        </div>
    </div>

    <!-- MODAL CONFIRMAÇÃO DE PERFIL -->
    <div id="modalConfirmacao" class="modal-overlay">
        <div class="modal-card">
            <h2 class="modal-title">Confirmação de Alteração de Dados</h2>

            <div class="confirmation-text">
                Confirmo a atualização de dados do meu perfil e informo que todas alterações realizadas são verídicas e estão preenchidas corretamente.
            </div>

            <div class="modal-actions">
                <button type="button" class="btn-cancel" onclick="fecharModalConfirmacao()">
                    Cancelar
                </button>

                <button type="button" class="btn-save" onclick="enviarPerfil()">
                    Confirmar
                </button>
            </div>
        </div>
    </div>

    <script>
        var coordMap = {};

        <%
        Set RS_COORD_JS = conexao_.Execute(SQL_COORD)

        Do Until RS_COORD_JS.EOF
            jsCoordId = "" & RS_COORD_JS("ID")
            jsGA = "" & Nz(RS_COORD_JS("GA"), "")
            jsGG = "" & Nz(RS_COORD_JS("GG"), "")
        %>
            coordMap["<%=JsSafe(jsCoordId)%>"] = {
                ga: "<%=JsSafe(jsGA)%>",
                gg: "<%=JsSafe(jsGG)%>",
                diretoria: "DIRETORIA EFVM"
            };
        <%
            RS_COORD_JS.MoveNext
        Loop
        %>

        function atualizarHierarquia() {
            var coordSelect = document.getElementById("inputCoord");
            var gaInput = document.getElementById("inputGA");
            var ggInput = document.getElementById("inputGG");
            var diretoriaInput = document.getElementById("inputDiretoria");

            var coordId = coordSelect.value;

            if (coordMap[coordId]) {
                gaInput.value = coordMap[coordId].ga;
                ggInput.value = coordMap[coordId].gg;
                diretoriaInput.value = coordMap[coordId].diretoria;
            } else {
                gaInput.value = "";
                ggInput.value = "";
                diretoriaInput.value = "DIRETORIA EFVM";
            }
        }

        function abrirModalSenha() {
            document.getElementById("modalSenha").classList.add("open");
        }

        function fecharModalSenha() {
            document.getElementById("modalSenha").classList.remove("open");
        }

        function abrirModalConfirmacao() {
            document.getElementById("modalConfirmacao").classList.add("open");
        }

        function fecharModalConfirmacao() {
            document.getElementById("modalConfirmacao").classList.remove("open");
        }

        function enviarPerfil() {
            document.getElementById("formPerfil").submit();
        }

        document.addEventListener("keydown", function(event) {
            if (event.key === "Escape") {
                fecharModalSenha();
                fecharModalConfirmacao();
            }
        });
    </script>

</body>
</html>
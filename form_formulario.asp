<!-- #include file="conexao.asp" -->

<%
If Session("matricula") <> "" Then

    If Session("funcao") = "INSPETOR" Or Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM" Then

        usuarioNome = "" & Session("name")
        usuarioMatricula = "" & Session("matricula")
        usuarioFuncao = "" & Session("funcao")
        usuarioNivel = "" & Session("nivel")

        dataHoje = Right("0" & Day(Date), 2) & "/" & Right("0" & Month(Date), 2) & "/" & Year(Date)

        Function HtmlSafe(valor)
            HtmlSafe = Server.HTMLEncode("" & valor)
        End Function

        Function IdSafe(valor)
            Dim t
            t = "" & valor
            t = Replace(t, " ", "_")
            t = Replace(t, "Ã", "A")
            t = Replace(t, "Á", "A")
            t = Replace(t, "À", "A")
            t = Replace(t, "Â", "A")
            t = Replace(t, "É", "E")
            t = Replace(t, "Ê", "E")
            t = Replace(t, "Í", "I")
            t = Replace(t, "Ó", "O")
            t = Replace(t, "Ô", "O")
            t = Replace(t, "Ú", "U")
            t = Replace(t, "Ç", "C")
            t = Replace(t, ".", "")
            t = Replace(t, "/", "")
            t = Replace(t, "\", "")
            t = Replace(t, "-", "_")
            IdSafe = "func_" & t
        End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Inserir Registro</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        :root {
            --vale-teal: #00857A;
            --vale-teal-dark: #006B63;
            --vale-yellow: #F6B800;
            --text-dark: #263238;
            --text-muted: #607D8B;
            --border-gray: #B0BEC5;
            --field-bg: #F7FAFA;
            --card-bg: rgba(255, 255, 255, 0.94);
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
                linear-gradient(rgba(0, 60, 55, 0.18), rgba(0, 60, 55, 0.18)),
                url("libs/img/login_bg.jpg") center center / cover no-repeat fixed;
            color: var(--text-dark);
        }

        .page-wrapper {
            width: 100%;
            min-height: 100vh;
            padding: clamp(18px, 3vw, 42px);
        }

        .form-card {
            width: min(100%, 1080px);
            margin: 0 auto;
            padding: clamp(22px, 3vw, 38px);
            border-radius: 28px;
            background: var(--card-bg);
            box-shadow: 0 22px 55px rgba(0, 0, 0, 0.25);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
        }

        .form-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 18px;
            margin-bottom: 28px;
            border-bottom: 1px solid rgba(0, 133, 122, 0.18);
            padding-bottom: 18px;
        }

        .form-title-area {
            flex: 1;
        }

        .form-title {
            margin: 0;
            color: var(--vale-teal-dark);
            font-size: clamp(26px, 3.2vw, 42px);
            font-weight: 900;
            line-height: 1.1;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .form-subtitle {
            margin-top: 6px;
            color: var(--text-muted);
            font-size: 14px;
            font-weight: 700;
        }

        .btn-back {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-height: 42px;
            padding: 0 18px;
            border-radius: 12px;
            background: #ECEFF1;
            color: var(--text-dark);
            text-decoration: none;
            font-size: 14px;
            font-weight: 800;
            transition: all 0.2s ease-in-out;
            white-space: nowrap;
        }

        .btn-back:hover {
            background: #ffffff;
            transform: translateY(-1px);
            box-shadow: 0 8px 18px rgba(0,0,0,0.14);
        }

        .alert-box {
            margin-bottom: 20px;
            padding: 14px 18px;
            border-radius: 14px;
            background: rgba(198, 40, 40, 0.92);
            color: #ffffff;
            font-weight: 800;
            text-align: center;
        }

        .section-title {
            margin: 26px 0 16px 0;
            color: var(--vale-teal-dark);
            font-size: 18px;
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }

        .form-grid {
            display: grid;
            grid-template-columns: repeat(12, 1fr);
            gap: 18px;
        }

        .field {
            display: flex;
            flex-direction: column;
            gap: 7px;
        }

        .field-12 {
            grid-column: span 12;
        }

        .field-8 {
            grid-column: span 8;
        }

        .field-6 {
            grid-column: span 6;
        }

        .field-4 {
            grid-column: span 4;
        }

        .field-3 {
            grid-column: span 3;
        }

        .field label {
            color: #37474F;
            font-size: 13px;
            font-weight: 900;
        }

        .form-control-custom {
            width: 100%;
            min-height: 44px;
            padding: 10px 13px;
            border: 1.5px solid var(--border-gray);
            border-radius: 12px;
            background: var(--field-bg);
            color: var(--text-dark);
            font-size: 14px;
            outline: none;
            transition: all 0.18s ease-in-out;
        }

        .form-control-custom:focus {
            border-color: var(--vale-yellow);
            background: #ffffff;
            box-shadow: 0 0 0 3px rgba(246, 184, 0, 0.32);
        }

        .form-control-custom[readonly] {
            background: #ECEFF1;
            color: #546E7A;
            cursor: not-allowed;
        }

        textarea.form-control-custom {
            min-height: 160px;
            resize: vertical;
            line-height: 1.5;
        }

        .tabs-card {
            margin-top: 10px;
            border: 1.5px solid rgba(176, 190, 197, 0.85);
            border-radius: 20px;
            overflow: hidden;
            background: rgba(255,255,255,0.62);
        }

        .tabs-nav {
            display: flex;
            background: rgba(0, 133, 122, 0.08);
            border-bottom: 1px solid rgba(176, 190, 197, 0.75);
        }

        .tab-btn {
            flex: 1;
            min-height: 48px;
            border: none;
            background: transparent;
            color: var(--vale-teal-dark);
            font-weight: 900;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
        }

        .tab-btn.active {
            background: var(--vale-teal);
            color: #ffffff;
        }

        .tab-btn:hover {
            background: rgba(0, 133, 122, 0.16);
        }

        .tab-btn.active:hover {
            background: var(--vale-teal-dark);
        }

        .tab-panel {
            display: none;
            padding: 20px;
        }

        .tab-panel.active {
            display: block;
        }

        .checkbox-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(180px, 1fr));
            gap: 12px;
        }

        .checkbox-item {
            display: flex;
            align-items: center;
            gap: 10px;
            min-height: 42px;
            padding: 10px 12px;
            border: 1.5px solid var(--border-gray);
            border-radius: 12px;
            background: #ffffff;
            cursor: pointer;
            transition: all 0.18s ease-in-out;
            font-size: 14px;
            font-weight: 700;
            color: #37474F;
        }

        .checkbox-item:hover {
            border-color: var(--vale-yellow);
            box-shadow: 0 0 0 3px rgba(246, 184, 0, 0.20);
        }

        .checkbox-item input {
            width: 17px;
            height: 17px;
            accent-color: var(--vale-teal);
        }

        .form-actions {
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            margin-top: 28px;
            padding-top: 22px;
            border-top: 1px solid rgba(0, 133, 122, 0.18);
        }

        .btn-primary-custom,
        .btn-secondary-custom {
            min-height: 46px;
            padding: 0 22px;
            border: none;
            border-radius: 12px;
            font-size: 15px;
            font-weight: 900;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
        }

        .btn-primary-custom {
            background: var(--vale-teal);
            color: #ffffff;
        }

        .btn-primary-custom:hover {
            background: var(--vale-teal-dark);
            transform: translateY(-1px);
            box-shadow: 0 10px 20px rgba(0, 107, 99, 0.25);
        }

        .btn-secondary-custom {
            background: #ECEFF1;
            color: var(--text-dark);
        }

        .btn-secondary-custom:hover {
            background: #ffffff;
            transform: translateY(-1px);
            box-shadow: 0 8px 18px rgba(0,0,0,0.14);
        }

        .help-text {
            margin-top: 8px;
            color: var(--text-muted);
            font-size: 12px;
            line-height: 1.4;
        }

        @media (max-width: 980px) {
            .form-header {
                flex-direction: column;
                align-items: stretch;
            }

            .btn-back {
                width: 100%;
            }

            .field-8,
            .field-6,
            .field-4,
            .field-3 {
                grid-column: span 12;
            }

            .checkbox-grid {
                grid-template-columns: repeat(2, minmax(160px, 1fr));
            }
        }

        @media (max-width: 620px) {
            .page-wrapper {
                padding: 14px;
            }

            .form-card {
                padding: 20px;
                border-radius: 22px;
            }

            .tabs-nav {
                flex-direction: column;
            }

            .checkbox-grid {
                grid-template-columns: 1fr;
            }

            .form-actions {
                flex-direction: column;
            }

            .btn-primary-custom,
            .btn-secondary-custom {
                width: 100%;
            }
        }

		/* =====================================================
   CORREÇÃO DOS BOTÕES DO FORMULÁRIO
===================================================== */

.form-actions {
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 12px;
    margin-top: 28px;
    padding-top: 22px;
    border-top: 1px solid rgba(0, 133, 122, 0.18);
}

.btn-primary-custom,
.btn-secondary-custom {
    min-height: 48px;
    min-width: 140px;
    padding: 0 24px;
    border-radius: 12px;
    border: none;
    font-size: 15px;
    font-weight: 900;
    cursor: pointer;
    text-align: center;
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    justify-content: center;
}

.btn-primary-custom {
    background: #00857A !important;
    color: #ffffff !important;
    box-shadow: 0 10px 22px rgba(0, 133, 122, 0.25);
}

.btn-primary-custom:hover {
    background: #006B63 !important;
    transform: translateY(-1px);
}

.btn-secondary-custom {
    background: #ECEFF1 !important;
    color: #263238 !important;
}

.btn-secondary-custom:hover {
    background: #ffffff !important;
    transform: translateY(-1px);
    box-shadow: 0 8px 18px rgba(0,0,0,0.14);
}

@media (max-width: 620px) {
    .form-actions {
        flex-direction: column;
    }

    .btn-primary-custom,
    .btn-secondary-custom {
        width: 100%;
    }
}

/* =====================================================
   CAMPO DE ANEXO / FOTO
===================================================== */

.upload-box {
    position: relative;
    min-height: 96px;
    border: 2px dashed #B0BEC5;
    border-radius: 16px;
    background: #F7FAFA;
    display: flex;
    align-items: center;
    gap: 16px;
    padding: 18px;
    transition: all 0.18s ease-in-out;
    cursor: pointer;
}

.upload-box:hover,
.upload-box:focus-within {
    border-color: #F6B800;
    background: #ffffff;
    box-shadow: 0 0 0 3px rgba(246, 184, 0, 0.24);
}

.upload-icon {
    width: 52px;
    height: 52px;
    border-radius: 14px;
    background: #00857A;
    color: #ffffff;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    flex-shrink: 0;
}

.upload-content {
    display: flex;
    flex-direction: column;
    gap: 4px;
    color: #263238;
}

.upload-content strong {
    font-size: 15px;
    color: #006B63;
}

.upload-content span {
    font-size: 13px;
    color: #607D8B;
}

.upload-warning {
    font-size: 12px !important;
    color: #8A6D00 !important;
}

.upload-input {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
}
    </style>
</head>

<body>
    <div class="page-wrapper">
        <main class="form-card">

            <div class="form-header">
                <div class="form-title-area">
                    <h1 class="form-title">Inserir informação</h1>
                    <div class="form-subtitle">
                        Central de Informações · Cadastro de novo registro
                    </div>
                </div>

                <a href="form_home.asp" class="btn-back">← Voltar</a>
            </div>

            <%
            If Session("alerta") <> "" Then
            %>
                <div class="alert-box">
                    <%=HtmlSafe(Session("alerta"))%>
                </div>
            <%
                Session("alerta") = ""
            End If
            %>

            <form name="form" id="form" action="valida_formulario.asp" method="POST">

                <div class="section-title">Dados do solicitante</div>

                <div class="form-grid">
                    <div class="field field-6">
                        <label for="inputNome">Nome</label>
                        <input id="inputNome" name="inputNome" class="form-control-custom" value="<%=HtmlSafe(usuarioNome)%>" type="text" required readonly>
                    </div>

                    <div class="field field-3">
                        <label for="inputMatricula">Matrícula</label>
                        <input id="inputMatricula" name="inputMatricula" class="form-control-custom" value="<%=HtmlSafe(usuarioMatricula)%>" type="text" required readonly>
                    </div>

                    <div class="field field-3">
                        <label for="inputData">Data</label>
                        <input id="inputData" name="inputData" class="form-control-custom" value="<%=dataHoje%>" type="text" required readonly>
                    </div>
                </div>

                <div class="section-title">Classificação</div>

                <div class="form-grid">
                    <div class="field field-6">
                        <label for="inputTipo">Tipo</label>
                        <select class="form-control-custom" id="inputTipo" name="inputTipo" required>
                            <option value=""></option>
                            <option value="GESTÃO">Gestão</option>
                            <option value="SAÚDE">Saúde</option>
                            <option value="ALERTA">Alerta</option>
                            <option value="MEIO AMBIENTE">Meio Ambiente</option>
                            <option value="SEGUNRANÇA">Segurança</option>
                            <option value="RECONHECIMENTO">Reconhecimento</option>
                            <option value="INSTRUÇÃO TÉCNICA OPERACIONAL">Instrução Técnica Operacional</option>
                            <option value="CAPACITAÇÃO">Capacitação</option>
                        </select>
                    </div>
                </div>

                <div class="section-title">Público destinatário</div>

                <div class="tabs-card">
                    <div class="tabs-nav">
                        <button type="button" class="tab-btn active" onclick="abrirAba('geral', this)">Geral</button>
                        <button type="button" class="tab-btn" onclick="abrirAba('funcionario', this)">Funcionário específico</button>
                    </div>

                    <div id="tab-geral" class="tab-panel active">
                        <div class="checkbox-grid">
                            <%
                            SQL = "SELECT TBL_USUARIO.FUNCAO FROM TBL_USUARIO GROUP BY TBL_USUARIO.FUNCAO ORDER BY TBL_USUARIO.FUNCAO;"
                            Set R_SQL = conexao_.Execute(SQL)

                            Do Until R_SQL.EOF
                                funcaoValor = "" & R_SQL("FUNCAO")
                                funcaoId = IdSafe(funcaoValor)
                            %>
                                <label class="checkbox-item" for="<%=funcaoId%>">
                                    <input id="<%=funcaoId%>" name="<%=HtmlSafe(funcaoValor)%>" type="checkbox" value="<%=HtmlSafe(funcaoValor)%>">
                                    <span><%=HtmlSafe(funcaoValor)%></span>
                                </label>
                            <%
                                R_SQL.MoveNext
                            Loop
                            %>
                        </div>

                        <div class="help-text">
                            Selecione uma ou mais funções para direcionar a informação ao público geral.
                        </div>
                    </div>

                    <div id="tab-funcionario" class="tab-panel">
                        <div class="field field-12">
                            <label for="inputFuncionarioEspecifico">Empregado</label>
                            <select class="form-control-custom" id="inputFuncionarioEspecifico" name="inputFuncionarioEspecifico">
                                <option value=""></option>

                                <%
                                SQL = "SELECT * FROM TBL_USUARIO ORDER BY TBL_USUARIO.NOME;"
                                Set R_SQL = conexao_.Execute(SQL)

                                Do Until R_SQL.EOF
                                %>
                                    <option value="<%=HtmlSafe(R_SQL("MATRICULA"))%>">
                                        <%=HtmlSafe(R_SQL("NOME"))%> - <%=HtmlSafe(R_SQL("FUNCAO"))%>
                                    </option>
                                <%
                                    R_SQL.MoveNext
                                Loop
                                %>
                            </select>

                            <div class="help-text">
                                Use esta opção caso a informação seja destinada a apenas um empregado específico.
                            </div>
                        </div>
                    </div>
                </div>

                <div class="section-title">Conteúdo da informação</div>

<div class="form-grid">

    <div class="field field-8">
        <label for="inputTitulo">Título</label>
        <input
            id="inputTitulo"
            name="inputTitulo"
            class="form-control-custom"
            value=""
            type="text"
            required>
    </div>

    <div class="field field-12">
        <label for="inputInformacao">Informação</label>
        <textarea
            class="form-control-custom"
            placeholder=""
            name="inputInformacao"
            id="inputInformacao"
            required></textarea>
    </div>

    <div class="field field-12">
        <label for="inputAnexoVisual">Anexar foto ou evidência</label>

        <div class="upload-box">
            <div class="upload-icon">📎</div>

            <div class="upload-content">
                <strong>Selecionar arquivo</strong>
                <span>Formatos recomendados: JPG, PNG, PDF ou PPTX.</span>
                <span class="upload-warning">
                    Após inserir a informação, o sistema poderá direcionar para a etapa de upload do anexo.
                </span>
                <span id="nomeArquivoSelecionado" class="upload-warning"></span>
            </div>

            <input
                id="inputAnexoVisual"
                name="inputAnexoVisual"
                type="file"
                class="upload-input"
                accept=".jpg,.jpeg,.png,.pdf,.ppt,.pptx"
                onchange="mostrarNomeArquivo(this)">
        </div>
    </div>

</div>

<div class="form-actions">
    form_home.asp
        Cancelar
    </button>

    <button
        class="btn-primary-custom"
        type="submit">
        Inserir
    </button>
</div>

</form>

    <script>
    function abrirAba(nomeAba, botao) {
        var paineis = document.querySelectorAll(".tab-panel");
        var botoes = document.querySelectorAll(".tab-btn");

        for (var i = 0; i < paineis.length; i++) {
            paineis[i].classList.remove("active");
        }

        for (var j = 0; j < botoes.length; j++) {
            botoes[j].classList.remove("active");
        }

        document.getElementById("tab-" + nomeAba).classList.add("active");
        botao.classList.add("active");
    }

    function mostrarNomeArquivo(input) {
        var label = document.getElementById("nomeArquivoSelecionado");

        if (input.files && input.files.length > 0) {
            label.innerHTML = "Arquivo selecionado: " + input.files[0].name;
        } else {
            label.innerHTML = "";
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
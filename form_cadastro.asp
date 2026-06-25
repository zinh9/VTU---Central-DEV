<!-- #include file="conexao.asp" -->

<%
VAR_ALERTA = Session("alerta")
Session("alerta") = ""
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Cadastro</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            min-height: 100vh;
            font-family: Arial, Helvetica, sans-serif;
            background:
                linear-gradient(rgba(0, 60, 55, 0.25), rgba(0, 60, 55, 0.25)),
                url("libs/img/login_bg.jpg") center center / cover no-repeat;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .cadastro-wrapper {
            width: 100%;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 32px;
        }

        .cadastro-card {
            width: 100%;
            max-width: 660px;
            padding: 40px 56px;
            border-radius: 28px;
            background: rgba(255, 255, 255, 0.54);
            box-shadow: 0 22px 55px rgba(0, 0, 0, 0.28);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
        }

        .cadastro-title {
            margin: 0 0 30px 0;
            text-align: center;
            color: #ffffff;
            font-size: 30px;
            font-weight: 700;
            letter-spacing: 1px;
            text-transform: uppercase;
            text-shadow: 0 2px 7px rgba(0, 0, 0, 0.42);
        }

        .form-group {
            margin-bottom: 18px;
        }

        .form-label {
            display: block;
            margin-bottom: 7px;
            color: #ffffff;
            font-size: 16px;
            font-weight: 700;
            text-shadow: 0 1px 5px rgba(0, 0, 0, 0.38);
        }

        .form-control {
            width: 100%;
            height: 48px;
            padding: 0 16px;
            border: none;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.88);
            color: #333333;
            font-size: 16px;
            outline: none;
        }

        .form-control::placeholder {
            color: #777777;
        }

        .form-control:focus {
            background: rgba(255, 255, 255, 0.98);
            box-shadow: 0 0 0 3px rgba(246, 184, 0, 0.75);
        }

        .form-row {
            display: flex;
            gap: 18px;
        }

        .form-row .form-group {
            width: 100%;
        }

        .btn-cadastro {
            display: block;
            width: 210px;
            height: 50px;
            margin: 30px auto 24px auto;
            border: none;
            border-radius: 10px;
            background: #00857A;
            color: #ffffff;
            font-size: 17px;
            font-weight: 700;
            text-transform: uppercase;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
        }

        .btn-cadastro:hover {
            background: #006B63;
            transform: translateY(-1px);
        }

        .btn-cadastro:focus {
            outline: 3px solid rgba(246, 184, 0, 0.85);
            outline-offset: 3px;
        }

        .cadastro-links {
            margin-top: 6px;
            text-align: left;
        }

        .cadastro-links a {
            display: block;
            margin-bottom: 7px;
            color: #ffffff;
            font-size: 15px;
            text-decoration: none;
            text-shadow: 0 1px 5px rgba(0, 0, 0, 0.45);
        }

        .cadastro-links a:hover {
            text-decoration: underline;
        }

        .alerta-cadastro {
            margin-bottom: 24px;
            padding: 13px 16px;
            border-radius: 8px;
            background: rgba(198, 40, 40, 0.92);
            color: #ffffff;
            font-size: 15px;
            font-weight: 700;
            text-align: center;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.18);
        }

        
         .info-cadastro {
         margin-top: 18px;
         color: #ffffff;
         font-size: 13px;
         line-height: 1.45;
         text-align: center;
         text-shadow: 0 1px 4px rgba(0, 0, 0, 0.35);
         }


        @media (max-width: 768px) {
            .cadastro-wrapper {
                padding: 20px;
            }

            .cadastro-card {
                max-width: 100%;
                padding: 32px 28px;
                border-radius: 22px;
            }

            .cadastro-title {
                font-size: 24px;
            }

            .form-row {
                display: block;
            }

            .btn-cadastro {
                width: 100%;
            }
        }
        /* =====================================================
   RESPONSIVIDADE - TELA DE CADASTRO
   Ajusta o layout para notebook, monitores menores e telas grandes
===================================================== */

html, body {
    width: 100%;
    min-height: 100vh;
    min-height: 100svh;
    overflow-x: hidden;
}

body {
    background-size: cover;
    background-position: center center;
}

.cadastro-wrapper {
    min-height: 100vh;
    min-height: 100svh;
    padding: clamp(12px, 2vw, 32px);
    align-items: center;
}

.cadastro-card {
    width: min(92vw, 660px);
    max-height: calc(100svh - 32px);
    overflow-y: auto;
    padding:
        clamp(22px, 3vw, 40px)
        clamp(24px, 4vw, 56px);
    border-radius: clamp(18px, 2vw, 28px);
}

.cadastro-title {
    font-size: clamp(22px, 2.4vw, 30px);
    margin-bottom: clamp(18px, 3vh, 30px);
}

.form-group {
    margin-bottom: clamp(12px, 1.8vh, 18px);
}

.form-label {
    font-size: clamp(13px, 1.2vw, 16px);
}

.form-control {
    height: clamp(40px, 5.4vh, 48px);
    font-size: clamp(14px, 1.2vw, 16px);
}

.btn-cadastro {
    width: min(100%, 210px);
    height: clamp(42px, 5.5vh, 50px);
    margin:
        clamp(18px, 3vh, 30px)
        auto
        clamp(12px, 2vh, 24px)
        auto;
}

/* Telas médias: notebook e monitores menores */
@media (max-width: 1100px) {
    .cadastro-card {
        width: min(88vw, 620px);
    }
}

/* Telas menores: empilha campos lado a lado */
@media (max-width: 900px) {
    .form-row {
        display: block;
    }

    .cadastro-card {
        width: min(92vw, 560px);
    }

    .btn-cadastro {
        width: 100%;
    }
}

/* Notebooks com pouca altura vertical */
@media (max-height: 760px) {
    .cadastro-wrapper {
        align-items: flex-start;
        padding-top: 18px;
        padding-bottom: 18px;
    }

    .cadastro-card {
        max-height: calc(100svh - 36px);
        padding: 22px 36px;
    }

    .cadastro-title {
        font-size: 24px;
        margin-bottom: 18px;
    }

    .form-group {
        margin-bottom: 11px;
    }

    .form-control {
        height: 40px;
    }

    .btn-cadastro {
        height: 42px;
        margin-top: 18px;
        margin-bottom: 12px;
    }

    .info-cadastro {
        display: none;
    }
}

/* Telas muito baixas */
@media (max-height: 640px) {
    .cadastro-card {
        padding: 18px 30px;
    }

    .cadastro-title {
        font-size: 22px;
        margin-bottom: 14px;
    }

    .form-label {
        margin-bottom: 4px;
    }

    .form-control {
        height: 38px;
    }
}
/* =====================================================
   CORREÇÃO DE CENTRALIZAÇÃO - CADASTRO
===================================================== */

html,
body {
    width: 100%;
    min-height: 100vh;
    min-height: 100dvh;
    margin: 0;
    overflow-x: hidden;
}

body {
    display: flex;
    align-items: center;
    justify-content: center;
}

.cadastro-wrapper {
    width: 100%;
    min-height: 100vh;
    min-height: 100dvh;
    display: flex;
    align-items: center !important;
    justify-content: center !important;
    padding: clamp(16px, 2vw, 32px);
}

.cadastro-card {
    width: min(92vw, 660px);
    max-height: calc(100dvh - 32px);
    overflow-y: auto;
}

/* Remove o comportamento que jogava o card para cima */
@media (max-height: 760px) {
    .cadastro-wrapper {
        align-items: center !important;
        justify-content: center !important;
    }

    .cadastro-card {
        max-height: calc(100dvh - 28px);
    }
}
/* =====================================================
   LINKS CLEAN - CADASTRO
   Layout:
   Fazer login                                      Ajuda?
===================================================== */

.cadastro-links-clean {
    width: 100%;
    margin-top: 18px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.cadastro-links-clean a,
.cadastro-links-clean a:link,
.cadastro-links-clean a:visited,
.cadastro-links-clean a:active {
    color: #ffffff !important;
    font-size: 14px;
    font-weight: 400 !important;
    text-decoration: none !important;
    text-shadow: 0 1px 4px rgba(0, 0, 0, 0.38);
    opacity: 0.95;
}

.cadastro-links-clean a:hover {
    color: #ffffff !important;
    text-decoration: underline !important;
    opacity: 1;
}

/* Ajuste para telas muito pequenas */
@media (max-width: 520px) {
    .cadastro-links-clean {
        flex-direction: column;
        justify-content: center;
        gap: 8px;
        text-align: center;
    }
}
/* =====================================================
   CORREÇÃO GLOBAL DE TIPOGRAFIA (CADASTRO)
===================================================== */

.cadastro-card,
.cadastro-card * {
    font-size: clamp(13px, 1vw, 15px) !important;
}

.cadastro-title {
    font-size: clamp(24px, 2.5vw, 32px) !important;
}

.form-control {
    font-size: clamp(14px, 1.1vw, 16px) !important;
}

.form-label {
    font-size: clamp(13px, 1vw, 15px) !important;
}

.btn-cadastro {
    font-size: clamp(14px, 1vw, 16px) !important;
}

.cadastro-links-clean a,
.cadastro-links a {
    font-size: clamp(12px, 0.9vw, 14px) !important;
    font-weight: 400 !important;
}
/* =====================================================
   CADASTRO - RODAPÉ FINAL CORRIGIDO
   Layout:
   Fazer login                                      Ajuda?
   Ambiente interno Vale · Cadastro destinado...
===================================================== */

.cadastro-footer-final {
    width: 100%;
    margin-top: 18px;
}

.cadastro-links-row-final {
    width: 100%;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.cadastro-links-row-final a,
.cadastro-links-row-final a:link,
.cadastro-links-row-final a:visited,
.cadastro-links-row-final a:active {
    color: #ffffff !important;
    font-size: clamp(12px, 0.9vw, 14px) !important;
    font-weight: 400 !important;
    text-decoration: none !important;
    text-shadow: 0 1px 4px rgba(0, 0, 0, 0.38);
    opacity: 0.95;
}

.cadastro-links-row-final a:hover {
    color: #ffffff !important;
    text-decoration: underline !important;
    opacity: 1;
}

.cadastro-info-final {
    width: 100%;
    margin-top: 12px;
    text-align: center;
    color: #ffffff;
    font-size: clamp(11px, 0.8vw, 13px) !important;
    font-weight: 400 !important;
    line-height: 1.35;
    text-shadow: 0 1px 4px rgba(0, 0, 0, 0.38);
    opacity: 0.88;
}

/* Evita interferência de classes antigas */
.cadastro-links,
.cadastro-links-clean,
.cadastro-links-final,
.cadastro-link-line,
.cadastro-link-bottom {
    all: unset;
}

/* Reativa o comportamento correto apenas do novo bloco */
.cadastro-footer-final,
.cadastro-footer-final * {
    box-sizing: border-box;
}

/* Telas pequenas */
@media (max-width: 520px) {
    .cadastro-links-row-final {
        flex-direction: column;
        gap: 8px;
        justify-content: center;
        text-align: center;
    }

    .cadastro-info-final {
        margin-top: 10px;
    }
}
    </style>

    <script>
        function validarCadastro() {
            var senha = document.getElementById("inputSenha").value;
            var confirmarSenha = document.getElementById("inputConfirmarSenha").value;
            var matricula = document.getElementById("inputMatricula").value;

            if (matricula.length < 4) {
                alert("Informe uma matrícula válida.");
                return false;
            }

            if (senha !== confirmarSenha) {
                alert("A senha e a confirmação de senha não conferem.");
                return false;
            }

            return true;
        }
    </script>
</head>

<body>
    <div class="cadastro-wrapper">
        <main class="cadastro-card">
            <h1 class="cadastro-title">Cadastro de usuário</h1>

            <% If VAR_ALERTA <> "" Then %>
                <div class="alerta-cadastro">
                    <%=VAR_ALERTA%>
                </div>
            <% End If %>

            <form method="post" action="valida_cadastro.asp" id="formCadastro" onsubmit="return validarCadastro();">

                <div class="form-group">
                    <label class="form-label" for="inputNome">Nome:</label>
                    <input
                        type="text"
                        id="inputNome"
                        name="inputNome"
                        class="form-control"
                        placeholder="Digite seu nome completo"
                        required>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label class="form-label" for="inputMatricula">Matrícula:</label>
                        <input
                            type="text"
                            id="inputMatricula"
                            name="inputMatricula"
                            class="form-control"
                            placeholder="Matrícula"
                            required>
                    </div>

                    <div class="form-group">
                        <label class="form-label" for="inputFuncao">Função:</label>
                        <select
                            id="inputFuncao"
                            name="inputFuncao"
                            class="form-control"
                            required>
                            <option value="">Selecione...</option>

                            <%
                                SQL = "SELECT TBL_USUARIO.FUNCAO FROM TBL_USUARIO GROUP BY TBL_USUARIO.FUNCAO ORDER BY TBL_USUARIO.FUNCAO;"
                                set R_SQL = conexao_.execute(SQL)

                                Do Until R_SQL.EOF
                            %>
                                <option value="<%=R_SQL("FUNCAO")%>"><%=R_SQL("FUNCAO")%></option>
                            <%
                                    R_SQL.movenext
                                Loop
                            %>
                        </select>
                    </div>
                </div>

                <div class="form-group">
                    <label class="form-label" for="inputLocalidade">Localidade:</label>
                    <select
                        id="inputLocalidade"
                        name="inputLocalidade"
                        class="form-control"
                        required>
                        <option value="">Selecione...</option>
                        <option value="Tubarão">Tubarão</option>
                        <option value="Intendente Câmara">Intendente Câmara</option>
                        <option value="Piraqueacu">Piraqueacu</option>
                        <option value="Gov. Valadares">Gov. Valadares</option>
                        <option value="Frederico Sellow">Frederico Sellow</option>
                        <option value="Mario Carvalho">Mario Carvalho</option>
                        <option value="VIC">VIC</option>
                        <option value="VTU">VTU</option>
                        <option value="VGV">VGV</option>
                    </select>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label class="form-label" for="inputSenha">Senha:</label>
                        <input
                            type="password"
                            id="inputSenha"
                            name="inputSenha"
                            class="form-control"
                            placeholder="Digite uma senha"
                            required>
                    </div>

                    <div class="form-group">
                        <label class="form-label" for="inputConfirmarSenha">Confirmar senha:</label>
                        <input
                            type="password"
                            id="inputConfirmarSenha"
                            name="inputConfirmarSenha"
                            class="form-control"
                            placeholder="Confirme a senha"
                            required>
                    </div>
                </div>

                <button type="submit" class="btn-cadastro">Cadastrar</button>

                <div class="cadastro-footer-final">

    <div class="cadastro-links-row-final">
        <a href="form_login.asp">Fazer login</a>

        <a href="#" onclick="alert('Procure a supervisão ou responsável pela Central de Informações para apoio no cadastro.'); return false;">Ajuda?</a>
    </div>

    <div class="cadastro-info-final">
        Ambiente interno Vale · Cadastro destinado a empregados autorizados
    </div>

</div>
        </main>
    </div>
</body>
</html>
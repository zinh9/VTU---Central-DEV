<%
VAR_ALERTA = Session("alerta")
Session("alerta") = ""
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Login</title>
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

        .login-wrapper {
            width: 100%;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 32px;
        }

        .login-card {
            width: 100%;
            max-width: 620px;
            padding: 44px 56px;
            border-radius: 28px;
            background: rgba(255, 255, 255, 0.52);
            box-shadow: 0 22px 55px rgba(0, 0, 0, 0.28);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
        }

        .login-title {
            margin: 0 0 36px 0;
            text-align: center;
            color: #ffffff;
            font-size: 30px;
            font-weight: 700;
            letter-spacing: 1px;
            text-transform: uppercase;
            text-shadow: 0 2px 7px rgba(0, 0, 0, 0.42);
        }

        .form-group {
            margin-bottom: 24px;
        }

        .form-label {
            display: block;
            margin-bottom: 8px;
            color: #ffffff;
            font-size: 18px;
            font-weight: 700;
            text-shadow: 0 1px 5px rgba(0, 0, 0, 0.38);
        }

        .form-control {
            width: 100%;
            height: 50px;
            padding: 0 16px;
            border: none;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.86);
            color: #333333;
            font-size: 17px;
            outline: none;
        }

        .form-control::placeholder {
            color: #777777;
        }

        .form-control:focus {
            background: rgba(255, 255, 255, 0.96);
            box-shadow: 0 0 0 3px rgba(246, 184, 0, 0.75);
        }

        .btn-login {
            display: block;
            width: 190px;
            height: 50px;
            margin: 34px auto 26px auto;
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

        .btn-login:hover {
            background: #006B63;
            transform: translateY(-1px);
        }

        .btn-login:focus {
            outline: 3px solid rgba(246, 184, 0, 0.85);
            outline-offset: 3px;
        }

        .login-links {
            margin-top: 6px;
            text-align: left;
        }

        .login-links a {
            display: block;
            margin-bottom: 7px;
            color: #ffffff;
            font-size: 16px;
            text-decoration: none;
            text-shadow: 0 1px 5px rgba(0, 0, 0, 0.45);
        }

        .login-links a:hover {
            text-decoration: underline;
        }

        .alerta-login {
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

        .login-footer {
            margin-top: 26px;
            text-align: center;
            color: rgba(255, 255, 255, 0.92);
            font-size: 13px;
            text-shadow: 0 1px 5px rgba(0, 0, 0, 0.45);
        }

        @media (max-width: 768px) {
            .login-wrapper {
                padding: 20px;
            }

            .login-card {
                max-width: 100%;
                padding: 34px 28px;
                border-radius: 22px;
            }

            .login-title {
                font-size: 24px;
            }

            .btn-login {
                width: 100%;
            }
        }
        /* =====================================================
   RESPONSIVIDADE - TELA DE LOGIN
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

.login-wrapper {
    min-height: 100vh;
    min-height: 100svh;
    padding: clamp(12px, 2vw, 32px);
}

.login-card {
    width: min(92vw, 620px);
    max-height: calc(100svh - 32px);
    overflow-y: auto;
    padding:
        clamp(24px, 3.2vw, 44px)
        clamp(26px, 4vw, 56px);
    border-radius: clamp(18px, 2vw, 28px);
}

.login-title {
    font-size: clamp(22px, 2.5vw, 30px);
    margin-bottom: clamp(22px, 4vh, 36px);
}

.form-group {
    margin-bottom: clamp(14px, 2.4vh, 24px);
}

.form-label {
    font-size: clamp(13px, 1vw, 16px);
}


.form-control {
    height: clamp(42px, 5.8vh, 50px);
    font-size: clamp(14px, 1.3vw, 17px);
}

.btn-login {
    width: min(100%, 190px);
    height: clamp(42px, 5.6vh, 50px);
    margin:
        clamp(22px, 3.8vh, 34px)
        auto
        clamp(14px, 2.5vh, 26px)
        auto;
}

/* Notebooks com pouca altura vertical */
@media (max-height: 720px) {
    .login-wrapper {
        align-items: flex-start;
        padding-top: 22px;
        padding-bottom: 22px;
    }

    .login-card {
        max-height: calc(100svh - 44px);
        padding: 26px 42px;
    }

    .login-title {
        font-size: 24px;
        margin-bottom: 22px;
    }

    .form-group {
        margin-bottom: 14px;
    }

    .form-control {
        height: 42px;
    }

    .btn-login {
        height: 42px;
        margin-top: 22px;
        margin-bottom: 14px;
    }

    .login-footer {
        font-size: clamp(11px, 0.8vw, 13px);
    }
}

/* Telas menores */
@media (max-width: 768px) {
    .login-card {
        width: min(92vw, 520px);
    }

    .btn-login {
        width: 100%;
    }
}
/* =====================================================
   CORREÇÃO DE CENTRALIZAÇÃO - LOGIN
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

.login-wrapper {
    width: 100%;
    min-height: 100vh;
    min-height: 100dvh;
    display: flex;
    align-items: center !important;
    justify-content: center !important;
    padding: clamp(16px, 2vw, 32px);
}

.login-card {
    width: min(92vw, 620px);
    max-height: calc(100dvh - 32px);
    overflow-y: auto;
}

/* Remove o comportamento que jogava o card para cima */
@media (max-height: 720px) {
    .login-wrapper {
        align-items: center !important;
        justify-content: center !important;
    }

    .login-card {
        max-height: calc(100dvh - 28px);
    }
}
/* =====================================================
   LINKS PADRONIZADOS - LOGIN
   Mantém os links brancos e posiciona "Ajuda?" à direita
===================================================== */

.login-links-row {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 24px;
    margin-top: 8px;
    width: 100%;
}

.login-links-left {
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 7px;
}

.login-links-right {
    display: flex;
    justify-content: flex-end;
    align-items: flex-start;
    min-width: 80px;
}

.login-links-row a,
.login-links-row a:link,
.login-links-row a:visited,
.login-links-row a:active {
    color: #ffffff !important;
    font-size: 15px;
    font-weight: 700;
    text-decoration: none !important;
    text-shadow: 0 1px 5px rgba(0, 0, 0, 0.45);
}

.login-links-row a:hover {
    color: #ffffff !important;
    text-decoration: underline !important;
}

.login-links-right a {
    text-align: right;
}

/* Em telas pequenas, centraliza os links para não quebrar layout */
@media (max-width: 520px) {
    .login-links-row {
        flex-direction: column;
        align-items: center;
        gap: 10px;
    }

    .login-links-left,
    .login-links-right {
        align-items: center;
        justify-content: center;
    }

    .login-links-right {
        min-width: auto;
    }
}
/* =====================================================
   LINKS CLEAN - LOGIN
   Layout:
   Cadastre-se
   Esqueci minha senha                 Ajuda?
===================================================== */

.login-links-clean {
    width: 100%;
    margin-top: 8px;
}

.login-link-line {
    width: 100%;
    display: flex;
    justify-content: flex-start;
    align-items: center;
    margin-bottom: 7px;
}

.login-link-bottom {
    justify-content: space-between;
    gap: 20px;
}

.login-links-clean a,
.login-links-clean a:link,
.login-links-clean a:visited,
.login-links-clean a:active {
    color: #ffffff !important;
    font-size: 14px;
    font-weight: 400 !important;
    text-decoration: none !important;
    text-shadow: 0 1px 4px rgba(0, 0, 0, 0.38);
    opacity: 0.95;
}

.login-links-clean a:hover {
    color: #ffffff !important;
    text-decoration: underline !important;
    opacity: 1;
}

/* Remove influência do estilo antigo, caso ainda exista */
.login-links-clean a {
    font-size: clamp(12px, 0.9vw, 14px);
}


/* Ajuste para telas muito pequenas */
@media (max-width: 520px) {
    .login-link-line,
    .login-link-bottom {
        justify-content: center;
        text-align: center;
    }

    .login-link-bottom {
        flex-direction: column;
        gap: 7px;
    }
}
/* =====================================================
   CORREÇÃO GLOBAL DE TIPOGRAFIA (LOGIN)
===================================================== */

/* Força todos os textos pequenos a se adaptarem ao tamanho da tela */
.login-card,
.login-card * {
    font-size: clamp(13px, 1vw, 15px) !important;
}

/* Título continua grande e bonito */
.login-title {
    font-size: clamp(24px, 2.5vw, 32px) !important;
}

/* Campos */
.form-control {
    font-size: clamp(14px, 1.1vw, 16px) !important;
}

/* Labels */
.form-label {
    font-size: clamp(13px, 1vw, 15px) !important;
}

/* Botão */
.btn-login {
    font-size: clamp(14px, 1vw, 16px) !important;
}

/* Links */
.login-links-clean a,
.login-links a {
    font-size: clamp(12px, 0.9vw, 14px) !important;
    font-weight: 400 !important;
}

/* Rodapé */
.login-footer {
    font-size: clamp(11px, 0.8vw, 13px) !important;
}

/* Centralização firme */
.login-card {
    margin: auto !important;
}
/* =====================================================
   LOGIN - LINKS FINAL (FUNCIONA EM TODAS AS TELAS)
===================================================== */

.login-links-final {
    width: 100%;
    margin-top: 18px;
}

/* linha 1 */
.linha-top {
    text-align: left;
    margin-bottom: 6px;
}

/* linha 2 */
.linha-bottom {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 10px;
}

/* links */
.login-links-final a {
    color: #ffffff !important;
    text-decoration: none !important;
    font-size: clamp(12px, 0.9vw, 14px);
    font-weight: 400 !important;
    text-shadow: 0 1px 4px rgba(0,0,0,0.35);
}

.login-links-final a:hover {
    text-decoration: underline !important;
}

/* rodapé */
.login-footer-final {
    text-align: center;
    font-size: clamp(11px, 0.8vw, 13px);
    color: #ffffff;
    opacity: 0.85;
    margin-top: 10px;
}
``
    </style>
</head>

<body>
    <div class="login-wrapper">
        <main class="login-card">
            <h1 class="login-title">Central de Informações</h1>

            <% If VAR_ALERTA <> "" Then %>
                <div class="alerta-login">
                    <%=VAR_ALERTA%>
                </div>
            <% End If %>

            <form method="post" action="valida_login.asp">
                <div class="form-group">
                    <label class="form-label" for="inputLogin">Login:</label>
                    <input
                        type="text"
                        id="inputLogin"
                        name="inputLogin"
                        class="form-control"
                        placeholder="Matrícula"
                        autocomplete="username"
                        required>
                </div>

                <div class="form-group">
                    <label class="form-label" for="inputPassword">Senha:</label>
                    <input
                        type="password"
                        id="inputPassword"
                        name="inputPassword"
                        class="form-control"
                        placeholder="Digite sua senha"
                        autocomplete="current-password"
                        required>
                </div>

                <button type="submit" class="btn-login">Entrar</button>

                <div class="login-links-final">
    
    <div class="linha-top">
        <a href="form_cadastro.asp">Cadastre-se</a>
    </div>

    <div class="linha-bottom">
        <a href="#" onclick="alert('Para recuperação de senha, procure a supervisão.'); return false;">
            Esqueci minha senha
        </a>

        <a href="#" onclick="alert('Verifique matrícula e senha. Persistindo, procure suporte.'); return false;">
            Ajuda?
        </a>
    </div>
</div>
            </form>

            <div class="login-footer">
                Ambiente interno Vale · Acesso restrito a empregados autorizados
            </div>
        </main>
    </div>
</body>
</html>
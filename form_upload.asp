<!-- #include file="conexao.asp" -->

<%
If Session("matricula") <> "" Then

    If Session("funcao") = "INSPETOR" Or Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM" Then

        usuarioNome = "" & Session("name")
        usuarioFuncao = "" & Session("funcao")
        usuarioNivel = "" & Session("nivel")
        usuarioMatricula = "" & Session("matricula")
        codRegistro = "" & Session("COD")

        Function HtmlSafe(valor)
            HtmlSafe = Server.HTMLEncode("" & valor)
        End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Anexar Arquivo</title>
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
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .upload-card {
            width: min(100%, 900px);
            padding: clamp(24px, 3vw, 42px);
            border-radius: 28px;
            background: var(--card-bg);
            box-shadow: 0 22px 55px rgba(0, 0, 0, 0.25);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
        }

        .upload-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 18px;
            margin-bottom: 28px;
            padding-bottom: 18px;
            border-bottom: 1px solid rgba(0, 133, 122, 0.18);
        }

        .upload-title-area {
            flex: 1;
        }

        .upload-title {
            margin: 0;
            color: var(--vale-teal-dark);
            font-size: clamp(26px, 3.2vw, 42px);
            font-weight: 900;
            line-height: 1.1;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .upload-subtitle {
            margin-top: 8px;
            color: var(--text-muted);
            font-size: 14px;
            font-weight: 700;
            line-height: 1.45;
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

        .info-row {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 14px;
            margin-bottom: 24px;
        }

        .info-pill {
            min-height: 64px;
            padding: 12px 16px;
            border-radius: 16px;
            background: rgba(0, 133, 122, 0.08);
            border: 1.5px solid rgba(0, 133, 122, 0.18);
        }

        .info-label {
            display: block;
            color: var(--vale-teal-dark);
            font-size: 12px;
            font-weight: 900;
            text-transform: uppercase;
            margin-bottom: 5px;
        }

        .info-value {
            display: block;
            color: var(--text-dark);
            font-size: 14px;
            font-weight: 800;
            word-break: break-word;
        }

        .section-title {
            margin: 22px 0 14px 0;
            color: var(--vale-teal-dark);
            font-size: 18px;
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }

        .upload-box {
            position: relative;
            min-height: 150px;
            border: 2px dashed var(--border-gray);
            border-radius: 20px;
            background: var(--field-bg);
            display: flex;
            align-items: center;
            gap: 18px;
            padding: 22px;
            transition: all 0.18s ease-in-out;
            cursor: pointer;
        }

        .upload-box:hover,
        .upload-box:focus-within,
        .upload-box.has-file {
            border-color: var(--vale-yellow);
            background: #ffffff;
            box-shadow: 0 0 0 3px rgba(246, 184, 0, 0.24);
        }

        .upload-icon {
            width: 64px;
            height: 64px;
            border-radius: 18px;
            background: var(--vale-teal);
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 28px;
            flex-shrink: 0;
        }

        .upload-content {
            display: flex;
            flex-direction: column;
            gap: 6px;
            color: var(--text-dark);
        }

        .upload-content strong {
            font-size: 18px;
            color: var(--vale-teal-dark);
        }

        .upload-content span {
            font-size: 14px;
            color: var(--text-muted);
            line-height: 1.35;
        }

        .upload-warning {
            font-size: 12px !important;
            color: #8A6D00 !important;
        }

        .upload-file-name {
            margin-top: 6px;
            color: var(--vale-teal-dark) !important;
            font-weight: 900;
        }

        .upload-input {
            position: absolute;
            inset: 0;
            opacity: 0;
            cursor: pointer;
        }

        .preview-area {
            display: none;
            margin-top: 18px;
            padding: 16px;
            border-radius: 18px;
            border: 1.5px solid var(--border-gray);
            background: #ffffff;
        }

        .preview-area.active {
            display: block;
        }

        .preview-title {
            color: var(--vale-teal-dark);
            font-size: 13px;
            font-weight: 900;
            text-transform: uppercase;
            margin-bottom: 10px;
        }

        .preview-image {
            max-width: 100%;
            max-height: 280px;
            border-radius: 16px;
            border: 1px solid #CFD8DC;
            object-fit: contain;
            display: none;
        }

        .preview-doc {
            display: none;
            padding: 18px;
            border-radius: 16px;
            background: rgba(0, 133, 122, 0.08);
            color: var(--text-dark);
            font-weight: 800;
        }

        .actions {
            display: flex;
            justify-content: flex-end;
            align-items: center;
            gap: 12px;
            margin-top: 28px;
            padding-top: 22px;
            border-top: 1px solid rgba(0, 133, 122, 0.18);
        }

        .btn-primary,
        .btn-secondary {
            min-height: 48px;
            min-width: 170px;
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
            transition: all 0.2s ease-in-out;
        }

        .btn-primary {
            background: var(--vale-teal);
            color: #ffffff;
            box-shadow: 0 10px 22px rgba(0, 133, 122, 0.25);
        }

        .btn-primary:hover {
            background: var(--vale-teal-dark);
            transform: translateY(-1px);
        }

        .btn-secondary {
            background: #ECEFF1;
            color: var(--text-dark);
        }

        .btn-secondary:hover {
            background: #ffffff;
            transform: translateY(-1px);
            box-shadow: 0 8px 18px rgba(0,0,0,0.14);
        }

        .alert-message {
            display: none;
            margin-top: 14px;
            padding: 12px 16px;
            border-radius: 14px;
            background: rgba(198, 40, 40, 0.92);
            color: #ffffff;
            font-weight: 800;
            text-align: center;
        }

        .alert-message.active {
            display: block;
        }

        @media (max-width: 820px) {
            .upload-header {
                flex-direction: column;
                align-items: stretch;
            }

            .btn-back {
                width: 100%;
            }

            .info-row {
                grid-template-columns: 1fr;
            }

            .upload-box {
                flex-direction: column;
                align-items: flex-start;
            }

            .actions {
                flex-direction: column;
            }

            .btn-primary,
            .btn-secondary {
                width: 100%;
            }
        }

        @media (max-width: 520px) {
            .page-wrapper {
                padding: 14px;
            }

            .upload-card {
                padding: 20px;
                border-radius: 22px;
            }
        }
    </style>
</head>

<body>
    <div class="page-wrapper">
        <main class="upload-card">

            <div class="upload-header">
                <div class="upload-title-area">
                    <h1 class="upload-title">Anexar arquivo</h1>
                    <div class="upload-subtitle">
                        Vincule uma imagem, documento ou evidência ao registro recém-cadastrado.
                    </div>
                </div>

                <a href="form_home.asp" class="btn-back">← Voltar para Home</a>
            </div>

            <div class="info-row">
                <div class="info-pill">
                    <span class="info-label">Registro</span>
                    <span class="info-value">
                        <% If codRegistro <> "" Then %>
                            #<%=HtmlSafe(codRegistro)%>
                        <% Else %>
                            Não identificado
                        <% End If %>
                    </span>
                </div>

                <div class="info-pill">
                    <span class="info-label">Usuário</span>
                    <span class="info-value"><%=HtmlSafe(usuarioNome)%></span>
                </div>

                <div class="info-pill">
                    <span class="info-label">Perfil</span>
                    <span class="info-value"><%=HtmlSafe(usuarioFuncao)%></span>
                </div>
            </div>

            <div class="section-title">Selecionar anexo</div>

            <form name="formUpload" id="formUpload" action="valida_upload.asp" method="post" enctype="multipart/form-data" onsubmit="return validarUpload();">

                <label class="upload-box" id="uploadBox" for="inputFile">
                    <div class="upload-icon">📎</div>

                    <div class="upload-content">
                        <strong>Clique para selecionar um arquivo</strong>
                        <span>Formatos recomendados: JPG, JPEG, PNG, PDF, PPT ou PPTX.</span>
                        <span class="upload-warning">
                            Caso não deseje anexar nenhum arquivo, utilize a opção “Salvar sem anexo”.
                        </span>
                        <span id="fileName" class="upload-file-name"></span>
                    </div>

                    <input
                        id="inputFile"
                        name="inputFile"
                        type="file"
                        class="upload-input"
                        accept=".jpg,.jpeg,.png,.pdf,.ppt,.pptx"
                        onchange="mostrarArquivo(this)">
                </label>

                <div id="previewArea" class="preview-area">
                    <div class="preview-title">Pré-visualização</div>
                    <img id="previewImage" class="preview-image" src="" alt="Pré-visualização do arquivo selecionado">
                    <div id="previewDoc" class="preview-doc"></div>
                </div>

                <div id="alertUpload" class="alert-message">
                    Selecione um arquivo antes de clicar em “Anexar e salvar”.
                </div>

                <div class="actions">
                    <a href="form_home.asp" class="btn-secondary">
                        Salvar sem anexo
                    </a>

                    <button type="submit" class="btn-primary">
                        Anexar e salvar
                    </button>
                </div>

            </form>

        </main>
    </div>

    <script>
        function mostrarArquivo(input) {
            var fileName = document.getElementById("fileName");
            var uploadBox = document.getElementById("uploadBox");
            var previewArea = document.getElementById("previewArea");
            var previewImage = document.getElementById("previewImage");
            var previewDoc = document.getElementById("previewDoc");
            var alertUpload = document.getElementById("alertUpload");

            alertUpload.classList.remove("active");

            previewImage.style.display = "none";
            previewDoc.style.display = "none";
            previewImage.src = "";
            previewDoc.innerHTML = "";

            if (input.files && input.files.length > 0) {
                var arquivo = input.files[0];
                var nome = arquivo.name;
                var tipo = arquivo.type;

                fileName.innerHTML = "Arquivo selecionado: " + nome;
                uploadBox.classList.add("has-file");
                previewArea.classList.add("active");

                if (tipo.indexOf("image/") === 0) {
                    var reader = new FileReader();

                    reader.onload = function(e) {
                        previewImage.src = e.target.result;
                        previewImage.style.display = "block";
                    };

                    reader.readAsDataURL(arquivo);
                } else {
                    previewDoc.innerHTML = "Arquivo selecionado para anexar: " + nome;
                    previewDoc.style.display = "block";
                }
            } else {
                fileName.innerHTML = "";
                uploadBox.classList.remove("has-file");
                previewArea.classList.remove("active");
            }
        }

        function validarUpload() {
            var input = document.getElementById("inputFile");
            var alertUpload = document.getElementById("alertUpload");

            if (!input.files || input.files.length === 0) {
                alertUpload.classList.add("active");
                return false;
            }

            alertUpload.classList.remove("active");
            return true;
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
<!-- #include file="conexao.asp" -->
<!-- #include file="clsUpload.asp" -->

<%
Response.Buffer = True

If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

If Not (Session("funcao") = "INSPETOR" Or Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM") Then
    Response.Redirect("form_home.asp")
    Response.End
End If

usuarioNome = "" & Session("name")
usuarioFuncao = "" & Session("funcao")
usuarioMatricula = "" & Session("matricula")
ID = "" & Session("COD")

uploadOk = False
mensagemTitulo = "Anexo não realizado"
mensagemTexto = "Não foi possível anexar o arquivo ao registro."
mensagemTipo = "erro"
imagem_ = ""

Function HtmlSafe(valor)
    HtmlSafe = Server.HTMLEncode("" & valor)
End Function

Function ExtensaoPermitida(ext)
    ext = LCase(Trim("" & ext))

    If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "jfif" Or ext = "pdf" Or ext = "ppt" Or ext = "pptx" Or ext = "bin" Then
        ExtensaoPermitida = True
    Else
        ExtensaoPermitida = False
    End If
End Function

If ID <> "" Then

    On Error Resume Next

    Dim Upload
    Dim FileName
    Dim Folder
    Dim CaminhoFisico
    Dim sql

    Set Upload = New clsUpload

    FileName = Upload.Fields("inputFile").FileExt
    FileName = LCase(Trim("" & FileName))

    If Err.Number <> 0 Then
        mensagemTitulo = "Erro ao processar o upload"
        mensagemTexto = "Ocorreu uma falha ao ler o arquivo enviado. Tente novamente."
        mensagemTipo = "erro"
        Err.Clear

    ElseIf FileName = "" Then
        mensagemTitulo = "Nenhum arquivo selecionado"
        mensagemTexto = "Nenhum arquivo foi selecionado para anexar ao registro."
        mensagemTipo = "erro"

    ElseIf Not ExtensaoPermitida(FileName) Then
        mensagemTitulo = "Formato não permitido"
        mensagemTexto = "O formato do arquivo selecionado não é permitido. Use JPG, JPEG, PNG, GIF, PDF, PPT ou PPTX."
        mensagemTipo = "erro"

    Else
        Folder = Server.MapPath("libs\Upload\Imagens") & "\"
        CaminhoFisico = Folder & ID & "." & FileName

        ' Caminho salvo no banco. Mantemos padrão web com barra normal.
        imagem_ = "libs/Upload/Imagens/" & ID & "." & FileName

        ' Salva arquivo físico
        Upload("inputFile").SaveAs CaminhoFisico

        If Err.Number <> 0 Then
            mensagemTitulo = "Erro ao salvar o arquivo"
            mensagemTexto = "O arquivo não pôde ser salvo na pasta de upload. Verifique permissão da pasta ou tente novamente."
            mensagemTipo = "erro"
            Err.Clear
        Else
            ' Atualiza IMG_ID do registro
            sql = "UPDATE TBL_INFORMACAO SET IMG_ID = '" & Replace(imagem_, "'", "''") & "' WHERE COD=" & ID
            conexao_.Execute(sql)

            If Err.Number <> 0 Then
                mensagemTitulo = "Arquivo salvo, mas registro não atualizado"
                mensagemTexto = "O arquivo foi salvo, porém não foi possível atualizar o campo IMG_ID no banco."
                mensagemTipo = "erro"
                Err.Clear
            Else
                uploadOk = True
                mensagemTitulo = "Anexo cadastrado com sucesso!"
                mensagemTexto = "O arquivo foi vinculado ao registro #" & ID & " com sucesso."
                mensagemTipo = "sucesso"
            End If
        End If
    End If

    Set Upload = Nothing

    On Error GoTo 0

Else
    mensagemTitulo = "Registro não identificado"
    mensagemTexto = "Não foi possível identificar o código do registro para anexar o arquivo."
    mensagemTipo = "erro"
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Central de Informações - Resultado do Upload</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        :root {
            --vale-teal: #00857A;
            --vale-teal-dark: #006B63;
            --vale-yellow: #F6B800;
            --success: #00857A;
            --error: #C62828;
            --text-dark: #263238;
            --text-muted: #607D8B;
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

        .result-card {
            width: min(100%, 850px);
            padding: clamp(26px, 3vw, 44px);
            border-radius: 30px;
            background: var(--card-bg);
            box-shadow: 0 22px 55px rgba(0, 0, 0, 0.25);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            text-align: center;
        }

        .logo-area {
            margin-bottom: 20px;
        }

        .logo-area img {
            width: 118px;
            height: auto;
        }

        .status-icon {
            width: 76px;
            height: 76px;
            margin: 0 auto 22px auto;
            border-radius: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
            font-size: 38px;
            font-weight: 900;
            box-shadow: 0 14px 28px rgba(0,0,0,0.18);
        }

        .status-icon.sucesso {
            background: var(--success);
        }

        .status-icon.erro {
            background: var(--error);
        }

        .result-title {
            margin: 0;
            color: var(--vale-teal-dark);
            font-size: clamp(28px, 3vw, 42px);
            font-weight: 900;
            line-height: 1.12;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .result-text {
            max-width: 680px;
            margin: 16px auto 0 auto;
            color: var(--text-muted);
            font-size: 16px;
            font-weight: 700;
            line-height: 1.45;
        }

        .info-row {
            margin: 28px auto 0 auto;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 14px;
            max-width: 740px;
        }

        .info-pill {
            min-height: 64px;
            padding: 12px 16px;
            border-radius: 16px;
            background: rgba(0, 133, 122, 0.08);
            border: 1.5px solid rgba(0, 133, 122, 0.18);
            text-align: left;
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

        .actions {
            margin-top: 34px;
            display: flex;
            justify-content: center;
            gap: 12px;
            flex-wrap: wrap;
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

        @media (max-width: 720px) {
            .info-row {
                grid-template-columns: 1fr;
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

            .result-card {
                padding: 22px;
                border-radius: 24px;
            }
        }
    </style>
</head>

<body>
    <div class="page-wrapper">
        <main class="result-card">

            <div class="logo-area">
                <img src="libs/img/logo-vale.png" alt="Logo Vale" onerror="this.onerror=null;this.src='libs/img/Logotipo_Vale.png';">
            </div>

            <div class="status-icon <%=mensagemTipo%>">
                <% If mensagemTipo = "sucesso" Then %>
                    ✓
                <% Else %>
                    !
                <% End If %>
            </div>

            <h1 class="result-title"><%=HtmlSafe(mensagemTitulo)%></h1>

            <div class="result-text">
                <%=HtmlSafe(mensagemTexto)%>
            </div>

            <div class="info-row">
                <div class="info-pill">
                    <span class="info-label">Registro</span>
                    <span class="info-value">
                        <% If ID <> "" Then %>
                            #<%=HtmlSafe(ID)%>
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

            <div class="actions">
                <a href="form_home.asp" class="btn-primary">
                    Ir para Home
                </a>

                <a href="form_formulario.asp" class="btn-secondary">
                    Inserir novo aviso
                </a>

                <a href="form_visualizar_registro.asp" class="btn-secondary">
                    Visualizar registros
                </a>
            </div>

        </main>
    </div>
</body>
</html>

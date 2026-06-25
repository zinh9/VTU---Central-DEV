<!-- #include file="conexao.asp" -->

<%
If session("matricula") = "" Then
    response.redirect("form_login.asp")
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>DSS Online - Em Produção</title>
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
                linear-gradient(rgba(0, 60, 55, 0.28), rgba(0, 60, 55, 0.28)),
                url("libs/img/login_bg.jpg") center center / cover no-repeat fixed;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
        }

        .production-card {
            width: min(92vw, 720px);
            padding: 48px 56px;
            border-radius: 28px;
            background: rgba(255, 255, 255, 0.78);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            box-shadow: 0 22px 55px rgba(0, 0, 0, 0.28);
            text-align: center;
            color: #263238;
        }

        .production-card img {
            width: 120px;
            margin-bottom: 22px;
        }

        .production-card h1 {
            margin: 0;
            color: #006B63;
            font-size: clamp(28px, 4vw, 42px);
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .production-card p {
            margin: 18px 0 32px 0;
            font-size: 17px;
            color: #607D8B;
        }

        .btn-voltar {
            display: inline-block;
            padding: 13px 28px;
            border-radius: 10px;
            background: #00857A;
            color: #ffffff;
            text-decoration: none;
            font-weight: 700;
            transition: all 0.2s ease-in-out;
        }

        .btn-voltar:hover {
            background: #006B63;
            transform: translateY(-2px);
        }
    </style>
</head>

<body>
    <main class="production-card">
        <img src="libs/img/logo-vale.png" alt="Vale">
        <h1>Em produção</h1>
        <p>Esta funcionalidade está em desenvolvimento e será disponibilizada futuramente.</p>
        <a class="btn-voltar" href="form_home.asp">Voltar para a Central</a>
    </main>
</body>
</html>
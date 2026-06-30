<!-- #include file="conexao.asp" -->
<!-- #include file="escopo_usuarios.asp" -->
<%
' ============================================================
'  valida_nivel.asp
'  Recebe POST de form_gestao_usuarios.asp
'  Valida escopo e permissão antes de gravar NIVEL em TBL_USUARIO
' ============================================================

Response.Buffer = True

' ── Proteção de sessão ──────────────────────────────────────
If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

If Not ESCOPO_ACESSO Then
    Session("msg_gestao") = "Sem permissão para acessar esta funcionalidade."
    Response.Redirect("form_home.asp")
    Response.End
End If

' ── Parâmetros recebidos ────────────────────────────────────
Dim matAlvo   : matAlvo   = Trim("" & Request.Form("matricula"))
Dim novoNivel : novoNivel = Trim("" & Request.Form("novo_nivel"))
Dim pgVoltar  : pgVoltar  = Trim("" & Request.Form("pg"))
Dim qs        : qs        = Trim("" & Request.Form("filtros"))

If pgVoltar = "" Then pgVoltar = "1"

Dim urlVoltar : urlVoltar = "form_gestao_usuarios.asp?pg=" & pgVoltar & qs

' ── Validação 1: campos obrigatórios ────────────────────────
If matAlvo = "" Or novoNivel = "" Then
    Session("msg_gestao") = "Dados inválidos. Tente novamente."
    Response.Redirect(urlVoltar)
    Response.End
End If

' ── Validação 2: nível permitido para o logado delegar ──────
If Not PodeDelegarNivel(novoNivel) Then
    Session("msg_gestao") = "Você não tem permissão para atribuir o nível """ & novoNivel & """."
    Response.Redirect(urlVoltar)
    Response.End
End If

' ── Validação 3: empregado alvo está no escopo do logado ────
If Not PodeEditarUsuario(matAlvo) Then
    Session("msg_gestao") = "O empregado não pertence ao seu escopo de gestão."
    Response.Redirect(urlVoltar)
    Response.End
End If

' ── Validação 4: nível alvo existe (whitelist) ──────────────
Dim niveisPermitidos : niveisPermitidos = "|USER|ADM_DEL|ADM_GA|ADM_GG|ADM_DEV|"
If InStr(niveisPermitidos, "|" & novoNivel & "|") = 0 Then
    Session("msg_gestao") = "Nível inválido."
    Response.Redirect(urlVoltar)
    Response.End
End If

' ── Busca nível atual para log ──────────────────────────────
Dim nivelAnterior : nivelAnterior = ""
Dim nomeAlvo      : nomeAlvo      = ""

Dim sqlBuscaAlvo : sqlBuscaAlvo = _
    "SELECT NOME, NIVEL FROM TBL_USUARIO WHERE MATRICULA='" & _
    Replace(matAlvo,"'","''") & "'"
Dim rsBusca : Set rsBusca = conexao_.Execute(sqlBuscaAlvo)

If rsBusca.EOF Then
    rsBusca.Close
    Session("msg_gestao") = "Empregado não encontrado no banco."
    Response.Redirect(urlVoltar)
    Response.End
End If

nivelAnterior = "" & rsBusca("NIVEL")
nomeAlvo      = "" & rsBusca("NOME")
rsBusca.Close

' ── Nada a fazer se nível já é o mesmo ──────────────────────
If UCase(Trim(nivelAnterior)) = UCase(Trim(novoNivel)) Then
    Session("msg_gestao") = "Nível já era """ & novoNivel & """. Nenhuma alteração realizada."
    Response.Redirect(urlVoltar)
    Response.End
End If

' ── Gravação no banco ────────────────────────────────────────
On Error Resume Next

Dim sqlUpdate : sqlUpdate = _
    "UPDATE TBL_USUARIO SET NIVEL='" & Replace(novoNivel,"'","''") & "' " & _
    "WHERE MATRICULA='" & Replace(matAlvo,"'","''") & "'"

conexao_.Execute sqlUpdate

If Err.Number <> 0 Then
    Dim errMsg : errMsg = "" & Err.Description
    On Error GoTo 0
    Session("msg_gestao") = "Erro ao atualizar nível: " & errMsg
    Response.Redirect(urlVoltar)
    Response.End
End If

On Error GoTo 0

' ── Mensagem de sucesso ──────────────────────────────────────
Session("msg_gestao") = _
    "✅ Nível de " & nomeAlvo & " (" & matAlvo & ") alterado com sucesso de " & _
    nivelAnterior & " para " & novoNivel & "."

Response.Redirect(urlVoltar)
%>

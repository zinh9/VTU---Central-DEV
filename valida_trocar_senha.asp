<!-- #include file="conexao.asp" -->

<%
Response.Buffer = True

If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

Function SqlSafe(valor)
    SqlSafe = Replace(Trim("" & valor), "'", "''")
End Function

matriculaSessao = "" & Session("matricula")

senhaAtual = Trim("" & Request.Form("senhaAtual"))
novaSenha = Trim("" & Request.Form("novaSenha"))
confirmaSenha = Trim("" & Request.Form("confirmaSenha"))

If senhaAtual = "" Then
    Session("msg_perfil") = "Informe a senha atual."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If novaSenha = "" Then
    Session("msg_perfil") = "Informe a nova senha."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If confirmaSenha = "" Then
    Session("msg_perfil") = "Confirme a nova senha."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If novaSenha <> confirmaSenha Then
    Session("msg_perfil") = "A nova senha e a confirmação não conferem."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If novaSenha = senhaAtual Then
    Session("msg_perfil") = "A nova senha deve ser diferente da senha atual."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

' Verifica se a senha atual está correta
SQL_SENHA = "SELECT * FROM TBL_USUARIO " & _
            "WHERE MATRICULA='" & SqlSafe(matriculaSessao) & "' " & _
            "AND SENHA='" & SqlSafe(senhaAtual) & "'"

Set RS_SENHA = conexao_.Execute(SQL_SENHA)

If RS_SENHA.EOF Then
    Session("msg_perfil") = "Senha atual incorreta."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

' Atualiza a senha
SQL_UPDATE = "UPDATE TBL_USUARIO SET " & _
             "SENHA='" & SqlSafe(novaSenha) & "' " & _
             "WHERE MATRICULA='" & SqlSafe(matriculaSessao) & "'"

conexao_.Execute(SQL_UPDATE)

Session("msg_perfil") = "Senha alterada com sucesso."

Response.Redirect("form_perfil.asp")
Response.End
%>
``
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

inputMatricula = Trim("" & Request.Form("inputMatricula"))
inputNome = Trim("" & Request.Form("inputNome"))
inputFuncao = Trim("" & Request.Form("inputFuncao"))
inputCoord = Trim("" & Request.Form("inputCoord"))

' Segurança: impede alterar outro usuário manipulando o HTML
If inputMatricula = "" Or inputMatricula <> matriculaSessao Then
    Session("msg_perfil") = "Não foi possível validar a matrícula do usuário logado."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If inputNome = "" Then
    Session("msg_perfil") = "O campo Nome é obrigatório."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If inputFuncao = "" Then
    Session("msg_perfil") = "O campo Função é obrigatório."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

If inputCoord = "" Or Not IsNumeric(inputCoord) Then
    Session("msg_perfil") = "Selecione uma coordenação válida."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

' Verifica se a coordenação existe
SQL_COORD = "SELECT Count(*) AS TOTAL FROM TBL_COORD WHERE ID=" & CLng(inputCoord)
Set RS_COORD = conexao_.Execute(SQL_COORD)

If RS_COORD.EOF Or CLng(RS_COORD("TOTAL")) = 0 Then
    Session("msg_perfil") = "A coordenação selecionada não foi encontrada."
    Response.Redirect("form_perfil.asp")
    Response.End
End If

' Atualiza a fonte principal do perfil/hierarquia
SQL_UPDATE_USUARIO = "UPDATE USUARIO SET " & _
                     "NOME='" & SqlSafe(inputNome) & "', " & _
                     "CARGO='" & SqlSafe(inputFuncao) & "', " & _
                     "ID_TBL_COORD=" & CLng(inputCoord) & " " & _
                     "WHERE MATRICULA='" & SqlSafe(matriculaSessao) & "'"

conexao_.Execute(SQL_UPDATE_USUARIO)

' Atualiza também a fonte de login para manter Home e sessão coerentes
SQL_UPDATE_LOGIN = "UPDATE TBL_USUARIO SET " & _
                   "NOME='" & SqlSafe(inputNome) & "', " & _
                   "FUNCAO='" & SqlSafe(inputFuncao) & "' " & _
                   "WHERE MATRICULA='" & SqlSafe(matriculaSessao) & "'"

conexao_.Execute(SQL_UPDATE_LOGIN)

' Atualiza sessão usada nas demais telas
Session("name") = inputNome
Session("funcao") = inputFuncao

Session("msg_perfil") = "Perfil atualizado com sucesso."

Response.Redirect("form_perfil.asp")
Response.End
%>
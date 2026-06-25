<!-- #include file="conexao.asp" -->

<%
Response.Buffer = True

If Session("matricula") <> "" Then

    If Session("funcao") = "INSPETOR" Or Session("funcao") = "SUPERVISOR" Or Session("nivel") = "ADM" Then

        Function SqlSafe(valor)
            SqlSafe = Replace(Trim("" & valor), "'", "''")
        End Function

        Sub Enviar_Infomacao_Funcionario_Especifico(matricula_)
    Dim sqlCheck, rsCheck, sql_

    sqlCheck = "SELECT Count(*) AS TOTAL FROM TBL_CHECK_INFORMACAO " & _
               "WHERE MATRICULA='" & SqlSafe(matricula_) & "' " & _
               "AND COD=" & CLng(Session("COD"))

    Set rsCheck = conexao_.Execute(sqlCheck)

    If CLng(rsCheck("TOTAL")) = 0 Then
        sql_ = "INSERT INTO TBL_CHECK_INFORMACAO (MATRICULA, COD, [CHECK]) " & _
               "VALUES ('" & SqlSafe(matricula_) & "', " & CLng(Session("COD")) & ", 'NÃO')"

        conexao_.Execute(sql_)
    End If
End Sub


Sub Enviar_Infomacao_publico(publico_)
    Dim sql, resultado_, sqlCheck, rsCheck, sql_

    sql = "SELECT MATRICULA, FUNCAO FROM TBL_USUARIO WHERE FUNCAO = '" & SqlSafe(publico_) & "'"

    Set resultado_ = conexao_.Execute(sql)

    If Not resultado_.EOF Then
        Do Until resultado_.EOF

            sqlCheck = "SELECT Count(*) AS TOTAL FROM TBL_CHECK_INFORMACAO " & _
                       "WHERE MATRICULA='" & SqlSafe(resultado_("MATRICULA")) & "' " & _
                       "AND COD=" & CLng(Session("COD"))

            Set rsCheck = conexao_.Execute(sqlCheck)

            If CLng(rsCheck("TOTAL")) = 0 Then
                sql_ = "INSERT INTO TBL_CHECK_INFORMACAO (MATRICULA, COD, [CHECK]) " & _
                       "VALUES ('" & SqlSafe(resultado_("MATRICULA")) & "', " & CLng(Session("COD")) & ", 'NÃO')"

                conexao_.Execute(sql_)
            End If

            resultado_.MoveNext
        Loop
    End If
End Sub

        matricula_ = Request.Form("inputMatricula")
        data_ = Request.Form("inputData")
        tipo_ = Request.Form("inputTipo")
        titulo_ = Request.Form("inputTitulo")
        informacao_ = Request.Form("inputInformacao")

        matricula_ = SqlSafe(matricula_)
        data_ = SqlSafe(data_)
        tipo_ = SqlSafe(tipo_)
        titulo_ = SqlSafe(titulo_)
        informacao_ = SqlSafe(informacao_)

        ' Verifica duplicidade
        sql = "SELECT * FROM TBL_INFORMACAO " & _
              "WHERE FK_MATRICULA='" & matricula_ & "' " & _
              "AND INFORMACAO='" & informacao_ & "' " & _
              "AND DATA='" & data_ & "' " & _
              "AND TIPO='" & tipo_ & "' " & _
              "AND TITULO='" & titulo_ & "'"

        Set resultado_ = conexao_.Execute(sql)

        If resultado_.EOF Then

            ' Insere informação
            sql = "INSERT INTO TBL_INFORMACAO " & _
                  "(FK_MATRICULA, INFORMACAO, DATA, TIPO, IMG_ID, TITULO) " & _
                  "VALUES " & _
                  "('" & matricula_ & "', '" & informacao_ & "', '" & data_ & "', '" & tipo_ & "', 'NOIMG', '" & titulo_ & "')"

            conexao_.Execute(sql)

            ' Recupera o código gerado
            sql = "SELECT @@IDENTITY AS NewID FROM TBL_INFORMACAO"
            Set resultado_ = conexao_.Execute(sql)

            If Not resultado_.EOF Then

                Session("COD") = resultado_("NewID")

                ' Público geral por função
                SQL_FUNCOES = "SELECT TBL_USUARIO.FUNCAO FROM TBL_USUARIO GROUP BY TBL_USUARIO.FUNCAO;"
                Set R_SQL_FUNCOES = conexao_.Execute(SQL_FUNCOES)

                Do Until R_SQL_FUNCOES.EOF

                    funcaoAtual = "" & R_SQL_FUNCOES("FUNCAO")

                    If Request.Form(funcaoAtual) = funcaoAtual Then
                        Enviar_Infomacao_publico(funcaoAtual)
                    End If

                    R_SQL_FUNCOES.MoveNext
                Loop

                ' Funcionário específico
                If Request.Form("inputFuncionarioEspecifico") <> "" Then
                    Enviar_Infomacao_Funcionario_Especifico(Request.Form("inputFuncionarioEspecifico"))
                End If

                ' Redireciona para a tela moderna de upload
                Response.Redirect("form_upload.asp")
                Response.End

            Else

                Session("alerta") = "Não foi possível recuperar o código do registro inserido."
                Response.Redirect("form_formulario.asp")
                Response.End

            End If

        Else

            Session("alerta") = "Uma informação com os mesmos dados já foi inserida."
            Response.Redirect("form_formulario.asp")
            Response.End

        End If

    Else

        Response.Redirect("form_home.asp")
        Response.End

    End If

Else

    Response.Redirect("form_login.asp")
    Response.End

End If
%>
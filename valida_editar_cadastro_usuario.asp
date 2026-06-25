<!-- #include file="conexao.asp" -->
<%
	If session("matricula") <> "" Then
		
		If session("nivel") = "ADM" or session("funcao") = "SUPERVISOR" Then
			
			VAR_MATRICULA = request.form("inputMatricula")
			VAR_NOME = request.form("inputNome")
			VAR_SENHA = request.form("inputSenha")
			VAR_NIVEL = request.form("inputNivel")
			VAR_FUNCAO = request.form("inputFuncao")
			
			
			SQL_1 = "UPDATE TBL_USUARIO SET NOME='"&VAR_NOME&"', SENHA='"&VAR_SENHA&"', NIVEL='"&VAR_NIVEL&"', FUNCAO='"&VAR_FUNCAO&"' WHERE MATRICULA='"&VAR_MATRICULA&"'"

			set resultado_=conexao_.execute(SQL_1)

			if NOT resultado_.EOF Then	
				
				Session("alerta") = "Alteração feita com suceso!"
				response.redirect("form_editar_cadastro_usuario.asp") 
				
			else		
				
				Session("alerta") = "Não foi possivel efetuar alteração"
				response.redirect("form_editar_cadastro_usuario.asp") 
			
			end if
		
		else
			
			response.redirect("form_home.asp")
		
		end if

	else
		
		response.redirect("form_login.asp")

	end if
%>
			
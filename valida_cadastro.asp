<!-- #include file="conexao.asp" -->
<%
	nome_=request.form("inputNome")
	matricula_=request.form("inputMatricula")
	funcao_=request.form("inputFuncao")
	senha_=request.form("inputConfirmarSenha")
	
	nivel_="USER"
	localidade_ =request.form("inputLocalidade")
	sql=" select * from TBL_USUARIO where MATRICULA='"&matricula_&"'"

	set resultado_=conexao_.execute(sql)

	if NOT resultado_.EOF Then	
	    
		Session("alerta")="JA EXISTE UM CADASTRO COM ESSA MATRICULA" 
		response.redirect("form_cadastro.asp") 
		
	else		
		
		sql="INSERT INTO TBL_USUARIO (MATRICULA, NOME, SENHA, NIVEL, FUNCAO, LOCALIDADE) values ('"&matricula_&"', '"&nome_&"', '"&Replace(senha_,"/","")&"', '"&nivel_&"', '"&funcao_&"', '"&localidade_&"')"
		set resultado_=conexao_.execute(sql)
	    
		if NOT resultado_.EOF Then	
	     
			Session("alerta")= "Cadastro efetuado com sucesso!"
		    response.redirect("form_cadastro.asp") 
		
		else
	
	        Session("alerta")= "Cadastro efetuado com sucesso!"
		    response.redirect("form_cadastro.asp")
		
		end if	
	
	end if

%>
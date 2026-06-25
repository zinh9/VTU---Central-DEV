<!-- #include file="conexao.asp" -->
<%
	
	matricula=request.form("inputLogin")
	
	senha=request.form("inputPassword")

	sql="SELECT " & _
	"TBL_USUARIO.MATRICULA, " & _
	"TBL_USUARIO.NIVEL, " & _
	"TBL_USUARIO.NOME, " & _
	"TBL_USUARIO.FUNCAO " & _
	"FROM TBL_USUARIO " & _
	"WHERE (((TBL_USUARIO.SENHA)='" & Replace(senha,"/","") & "') " & _
	"AND ((TBL_USUARIO.MATRICULA)='" & matricula & "'));"
	
	set resultado_=conexao_.execute(sql)

	if NOT resultado_.EOF Then	
	    session("matricula")= resultado_.fields("MATRICULA").value
		session("nivel")= resultado_.fields("NIVEL").value
		Session("name") = resultado_.fields("NOME").value
		Session("funcao") = resultado_.fields("FUNCAO").value
		
		response.redirect("form_home.asp")
	else		
		Session("alerta")="Login ou Senha incorretos" 
		response.redirect("form_login.asp")
	end if

%>
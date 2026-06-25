<!-- #include file="conexao.asp" -->
<%
	
	
	sql_1="DELETE FROM TBL_INFORMACAO WHERE COD=" & request.form("VAR_ID")
	sql_2="DELETE FROM TBL_CHECK_INFORMACAO WHERE COD=" & request.form("VAR_ID")

	
	
	set resultado_1=conexao_.execute(sql_1)
	set resultado_2=conexao_.execute(sql_2)
	
	if NOT resultado_1.EOF and resultado_2.EOF   Then	

		response.write "excluido"
		
	else		
		
		response.write "erro"
	
	end if

%>

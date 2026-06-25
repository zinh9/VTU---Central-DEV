
    <%
        strCN = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
			"DBQ=\\BRAZU1VALEAS076\wwwroot\CCO\Passageiro\Sites\VTU - Central DEV\banco_teste.mdb;charset=utf8"
	Set conexao_ = Server.CreateObject("ADODB.Connection")
	
	On Error Resume Next
	
	conexao_.Open strCN 
    
	If Err.Number <> 0 Then

	Response.Write "Erro no banco de dados"
	'Executa outras coisas
	'else
	
	'Response.Write "Bem vindo"
	
	End If

	function ValidateUser(Usr, Page)
		Dim R ,i, Caminho
		
		set r= Server.CreateObject("ADODB.Recordset")
		
		For i =len(page) to 1 step -1
		    if mid(page,i,1)="/" then
			 
				caminho = mid(page,i+1,len(page)-i)
				exit for
		    end if
		next
		
		r.open "SELECT TU_PA.TipoUsuario, Pagina.DescPagina  FROM Pagina INNER JOIN TU_PA ON Pagina.IdPagina = TU_PA.IdPagina WHERE (((TU_PA.TipoUsuario)='" & usr & "') AND ((Pagina.DescPagina)='" & caminho & "'));", conexao
		
		if r.eof=true then 
			ValidateUser=0
		else
			ValidateUser=1
		end if	
	end function
    %>

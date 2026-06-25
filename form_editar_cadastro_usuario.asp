<%
If session("matricula") <> "" Then
	If session("nivel") = "ADM" or session("funcao") = "SUPERVISOR" Then
%>
<HTML xmlns="http://www.w3.org/1999/xhtml">
    
	<head>

	    <meta charset="utf-8">
        
	    <title>Central de Informações</title>
        
		<meta name="viewport" content="width=device-width, initial-scale=1">

	    
		
		<script src="libs/js/jquery.js"></script>
		<script src="libs/js/bootstrap.min.js"></script>
		
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
		  
	    <link rel="stylesheet" href="libs/css/bootstrap.min.css">
        <link href="libs/style.css" type="text/css" rel="stylesheet">
		
		<script src="TESTE.js"></script>
		
        <script src="libs/Bootstrap-datepicker/js/bootstrap-datepicker.min.js" charset="utf-8"></script> 	
        <script src="libs/Bootstrap-datepicker/js/bootstrap-datepicker.pt-BR.min.js" charset="UTF-8"></script>		
        <link href="libs/Bootstrap-datepicker/css/bootstrap-datepicker.css" rel="stylesheet"/>
			
		

	</head>


    <body>
	    
		<!-- #include file="menu.html" -->
		<!-- #include file="conexao.asp" -->
			
        <div class="container containerPrincipal">
			
			<div class="row cabecalhoConteiner">
				<div class="col-md-12">
				
					<label>Edição de Cadastro</label>
				   
				</div>
			</div>
			<div class="row">	
				<div class="col-md-2 col-md-offset-0 form-group" style="margin-top: 15;">
				
					<a href="form_home.asp" class="btn btn-lg btn-primary btn-block"role="button"><span class="glyphicon glyphicon-circle-arrow-left"></span> Voltar</a>
				
				</div>

			</div>
			
				
			<div class="col-md-6 col-md-offset-3 form-group"  >
				<input id="inputFiltroMatricula" name="inputFiltroMatricula" class="form-control" placeholder="Pesquise..." type="text">
			</div>
			
			
			<table  class="table table-bordered">
				<thead>
					<tr>
						<th>Matricula</th>
						<th>Nome</th>
						<th>Função</th>
					</tr>
				</thead>
				<tbody id="myTable">
					
					<%
					SQL_1 = "SELECT MATRICULA, NOME, SENHA, NIVEL, FUNCAO, LOCALIDADE FROM TBL_USUARIO"

					set RESULTADO_SQL_1 = conexao_.execute(SQL_1)

					if NOT RESULTADO_SQL_1.EOF Then	
						
						Do Until RESULTADO_SQL_1.EOF
							
							VAR_SQL_1_MATRICULA = RESULTADO_SQL_1("MATRICULA")
							VAR_SQL_1_NOME = RESULTADO_SQL_1("NOME")
							VAR_SQL_1_SENHA = RESULTADO_SQL_1("SENHA")
							VAR_SQL_1_NIVEL = RESULTADO_SQL_1("NIVEL")
							VAR_SQL_1_FUNCAO = RESULTADO_SQL_1("FUNCAO")
							VAR_SQL_1_LOCALIDADE = RESULTADO_SQL_1("LOCALIDADE")							
							%>
							<tr>
							
								<th><%=VAR_SQL_1_MATRICULA%></th>
								<td><%=VAR_SQL_1_NOME%></td>
								<td><%=VAR_SQL_1_FUNCAO%></td>
								
								<td><button id="botao_cod-<%=VAR_SQL_1_MATRICULA%>" value="<%=VAR_SQL_1_MATRICULA%>||<%=VAR_SQL_1_NOME%>||<%=VAR_SQL_1_SENHA%>||<%=VAR_SQL_1_NIVEL%>||<%=VAR_SQL_1_FUNCAO%>||<%=VAR_SQL_1_LOCALIDADE%>" type="button" class="btn btn-primary"><span class="glyphicon glyphicon-pencil"></span> Editar</button></td>
								<script language="javascript">
								jQuery(document).on('click', '#botao_cod-<%=VAR_SQL_1_MATRICULA%>', function (e) {
									str = document.getElementById('botao_cod-<%=VAR_SQL_1_MATRICULA%>').value;
									list = str.split('||');
									
									document.getElementById('inputMatricula').value = list[0];
									document.getElementById('inputNome').value = list[1];
									document.getElementById('inputSenha').value = list[2];
									document.getElementById('inputNivel').value = list[3];
									document.getElementById('inputFuncao').value = list[4];
									document.getElementById('inputLocalidade').value = list[5];
									
									$('#ModalTESTE').modal('show');
								});					
								</script>
							
							</tr>
							<%
							RESULTADO_SQL_1.movenext
						Loop
					end if
					%>
				</tbody>
			</table>				
	    </div>
		<script>
		$(document).ready(function(){
		  $("#inputFiltroMatricula").on("keyup", function() {
			var value = $(this).val().toLowerCase();
			$("#myTable tr").filter(function() {
			  $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1)
			});
		  });
		});
		</script>
		
		<!-- Modal -->
		<div class="modal fade" id="ModalTESTE" tabindex="-1" role="dialog" aria-labelledby="exampleModalLongTitle" aria-hidden="true">
			<div class="modal-dialog" role="document">
				<div class="modal-content">
					
					<div class="modal-header">
						
						<h5 class="modal-title" id="exampleModalLongTitle">Dados</h5>
						
						<button type="button" class="close" data-dismiss="modal" aria-label="Close">
							<span aria-hidden="true">&times;</span>
						</button>
					
					</div>
					<form name="form" id="form" action="valida_editar_cadastro_usuario.asp" method="POST">
					<div class="modal-body janela-modal">
						
						
							
							<div class="form-group">
								<label for="inputMatricula" class="control-label" >Matricula:</label>
								<input id="inputMatricula" name="inputMatricula" class="form-control" type="text" required readonly>
							</div>
							
							<div class="form-group">								
								<label for="inputNome" class="control-label" >Nome:</label>
								<input id="inputNome" name="inputNome" class="form-control" type="text" required>						
							</div>
							
							<div class="form-group">							
								<label for="inputSenha" class="control-label" >Senha:</label>
								<input id="inputSenha" name="inputSenha" class="form-control" type="text" required>					
							</div>
							
							<div class="form-group">				
								<label for="inputNivel" class="control-label">Nivel:</label>
								
								<select class="form-control" id="inputNivel" name="inputNivel"required>
								 
									<option value="USER">USER</option>
									<option value="ADM">ADM</option>

								</select>		 		
							</div>
							
							<div class="form-group">
							
								<label for="inputFuncao" class="control-label">Função:</label>
								
								<select class="form-control" id="inputFuncao" value="<%=Session("funcao")%>" name="inputFuncao"required>

									<%
								SQL = "SELECT TBL_USUARIO.FUNCAO FROM TBL_USUARIO GROUP BY TBL_USUARIO.FUNCAO;"

								set R_SQL = conexao_.execute(SQL)

								Do Until R_SQL.EOF 
									%>
										
										<option value="<%=R_SQL("FUNCAO")%>"><%=R_SQL("FUNCAO")%></option>
									<%
									R_SQL.movenext
								Loop
								%>
								  
								</select>		 
							
							</div>

							<div class="form-group">
							
								<label for="inputLocalidade" class="control-label">Localidade:</label>
								
								<select class="form-control" id="inputLocalidade" name="inputLocalidade"required>
								 
									<option value="VTU">Tubarão</option>
									<option value="VIC">Intendente Camara</option>
									<option value="VPA">Piraqueacu</option>
									<option value="VGV">Gov. Valadares</option>
									<option value="VFS">Frederico Sellow</option>
									<option value="VMR">Mario Carvalho</option>
								  
								</select>		 
							
							</div>
			   
					</div>
					
					<div class="modal-footer">
						
						<button type="submit" class="btn btn-primary"><span class="glyphicon glyphicon-floppy-disk"></span> Salvar</button>
						<button type="button" class="btn btn-secondary" data-dismiss="modal">Fechar</button>
						
					</div>
					</form>
				</div>
			</div>
		</div>
		<!-- Modal -->
			<div class="modal fade" id="Modal-Alerta" tabindex="-1" role="dialog" aria-labelledby="exampleModalLongTitle" aria-hidden="true">
				<div class="modal-dialog" role="document">
					<div class="modal-content">
						
						<div class="modal-header">
							
							<h5 class="modal-title" id="exampleModalLongTitle">Alerta!</h5>
							
							<button type="button" class="close" data-dismiss="modal" aria-label="Close">
								<span aria-hidden="true">&times;</span>
							</button>
						
						</div>
						
						<div class="modal-body janela-modal">
						
							<%VAR_ALERTA = Session("alerta")%>
							<%=VAR_ALERTA%>
							<%Session("alerta") = ""%>

						</div>
						
						<div class="modal-footer">
							<button type="button" class="btn btn-secondary" data-dismiss="modal">Fechar</button>
							
						</div>
					</div>
				</div>
			</div>
			<script language="javascript">		
			if ("<%=VAR_ALERTA%>" == "")
			{
			}
			else
			{
				$('#Modal-Alerta').modal('show');
			}		
			</script>
		
	</body>
</html>
<%
	else
		response.redirect("form_home.asp")
	end if
else
	response.redirect("form_login.asp")
end if
%>
			
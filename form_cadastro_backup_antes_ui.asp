<!doctype html>
<html lang="pt-br">  

    <head>
    
	    <meta charset="utf-8"> 
        
	    <title>Central de Informações</title>
        
		<meta name="viewport" content="width=device-width, initial-scale=1">  
		
		<script src="libs/js/jquery.js"></script>
		<script src="libs/js/bootstrap.min.js"></script>
		
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
		  
	    <link rel="stylesheet" href="libs/css/bootstrap.min.css">
        
		<link href="libs/style.css" type="text/css" rel="stylesheet">
 
        <script src="libs/Bootstrap-datepicker/js/bootstrap-datepicker.min.js" charset="utf-8"></script> 	
        <script src="libs/Bootstrap-datepicker/js/bootstrap-datepicker.pt-BR.min.js" charset="UTF-8"></script>		
        <link href="libs/Bootstrap-datepicker/css/bootstrap-datepicker.css" rel="stylesheet"/>
	
	</head>

    <body>  
		
		<nav class="navbar navbar-default" role="navigation">	
			
			<div class="container">
				
				<a class="navbar-brand"><img src="libs/img/logo-vale.png" style="margin-top: -10px"></a>
				
				<a class="navbar-brand">Central de Informações</a>

			</div>
		  
		</nav>
	
        <div class="container containerLogin">
		    
			<form name="form" id="form" action="valida_cadastro.asp" method="POST">
			
			    <div class="row">
			        
					<div class="col-md-12 cabecalho">
			            
						<h3 style="white-space:nowrap" class="form-signin-heading">Cadastro de usuário</h3>
			       
				   </div>
			    
				</div>
                
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
							
							<label style="color: #ffffff" for="inputNome" class="control-label">Nome:</label>
                            <input id="inputNome" name="inputNome" class="form-control" placeholder="Digite seu nome..." type="text" required>
						
						</div>
			       
				    </div>
					
				
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
				
							<label style="color: #ffffff" for="inputMatricula" class="control-label">Matricula:</label>
							<input id="inputMatricula" name="inputMatricula" class="form-control" placeholder="Matricula" type="text"required>
						
						</div>
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
			    		
						    <label style="color: #ffffff" for="inputFuncao" class="control-label">Função:</label>
                            
			    	        <select class="form-control" id="inputFuncao" name="inputFuncao"required>
                             
			    		        <option value=""></option>
								<!-- #include file="conexao.asp" -->
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
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
			    		
						    <label style="color: #ffffff" for="inputLocalidade" class="control-label">Localidade:</label>
                            
			    	        <select class="form-control" id="inputLocalidade" name="inputLocalidade"required>
                             
			    		        <option value=""></option>
								<option value="VTU">Tubarão</option>
			    		        <option value="VIC">Intendente Câmara</option>
								                              
			    		    </select>		 
			    	    
						</div>
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
							
							<label style="color: #ffffff" for="inputSenha" class="control-label">Senha: (data de nascimento)</label>
							<input id="inputSenha" name="inputSenha" class="form-control" placeholder="Digite uma senha..." type="password"required>
							
						</div>
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
							
							<label style="color: #ffffff" for="inputConfirmarSenha" class="control-label">Confirmar senha:</label>
							<input id="inputConfirmarSenha" name="inputConfirmarSenha" class="form-control" oninput="validaSenha(this)" placeholder="Confirme a Senha..." type="password"required>
							
							<script>
							
							function validaSenha (input)
							{ 
								
								if (input.value != document.getElementById('inputSenha').value) {
									
									input.setCustomValidity('Repita a senha corretamente');
								
								}
								else
								{
									
									input.setCustomValidity('');
								
								}
                            
							} 
	
							</script>
						
						</div>
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
							
							<button class="btn btn-lg btn-primary btn-block botaoenviar"  type="submit">Cadastrar</button>
							
						</div>
			       
				    </div>
					
				</div>
					
				<div class="row">
				    
					<div class="col-md-12">
				
						<div class="form-group">	
							
							<a class="link-branco" href="form_login.asp" style="white-space:nowrap" >Fazer login</a>
							
							<p><a class="link-branco" href="#" style="white-space:nowrap"onclick="Ajuda()" >Ajuda?</a></p>
							
							<script language="javascript">
									
								function Ajuda(){
									
									alert ("Para suporte ou sugestões, envie um e mail para mariely.prottes@vale.com e gilmar.oliveira@vale.com.\nAssim que possível responderemos.");
								
								}
								
							</script>							
						
						</div>
					
					</div>
				
				</div>
	    
			</form>
       
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
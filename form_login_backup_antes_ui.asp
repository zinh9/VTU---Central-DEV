<HTML xmlns="http://www.w3.org/1999/xhtml">
    
	<head>
    
	    <meta charset="utf-8"> 
        
	    <title>Central de Informações</title>
        
		<meta name="viewport" content="width=device-width, initial-scale=1">

	    
		
		<script src="libs/js/jquery.js"></script>
		<script src="libs/js/bootstrap.min.js"></script>
		
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
		<!--<link rel="stylesheet" href="libs/css/v3.3.7/bootstrap.min.css">-->
		  
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
		    
			<form name="form" id="form" action="valida_login.asp" method="POST">
			
			    <div class="row">
			        
					<div class="col-md-12 cabecalho">
			            
						<h3 style="white-space:nowrap" class="form-signin-heading">Login</h3>
			       
				   </div>
			    
				</div>
                
			    
				
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
							
							<label style="color: #ffffff" for="inputLogin" class="control-label" >Matricula:</label>
							<input id="inputLogin" name="inputLogin" class="form-control" placeholder="Matricula (8 dígitos)" type="text" required>
						
						</div>
			       
				    </div>
					
				
				</div>
				
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
				
							<label style="color: #ffffff" for="inputPassword" class="control-label" >Senha: (data de nascimento)</label>
							<input id="inputPassword" name="inputPassword" class="form-control" placeholder="senha" type="password" required>
						
						</div>
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">
							
							<button class="btn btn-lg btn-primary btn-block botaoenviar"  type="submit">Entrar</button>
							
						</div>
			       
				    </div>
					
				</div>
				
				<div class="row">
				    
					<div class="col-md-12">
			            
						<div class="form-group">	
							
							<a class="link-branco" href="form_cadastro.asp">Cadastre-se</a>
							<p><a class="link-branco" href="#" style="white-space:nowrap" onclick="Ajuda()" >Precisa de ajuda?</a></p>		
						
                        </div>
						
						<script language="javascript">
									
								function Ajuda(){
									
									alert ("Para suporte ou sugestões, envie um e mail para mariely.prottes@vale.com ou gilmar.oliveira@vale.com.\nAssim que possível responderemos.");
								
								}
								
							</script>
			       
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
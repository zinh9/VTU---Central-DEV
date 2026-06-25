
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
			<table  class="table table-bordered">
				<thead>
					<tr>
						<th>COD</th>
						<th>DATA</th>
						<th>AUTOR</th>
						<th>TITULO</th>
						<th>GRAFICO</th>
					</tr>
				</thead>
				<tbody>
					
					<%
					SQL_1 = "SELECT DISTINCT TBL_INFORMACAO.COD, TBL_INFORMACAO.DATA, TBL_USUARIO.NOME, TBL_INFORMACAO.TITULO FROM TBL_INFORMACAO INNER JOIN TBL_USUARIO ON TBL_INFORMACAO.FK_MATRICULA = TBL_USUARIO.MATRICULA"

					set RESULTADO_SQL_1 = conexao_.execute(SQL_1)

					if NOT RESULTADO_SQL_1.EOF Then	
						
						Do Until RESULTADO_SQL_1.EOF
							
							VAR_SQL_1_COD = RESULTADO_SQL_1("COD")
							VAR_SQL_1_DATA = RESULTADO_SQL_1("DATA")
							VAR_SQL_1_TITULO = RESULTADO_SQL_1("TITULO")
							VAR_SQL_1_MATRICULA = RESULTADO_SQL_1("FK_MATRICULA")
							VAR_SQL_2_NOME = RESULTADO_SQL_1("NOME")	   
					        %>
							<tr>
							
								<th scope="row"><%=VAR_SQL_1_COD%></th>
								<td><%=VAR_SQL_1_DATA%></td>
								<td><%=VAR_SQL_2_NOME%></td>
								<td><%=VAR_SQL_1_TITULO%></td>
								<td><button data-toggle="modal" data-target="#Modal-<%=VAR_SQL_1_COD%>" type="button" class="btn btn-primary"
								<%
								MODAL_SQL_1= "SELECT DISTINCT TBL_CHECK_INFORMACAO.MATRICULA, TBL_USUARIO.NOME, Date()-[DATA] AS PENDENCIA_NAO FROM (TBL_CHECK_INFORMACAO INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD) INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA WHERE (((TBL_CHECK_INFORMACAO.CHECK)='NÃO') AND ((TBL_INFORMACAO.COD)="&VAR_SQL_1_COD&"))"
													
								set MODAL_RESULTADO_SQL_1 = conexao_.execute(MODAL_SQL_1)

								if NOT MODAL_RESULTADO_SQL_1.EOF Then	
									
									Do Until MODAL_RESULTADO_SQL_1.EOF
										
										MODAL_VAR_SQL_1_MATRICULA = MODAL_RESULTADO_SQL_1("MATRICULA")
										MODAL_VAR_SQL_1_PENDENCIA = MODAL_RESULTADO_SQL_1("PENDENCIA_NAO")
										MODAL_VAR_SQL_2_NOME = MODAL_RESULTADO_SQL_1("NOME")
										
										VAR_STRING_NAO = VAR_STRING_NAO & "" & MODAL_VAR_SQL_1_MATRICULA & "|" & MODAL_VAR_SQL_2_NOME & "|" & MODAL_VAR_SQL_1_PENDENCIA & ""
										
										MODAL_RESULTADO_SQL_1.movenext
									Loop
									
								end if
								%>
								value=""
								><span class="glyphicon glyphicon-signal"></span> Grafico</button></td>
							
							</tr>
					        <%
							RESULTADO_SQL_1.movenext
						Loop
					end if
					%>
				</tbody>
			</table>
				
				<%
				SQL_1 = "TRANSFORM Count(TBL_CHECK_INFORMACAO.CHECK) AS ContarDeCHECK SELECT TBL_CHECK_INFORMACAO.COD FROM TBL_CHECK_INFORMACAO GROUP BY TBL_CHECK_INFORMACAO.COD PIVOT TBL_CHECK_INFORMACAO.CHECK"

				set RESULTADO_SQL_1 = conexao_.execute(SQL_1)

				if NOT RESULTADO_SQL_1.EOF Then	
					
					Do Until RESULTADO_SQL_1.EOF
						
						VAR_SQL_1_COD = RESULTADO_SQL_1("COD")
						
						VAR_SQL_NAO_QUANTIDADE = CInt(RESULTADO_SQL_1("NÃO"))

						VAR_SQL_SIM_QUANTIDADE = CInt(RESULTADO_SQL_1("SIM"))
						
						%>
						<!--erro inicio-->
						<div class="modal fade" id="Modal-<%=VAR_SQL_1_COD%>" tabindex="-1" role="dialog" aria-labelledby="exampleModalLongTitle" aria-hidden="true">
							<div class="modal-dialog" role="document" style="width:50% !important;">
								<div class="modal-content">
									
									<div class="modal-header">
										
										<h5 class="modal-title" id="exampleModalLongTitle">COD - <%=VAR_SQL_1_COD%></h5>
										
										<button type="button" class="close" data-dismiss="modal" aria-label="Close">
											<span aria-hidden="true">&times;</span>
										</button>
									
									</div>
									
									<div class="modal-body janela-modal">
										
										
										
											<div id="piechart_<%=VAR_SQL_1_COD%>" style="height: 400px; width: auto;"></div>
											<script>
											
												CanvasJS.addColorSet("greenShades",
															[//colorSet Array

															"#109618",
															"#DC3912"           
															]);

											var chart = new CanvasJS.Chart("piechart_<%=VAR_SQL_1_COD%>", {
												colorSet: "greenShades",
												title: {
													text: "Porcentagem Visualização"
												},
												
												data: [{
													type: "pie",
													startAngle: 270,
													toolTipContent: "{label} - ({y})",
													indexLabel: "{label} (#percent%)",
													percentFormatString: "#0.#",
													dataPoints: [
														{y: <%=VAR_SQL_SIM_QUANTIDADE%>, label: "Visualizados"},
														{y: <%=VAR_SQL_NAO_QUANTIDADE%>, label: "Não visualizados"}
													]
												}]
											});
											chart.render();
											</script>
											<div class="cabecalhoConteinervermelho">

													<label>Publico Pendente</label>

											</div>
											<table  class="table table-bordered">
												<thead>
													<tr>
														<th>MATRICULA</th>
														<th>NOME</th>
														<th>TEMPO PENDENTE(DIA)</th>
							
													</tr>
												</thead>
												<tbody>
													<%
													MODAL_SQL_1= "SELECT DISTINCT TBL_CHECK_INFORMACAO.MATRICULA, TBL_USUARIO.NOME, Date()-[DATA] AS PENDENCIA_NAO FROM (TBL_CHECK_INFORMACAO INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD) INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA WHERE (((TBL_CHECK_INFORMACAO.CHECK)='NÃO') AND ((TBL_INFORMACAO.COD)="&VAR_SQL_1_COD&"))"
													
													set MODAL_RESULTADO_SQL_1 = conexao_.execute(MODAL_SQL_1)

													if NOT MODAL_RESULTADO_SQL_1.EOF Then	
														
														Do Until MODAL_RESULTADO_SQL_1.EOF
															
															MODAL_VAR_SQL_1_MATRICULA = MODAL_RESULTADO_SQL_1("MATRICULA")
                                                            MODAL_VAR_SQL_1_PENDENCIA = MODAL_RESULTADO_SQL_1("PENDENCIA_NAO")
															MODAL_VAR_SQL_2_NOME = MODAL_RESULTADO_SQL_1("NOME")
													        %>
													<tr>
													
														<th scope="row"><%=MODAL_VAR_SQL_1_MATRICULA%></th>
														<td><%=MODAL_VAR_SQL_2_NOME%></td>
														<td><%=MODAL_VAR_SQL_1_PENDENCIA%></td>
														
													</tr>
													<%
															MODAL_RESULTADO_SQL_1.movenext
														Loop
													end if
													%>
												</tbody>
											</table>
											<div class="cabecalhoConteiner">

													<label>Publico que Visualizou</label>

											</div>
											<table  class="table table-bordered">
												<thead>
													<tr>
														<th>MATRICULA</th>
														<th>NOME</th>
							                            <th>TEMPO CORRIDO(DIA)</th>
													</tr>
												</thead>
												<tbody>
													
													
													
													
													<%
													MODAL_SQL_1 = "SELECT DISTINCT TBL_CHECK_INFORMACAO.MATRICULA, TBL_USUARIO.NOME, [DATA_VISUALIZACAO]-[DATA] AS PENDENCIA_SIM FROM (TBL_CHECK_INFORMACAO INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA) INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD WHERE (((TBL_CHECK_INFORMACAO.CHECK)='SIM') AND ((TBL_CHECK_INFORMACAO.COD)="&VAR_SQL_1_COD&"))"
														
													set MODAL_RESULTADO_SQL_1 = conexao_.execute(MODAL_SQL_1)

													if NOT MODAL_RESULTADO_SQL_1.EOF Then	
														
														Do Until MODAL_RESULTADO_SQL_1.EOF
															
															
																MODAL_VAR_SQL_1_MATRICULA = MODAL_RESULTADO_SQL_1("MATRICULA")
																MODAL_VAR_SQL_1_NOME = MODAL_RESULTADO_SQL_1("NOME")
																MODAL_VAR_SQL_1_PENDENCIA = MODAL_RESULTADO_SQL_1("PENDENCIA_SIM")
                                                     
													 response.write "<tr>"
														 response.write "<th scope='row'>"
															response.write MODAL_VAR_SQL_1_MATRICULA
														 response.write "</th>"
														 response.write "<td>"
															response.write MODAL_VAR_SQL_1_NOME
														 response.write "</td>"
														 response.write "<td>"
															response.write MODAL_VAR_SQL_1_PENDENCIA
														 response.write "</td>"													 
													 response.write "</tr>"
													 
													 
													 
													 
															MODAL_RESULTADO_SQL_1.movenext
														Loop
													end if
													%>
													
													
													
													
												</tbody>
											</table>	
									</div>
																						
									<div class="modal-footer">
										<button type="button" class="btn btn-secondary" data-dismiss="modal">Fechar</button>
									</div>
								</div>	
							</div>	
						</div>	
						<!--erro fim-->
						<%
						RESULTADO_SQL_1.movenext
					Loop
				end if
				%>
				

	    </div>
	</body>
</html>


<%
VAR_SQL_1_COD = Request.Item("COD")
%>

<!-- #include file="conexao.asp" -->

<!--erro inicio-->

<div class="modal-header">
	
	<h5 class="modal-title" id="exampleModalLongTitle">COD - <%=VAR_SQL_1_COD%></h5>
	
	<button type="button" class="close" data-dismiss="modal" aria-label="Close">
		<span aria-hidden="true">&times;</span>
	</button>

</div>

<div class="modal-body janela-modal">

	<div id="piechart" style="height: 400px; width: auto;"></div>
	
	<%
	SQL_1 = "TRANSFORM Count(TBL_CHECK_INFORMACAO.CHECK) AS ContarDeCHECK SELECT TBL_CHECK_INFORMACAO.COD FROM TBL_CHECK_INFORMACAO WHERE TBL_CHECK_INFORMACAO.COD = " & VAR_SQL_1_COD & " GROUP BY TBL_CHECK_INFORMACAO.COD PIVOT TBL_CHECK_INFORMACAO.CHECK"
	set RESULTADO_SQL_1 = conexao_.execute(SQL_1)
	if NOT RESULTADO_SQL_1.EOF Then    
		VAR_SQL_NAO_QUANTIDADE = RESULTADO_SQL_1("NÃO")
		VAR_SQL_SIM_QUANTIDADE = RESULTADO_SQL_1("SIM")
	end if
	%>
	<script>
	
		CanvasJS.addColorSet("greenShades",
					[//colorSet Array

					"#109618",
					"#DC3912"           
					]);

	var chart = new CanvasJS.Chart("piechart", {
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

			MODAL_SQL_1 = "SELECT DISTINCT TBL_CHECK_INFORMACAO.MATRICULA, TBL_USUARIO.NOME, Date()-[DATA] AS PENDENCIA_NAO FROM (TBL_CHECK_INFORMACAO INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD) INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA WHERE (((TBL_CHECK_INFORMACAO.CHECK)='NÃO') AND ((TBL_INFORMACAO.COD)=" & VAR_SQL_1_COD & "))"
			
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
	<button type="button" class="btn btn-secondary" onclick="$('#ContentModalGrafico').html('');" data-dismiss="modal">Fechar</button>
</div>

<!--erro fim-->
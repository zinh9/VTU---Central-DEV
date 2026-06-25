<!-- #include file="conexao.asp" -->
<%
' ============================================================
'  api_relatorios.asp
'  API JSON para a tela de Relatórios da Central de Informações
'  Retorna dados consolidados de leitura para Chart.js
'  Requer: Session("matricula") ativo
'          Nível: INSPETOR, SUPERVISOR ou ADM
' ============================================================

Response.ContentType = "application/json"
Response.CharSet     = "utf-8"

' --- helpers -------------------------------------------------
Function JS(v)
    Dim s
    s = "" & v
    s = Replace(s, "\",  "\\")
    s = Replace(s, """",  "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr,   "\n")
    s = Replace(s, vbLf,   "\n")
    JS = s
End Function

Function NumSafe(v)
    If IsNumeric(v) And Not IsNull(v) Then
        NumSafe = CLng(v)
    Else
        NumSafe = 0
    End If
End Function

' --- autenticacao --------------------------------------------
If Session("matricula") = "" Then
    Response.Write "{""erro"":""Sessao expirada""}"
    Response.End
End If

NIVEL_  = "" & Session("nivel")
FUNCAO_ = "" & Session("funcao")

If FUNCAO_ <> "INSPETOR" And FUNCAO_ <> "SUPERVISOR" And NIVEL_ <> "ADM" Then
    Response.Write "{""erro"":""Sem permissao""}"
    Response.End
End If

' --- parametro de acao --------------------------------------
'  ?acao=resumo          -> KPIs gerais (cards do topo)
'  ?acao=leituras_dia    -> leituras confirmadas por dia (linha/barra)
'  ?acao=por_tipo        -> registros por tipo (doughnut)
'  ?acao=tempo_medio     -> tempo medio de leitura por registro (barras horizontais)
'  ?acao=pendencias_user -> top usuarios com mais pendencias (barras)
'  ?acao=historico       -> tabela historico completo paginada
'  ?acao=por_registro    -> detalhes de um registro (pie SIM/NAO + lista)
'  Filtros opcionais em qualquer acao:
'    &data_ini=DD/MM/AAAA  &data_fim=DD/MM/AAAA  &tipo=TEXTO  &cod=N

acao     = LCase(Trim("" & Request("acao")))
data_ini = Trim("" & Request("data_ini"))
data_fim = Trim("" & Request("data_fim"))
filtTipo = Trim("" & Request("tipo"))
filtCod  = Trim("" & Request("cod"))
pg       = Trim("" & Request("pg"))
If pg = "" Or Not IsNumeric(pg) Then pg = 1
pg = CInt(pg)
If pg < 1 Then pg = 1

' --- montagem de filtro de data comum -----------------------
'  O banco guarda DATA como texto DD/MM/AAAA
'  Usamos LIKE '%AAAA%' para filtro de ano, ou comparamos direto
whereData = ""
If data_ini <> "" Then
    whereData = whereData & " AND TBL_INFORMACAO.DATA >= '" & Replace(data_ini,"'","''") & "' "
End If
If data_fim <> "" Then
    whereData & " AND TBL_INFORMACAO.DATA <= '" & Replace(data_fim,"'","''") & "' "
End If
If filtTipo <> "" Then
    whereData = whereData & " AND TBL_INFORMACAO.TIPO LIKE '%" & Replace(filtTipo,"'","''") & "%' "
End If
If filtCod <> "" And IsNumeric(filtCod) Then
    whereData = whereData & " AND TBL_INFORMACAO.COD = " & CLng(filtCod) & " "
End If

' ============================================================
Select Case acao

' ------------------------------------------------------------
' RESUMO – KPIs gerais
' ------------------------------------------------------------
Case "resumo"

    ' Total de registros publicados
    Set rs1 = conexao_.Execute("SELECT Count(*) AS T FROM TBL_INFORMACAO WHERE 1=1 " & whereData)
    totalRegistros = NumSafe(rs1("T"))
    rs1.Close

    ' Total de leituras confirmadas (SIM)
    Set rs2 = conexao_.Execute( _
        "SELECT Count(*) AS T FROM TBL_CHECK_INFORMACAO " & _
        "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
        "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='SIM' " & whereData)
    totalSim = NumSafe(rs2("T"))
    rs2.Close

    ' Total de pendencias (NAO)
    Set rs3 = conexao_.Execute( _
        "SELECT Count(*) AS T FROM TBL_CHECK_INFORMACAO " & _
        "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
        "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='NÃO' " & whereData)
    totalNao = NumSafe(rs3("T"))
    rs3.Close

    ' Total de usuarios destinatarios unicos
    Set rs4 = conexao_.Execute( _
        "SELECT Count(DISTINCT TBL_CHECK_INFORMACAO.MATRICULA) AS T " & _
        "FROM TBL_CHECK_INFORMACAO " & _
        "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
        "WHERE 1=1 " & whereData)
    totalUsuarios = NumSafe(rs4("T"))
    rs4.Close

    ' Taxa global de leitura
    totalGeral = totalSim + totalNao
    If totalGeral > 0 Then
        taxaLeitura = FormatNumber((totalSim / totalGeral) * 100, 1)
    Else
        taxaLeitura = "0.0"
    End If

    Response.Write "{" & _
        """total_registros"":" & totalRegistros & "," & _
        """total_leituras_sim"":" & totalSim & "," & _
        """total_pendencias"":" & totalNao & "," & _
        """total_usuarios"":" & totalUsuarios & "," & _
        """taxa_leitura_pct"":""" & taxaLeitura & """" & _
    "}"

' ------------------------------------------------------------
' LEITURAS_DIA – quantidade de leituras confirmadas por dia
' ------------------------------------------------------------
Case "leituras_dia"

    SQL_LD = "SELECT TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO AS DIA, " & _
             "Count(*) AS TOTAL " & _
             "FROM TBL_CHECK_INFORMACAO " & _
             "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
             "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='SIM' " & _
             "AND TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO IS NOT NULL " & whereData & _
             "GROUP BY TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO " & _
             "ORDER BY TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO DESC"

    Set rsLD = conexao_.Execute(SQL_LD)

    labels = ""
    valores = ""
    primeiro = True

    Do Until rsLD.EOF
        If Not primeiro Then
            labels  = labels  & ","
            valores = valores & ","
        End If
        labels  = labels  & """" & JS(rsLD("DIA")) & """"
        valores = valores & NumSafe(rsLD("TOTAL"))
        primeiro = False
        rsLD.MoveNext
    Loop
    rsLD.Close

    Response.Write "{""labels"":[" & labels & "],""valores"":[" & valores & "]}"

' ------------------------------------------------------------
' POR_TIPO – contagem de registros por tipo (TIPO)
' ------------------------------------------------------------
Case "por_tipo"

    SQL_PT = "SELECT TBL_INFORMACAO.TIPO, Count(*) AS TOTAL " & _
             "FROM TBL_INFORMACAO " & _
             "WHERE 1=1 " & whereData & _
             "GROUP BY TBL_INFORMACAO.TIPO " & _
             "ORDER BY TOTAL DESC"

    Set rsPT = conexao_.Execute(SQL_PT)

    labels  = ""
    valores = ""
    primeiro = True

    Do Until rsPT.EOF
        If Not primeiro Then
            labels  = labels  & ","
            valores = valores & ","
        End If
        tipoVal = "" & rsPT("TIPO")
        If tipoVal = "" Then tipoVal = "Sem tipo"
        labels  = labels  & """" & JS(tipoVal) & """"
        valores = valores & NumSafe(rsPT("TOTAL"))
        primeiro = False
        rsPT.MoveNext
    Loop
    rsPT.Close

    Response.Write "{""labels"":[" & labels & "],""valores"":[" & valores & "]}"

' ------------------------------------------------------------
' TEMPO_MEDIO – dias corridos até leitura por registro
'   DATA_VISUALIZACAO - DATA (ambos armazenados como texto DD/MM/AAAA)
'   Usamos a coluna calculada já existente em aux_grafico
' ------------------------------------------------------------
Case "tempo_medio"

    ' Retorna os 10 registros com maior tempo medio de leitura
    SQL_TM = "SELECT TOP 10 " & _
             "TBL_INFORMACAO.COD, " & _
             "TBL_INFORMACAO.TITULO, " & _
             "Avg([TBL_CHECK_INFORMACAO].[DATA_VISUALIZACAO]-[TBL_INFORMACAO].[DATA]) AS MEDIA_DIAS " & _
             "FROM TBL_CHECK_INFORMACAO " & _
             "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
             "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='SIM' " & _
             "AND TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO IS NOT NULL " & whereData & _
             "GROUP BY TBL_INFORMACAO.COD, TBL_INFORMACAO.TITULO " & _
             "ORDER BY MEDIA_DIAS DESC"

    Set rsTM = conexao_.Execute(SQL_TM)

    labels  = ""
    valores = ""
    primeiro = True

    Do Until rsTM.EOF
        If Not primeiro Then
            labels  = labels  & ","
            valores = valores & ","
        End If
        tituloTM = "" & rsTM("TITULO")
        If Len(tituloTM) > 35 Then tituloTM = Left(tituloTM, 35) & "..."
        labels  = labels  & """#" & NumSafe(rsTM("COD")) & " " & JS(tituloTM) & """"
        mediaDias = rsTM("MEDIA_DIAS")
        If IsNumeric(mediaDias) Then
            valores = valores & FormatNumber(CDbl(mediaDias), 1)
        Else
            valores = valores & "0"
        End If
        primeiro = False
        rsTM.MoveNext
    Loop
    rsTM.Close

    Response.Write "{""labels"":[" & labels & "],""valores"":[" & valores & "]}"

' ------------------------------------------------------------
' PENDENCIAS_USER – top 10 usuarios com mais pendencias abertas
' ------------------------------------------------------------
Case "pendencias_user"

    SQL_PU = "SELECT TOP 10 TBL_USUARIO.NOME, Count(*) AS TOTAL " & _
             "FROM TBL_CHECK_INFORMACAO " & _
             "INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA " & _
             "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
             "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='NÃO' " & whereData & _
             "GROUP BY TBL_USUARIO.NOME " & _
             "ORDER BY TOTAL DESC"

    Set rsPU = conexao_.Execute(SQL_PU)

    labels  = ""
    valores = ""
    primeiro = True

    Do Until rsPU.EOF
        If Not primeiro Then
            labels  = labels  & ","
            valores = valores & ","
        End If
        labels  = labels  & """" & JS(rsPU("NOME")) & """"
        valores = valores & NumSafe(rsPU("TOTAL"))
        primeiro = False
        rsPU.MoveNext
    Loop
    rsPU.Close

    Response.Write "{""labels"":[" & labels & "],""valores"":[" & valores & "]}"

' ------------------------------------------------------------
' HISTORICO – tabela paginada de todos os registros com status
' ------------------------------------------------------------
Case "historico"

    limite = 15
    inicio = (pg - 1) * limite

    SQL_HC_BASE = "FROM TBL_INFORMACAO " & _
                  "INNER JOIN TBL_USUARIO ON TBL_INFORMACAO.FK_MATRICULA = TBL_USUARIO.MATRICULA " & _
                  "WHERE 1=1 " & whereData

    ' Total
    Set rsHCTotal = conexao_.Execute("SELECT Count(*) AS T " & SQL_HC_BASE)
    totalH = NumSafe(rsHCTotal("T"))
    rsHCTotal.Close

    totalPgH = 1
    If totalH > 0 Then totalPgH = Int((totalH + limite - 1) / limite)
    If pg > totalPgH Then pg = totalPgH

    ' Paginacao Access (sem OFFSET)
    If inicio = 0 Then
        sqlHC = "SELECT TOP " & limite & " " & _
                "TBL_INFORMACAO.COD, TBL_INFORMACAO.DATA, " & _
                "TBL_INFORMACAO.TITULO, TBL_INFORMACAO.TIPO, " & _
                "TBL_USUARIO.NOME AS AUTOR " & _
                SQL_HC_BASE & " ORDER BY TBL_INFORMACAO.COD DESC"
    Else
        sqlHC = "SELECT TOP " & limite & " " & _
                "TBL_INFORMACAO.COD, TBL_INFORMACAO.DATA, " & _
                "TBL_INFORMACAO.TITULO, TBL_INFORMACAO.TIPO, " & _
                "TBL_USUARIO.NOME AS AUTOR " & _
                SQL_HC_BASE & _
                " AND TBL_INFORMACAO.COD NOT IN (" & _
                    "SELECT TOP " & inicio & " TBL_INFORMACAO.COD " & _
                    SQL_HC_BASE & " ORDER BY TBL_INFORMACAO.COD DESC" & _
                ") ORDER BY TBL_INFORMACAO.COD DESC"
    End If

    Set rsHC = conexao_.Execute(sqlHC)

    ' Para cada registro, busca contadores SIM/NAO
    itens = ""
    primeiroH = True

    Do Until rsHC.EOF
        codH  = NumSafe(rsHC("COD"))
        dataH = "" & rsHC("DATA")
        titH  = "" & rsHC("TITULO")
        tipH  = "" & rsHC("TIPO")
        autH  = "" & rsHC("AUTOR")

        Set rsSN = conexao_.Execute( _
            "SELECT " & _
            "Sum(IIf(UCase(Trim([CHECK]))='SIM',1,0)) AS QTD_SIM, " & _
            "Sum(IIf(UCase(Trim([CHECK]))='NÃO',1,0)) AS QTD_NAO " & _
            "FROM TBL_CHECK_INFORMACAO WHERE COD=" & codH)

        qSim = 0 : qNao = 0
        If Not rsSN.EOF Then
            qSim = NumSafe(rsSN("QTD_SIM"))
            qNao = NumSafe(rsSN("QTD_NAO"))
        End If
        rsSN.Close

        If Not primeiroH Then itens = itens & ","
        itens = itens & "{" & _
            """cod"":" & codH & "," & _
            """data"":""" & JS(dataH) & """," & _
            """titulo"":""" & JS(titH) & """," & _
            """tipo"":""" & JS(tipH) & """," & _
            """autor"":""" & JS(autH) & """," & _
            """sim"":" & qSim & "," & _
            """nao"":" & qNao & _
        "}"
        primeiroH = False
        rsHC.MoveNext
    Loop
    rsHC.Close

    Response.Write "{" & _
        """pagina"":" & pg & "," & _
        """total_paginas"":" & totalPgH & "," & _
        """total_registros"":" & totalH & "," & _
        """itens"":[" & itens & "]" & _
    "}"

' ------------------------------------------------------------
' POR_REGISTRO – pie SIM/NAO + lista de leitores de 1 registro
' ------------------------------------------------------------
Case "por_registro"

    If filtCod = "" Or Not IsNumeric(filtCod) Then
        Response.Write "{""erro"":""Informe cod valido""}"
        Response.End
    End If
    codPR = CLng(filtCod)

    ' Titulo do registro
    Set rsTit = conexao_.Execute( _
        "SELECT TITULO, DATA, TIPO FROM TBL_INFORMACAO WHERE COD=" & codPR)
    tituloPR = "" : dataPR = "" : tipoPR = ""
    If Not rsTit.EOF Then
        tituloPR = "" & rsTit("TITULO")
        dataPR   = "" & rsTit("DATA")
        tipoPR   = "" & rsTit("TIPO")
    End If
    rsTit.Close

    ' Contagem SIM / NAO
    Set rsCnt = conexao_.Execute( _
        "SELECT " & _
        "Sum(IIf(UCase(Trim([CHECK]))='SIM',1,0)) AS QTD_SIM, " & _
        "Sum(IIf(UCase(Trim([CHECK]))='NÃO',1,0)) AS QTD_NAO " & _
        "FROM TBL_CHECK_INFORMACAO WHERE COD=" & codPR)
    qSimPR = 0 : qNaoPR = 0
    If Not rsCnt.EOF Then
        qSimPR = NumSafe(rsCnt("QTD_SIM"))
        qNaoPR = NumSafe(rsCnt("QTD_NAO"))
    End If
    rsCnt.Close

    ' Lista de quem JA leu (SIM) com data e dias corridos
    SQL_SIM = "SELECT TBL_CHECK_INFORMACAO.MATRICULA, TBL_USUARIO.NOME, " & _
              "TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO, " & _
              "[TBL_CHECK_INFORMACAO].[DATA_VISUALIZACAO]-[TBL_INFORMACAO].[DATA] AS DIAS_CORRIDOS " & _
              "FROM (TBL_CHECK_INFORMACAO " & _
              "INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA) " & _
              "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
              "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='SIM' " & _
              "AND TBL_CHECK_INFORMACAO.COD=" & codPR & _
              " ORDER BY TBL_CHECK_INFORMACAO.DATA_VISUALIZACAO ASC"

    Set rsSim = conexao_.Execute(SQL_SIM)
    itensSim = ""
    primeiroSim = True
    Do Until rsSim.EOF
        If Not primeiroSim Then itensSim = itensSim & ","
        diasC = rsSim("DIAS_CORRIDOS")
        If IsNumeric(diasC) Then diasC = CLng(diasC) Else diasC = 0
        itensSim = itensSim & "{" & _
            """matricula"":""" & JS(rsSim("MATRICULA")) & """," & _
            """nome"":""" & JS(rsSim("NOME")) & """," & _
            """data_visualizacao"":""" & JS(rsSim("DATA_VISUALIZACAO")) & """," & _
            """dias_corridos"":" & diasC & _
        "}"
        primeiroSim = False
        rsSim.MoveNext
    Loop
    rsSim.Close

    ' Lista de pendentes (NAO)
    SQL_NAO = "SELECT TBL_CHECK_INFORMACAO.MATRICULA, TBL_USUARIO.NOME, " & _
              "Date()-[TBL_INFORMACAO].[DATA] AS DIAS_PENDENTE " & _
              "FROM (TBL_CHECK_INFORMACAO " & _
              "INNER JOIN TBL_USUARIO ON TBL_CHECK_INFORMACAO.MATRICULA = TBL_USUARIO.MATRICULA) " & _
              "INNER JOIN TBL_INFORMACAO ON TBL_CHECK_INFORMACAO.COD = TBL_INFORMACAO.COD " & _
              "WHERE UCase(Trim(TBL_CHECK_INFORMACAO.[CHECK]))='NÃO' " & _
              "AND TBL_CHECK_INFORMACAO.COD=" & codPR & _
              " ORDER BY DIAS_PENDENTE DESC"

    Set rsNao = conexao_.Execute(SQL_NAO)
    itensNao = ""
    primeiroNao = True
    Do Until rsNao.EOF
        If Not primeiroNao Then itensNao = itensNao & ","
        diasP = rsNao("DIAS_PENDENTE")
        If IsNumeric(diasP) Then diasP = CLng(diasP) Else diasP = 0
        itensNao = itensNao & "{" & _
            """matricula"":""" & JS(rsNao("MATRICULA")) & """," & _
            """nome"":""" & JS(rsNao("NOME")) & """," & _
            """dias_pendente"":" & diasP & _
        "}"
        primeiroNao = False
        rsNao.MoveNext
    Loop
    rsNao.Close

    Response.Write "{" & _
        """cod"":" & codPR & "," & _
        """titulo"":""" & JS(tituloPR) & """," & _
        """data"":""" & JS(dataPR) & """," & _
        """tipo"":""" & JS(tipoPR) & """," & _
        """sim"":" & qSimPR & "," & _
        """nao"":" & qNaoPR & "," & _
        """leitores"":[" & itensSim & "]," & _
        """pendentes"":[" & itensNao & "]" & _
    "}"

' ------------------------------------------------------------
' DEFAULT
' ------------------------------------------------------------
Case Else
    Response.Write "{""erro"":""Acao invalida. Use: resumo, leituras_dia, por_tipo, tempo_medio, pendencias_user, historico, por_registro""}"

End Select
%>

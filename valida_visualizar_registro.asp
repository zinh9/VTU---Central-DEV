<!-- #include file="conexao.asp" -->

<%
If Session("matricula") = "" Then
    Response.Redirect("form_login.asp")
    Response.End
End If

COD_ = Trim("" & Request.Form("inputCOD"))
MATRICULA_ = "" & Session("matricula")

If COD_ = "" Or Not IsNumeric(COD_) Then
    Response.Write "Código inválido."
    Response.End
End If

DATA_ = Right("0" & Day(Date), 2) & "/" & Right("0" & Month(Date), 2) & "/" & Year(Date)

SQL_1 = "UPDATE TBL_CHECK_INFORMACAO " & _
        "SET [CHECK]='SIM', DATA_VISUALIZACAO='" & DATA_ & "' " & _
        "WHERE MATRICULA='" & Replace(MATRICULA_, "'", "''") & "' " & _
        "AND COD=" & CLng(COD_)

conexao_.Execute(SQL_1)

Response.Redirect("form_visualizar_registro.asp")
Response.End
%>
<%
' ============================================================
'  escopo_usuarios.asp
'  Inclua este arquivo em qualquer tela que precise filtrar
'  empregados pelo escopo do usuário logado.
'
'  USO:
'    <!-- #include file="escopo_usuarios.asp" -->
'
'    Dim where : where = GetEscopoWhere()
'    Dim sql   : sql   = "SELECT * FROM TBL_USUARIO WHERE 1=1 " & where
'
'  A função também expõe:
'    ESCOPO_NIVEL   → nível do usuário logado (string)
'    ESCOPO_MAT     → matrícula do usuário logado
'    ESCOPO_ACESSO  → True se pode acessar gestão de usuários
' ============================================================

Dim ESCOPO_NIVEL  : ESCOPO_NIVEL  = "" & Session("nivel")
Dim ESCOPO_MAT    : ESCOPO_MAT    = "" & Session("matricula")
Dim ESCOPO_ACESSO : ESCOPO_ACESSO = False

' Níveis que têm acesso à gestão de usuários
If ESCOPO_NIVEL = "ADM_DEV" Or _
   ESCOPO_NIVEL = "ADM_GG"  Or _
   ESCOPO_NIVEL = "ADM_GA"  Then
    ESCOPO_ACESSO = True
End If

' ────────────────────────────────────────────────────────────
' GetEscopoWhere()
'   Retorna cláusula WHERE (com AND inicial) filtrando
'   TBL_USUARIO pelo escopo do usuário logado.
'   Assume que a query principal já referencia TBL_USUARIO.
' ────────────────────────────────────────────────────────────
Function GetEscopoWhere()
    Dim mat : mat = Replace(ESCOPO_MAT, "'", "''")
    Dim w   : w   = ""

    Select Case ESCOPO_NIVEL

    ' ── DEV: acesso irrestrito ──────────────────────────────
    Case "ADM_DEV"
        w = ""

    ' ── GG: todos os usuários cuja COOR pertence à GA que
    '        pertence à GG cujo logado é o GG ───────────────
    '
    '  TBL_USUARIO.LOCALIDADE  → TBL_COOR.ID
    '  TBL_COOR.FK_GA          → TBL_GA.ID
    '  TBL_GA.FK_GG            → TBL_GG.ID
    '  TBL_GG.ID_TBL_GG        → matricula do GG em TBL_USUARIO
    '
    Case "ADM_GG"
        w = " AND TBL_USUARIO.LOCALIDADE IN (" & _
                "SELECT TBL_COOR.ID FROM TBL_COOR " & _
                "INNER JOIN TBL_GA  ON TBL_COOR.FK_GA  = TBL_GA.ID " & _
                "INNER JOIN TBL_GG  ON TBL_GA.FK_GG    = TBL_GG.ID " & _
                "WHERE TBL_GG.ID_TBL_GG = '" & mat & "'" & _
            ")"

    ' ── GA: todos os usuários cuja COOR pertence à GA
    '        cujo logado é o GA ─────────────────────────────
    '
    '  TBL_USUARIO.LOCALIDADE → TBL_COOR.ID
    '  TBL_COOR.FK_GA         → TBL_GA.ID
    '  TBL_GA.ID_TBL_GA       → matricula do GA em TBL_USUARIO
    '
    Case "ADM_GA"
        w = " AND TBL_USUARIO.LOCALIDADE IN (" & _
                "SELECT TBL_COOR.ID FROM TBL_COOR " & _
                "INNER JOIN TBL_GA ON TBL_COOR.FK_GA = TBL_GA.ID " & _
                "WHERE TBL_GA.ID_TBL_GA = '" & mat & "'" & _
            ")"

    ' ── ADM_DEL e USER: sem acesso ─────────────────────────
    Case Else
        ' Retorna cláusula impossível para bloquear qualquer resultado
        w = " AND 1=0"

    End Select

    ' Garante que GGs e GAs (sem LOCALIDADE) nunca apareçam
    ' na lista de empregados gerenciáveis
    If w <> " AND 1=0" Then
        w = w & " AND (TBL_USUARIO.NIVEL NOT IN ('ADM_GG','ADM_GA','ADM_DEV'))"
    End If

    GetEscopoWhere = w
End Function

' ────────────────────────────────────────────────────────────
' GetEscopoLabel()
'   Retorna texto descritivo do escopo para exibição na UI.
' ────────────────────────────────────────────────────────────
Function GetEscopoLabel()
    Select Case ESCOPO_NIVEL
        Case "ADM_DEV" : GetEscopoLabel = "Todos os empregados"
        Case "ADM_GG"  : GetEscopoLabel = "Empregados da sua GG"
        Case "ADM_GA"  : GetEscopoLabel = "Empregados da sua GA"
        Case Else      : GetEscopoLabel = ""
    End Select
End Function

' ────────────────────────────────────────────────────────────
' PodeEditarUsuario(matriculaAlvo)
'   Verifica se o usuário logado pode editar um usuário
'   específico. Use antes de qualquer operação de escrita.
' ────────────────────────────────────────────────────────────
Function PodeEditarUsuario(matriculaAlvo)
    PodeEditarUsuario = False

    If Not ESCOPO_ACESSO Then Exit Function

    ' DEV edita qualquer um
    If ESCOPO_NIVEL = "ADM_DEV" Then
        PodeEditarUsuario = True
        Exit Function
    End If

    ' Verifica se o alvo está dentro do escopo via banco
    Dim sqlVerifica : sqlVerifica = _
        "SELECT Count(*) AS T FROM TBL_USUARIO WHERE 1=1 " & _
        GetEscopoWhere() & _
        " AND TBL_USUARIO.MATRICULA = '" & Replace(matriculaAlvo, "'", "''") & "'"

    Dim rsV : Set rsV = conexao_.Execute(sqlVerifica)
    If Not rsV.EOF Then
        PodeEditarUsuario = (rsV("T") > 0)
    End If
    rsV.Close
    Set rsV = Nothing
End Function

' ────────────────────────────────────────────────────────────
' PodeDelegarNivel(nivelAlvo)
'   Verifica se o usuário logado pode atribuir um nível
'   específico a um empregado.
'   GG  → pode dar ADM_DEL e USER
'   GA  → pode dar ADM_DEL e USER
'   DEV → pode dar qualquer nível
' ────────────────────────────────────────────────────────────
Function PodeDelegarNivel(nivelAlvo)
    PodeDelegarNivel = False

    Select Case ESCOPO_NIVEL
        Case "ADM_DEV"
            PodeDelegarNivel = True
        Case "ADM_GG", "ADM_GA"
            ' Podem conceder apenas ADM_DEL ou USER
            ' Não podem promover outro GG/GA nem DEV
            If nivelAlvo = "ADM_DEL" Or nivelAlvo = "USER" Then
                PodeDelegarNivel = True
            End If
        Case Else
            PodeDelegarNivel = False
    End Select
End Function
%>

' ================================================================
' DIARIZAÇÃO v9.5 PRO — Plano Kaua
' Física Real 1:1 + Lote Mínimo 10 (ESTRITO)
' VALLEY-FILL: distribui sempre no dia mais vazio (anti-cascata)
' ================================================================
'
' CORREÇÃO v9.5 vs v9.4 (causa-raiz da cascata / dias abaixo):
'   As duas fases andavam da esquerda p/ direita (dia 1 -> N),
'   empilhando carga nos primeiros dias e afundando os do meio/fim.
'
'   AGORA ambas as fases ENCHEM O VALE PRIMEIRO:
'   1. FASE 1 (DistribuirNivelado): coloca cada lote cheio no dia
'      mais vazio disponivel (sob o alvoDia), respeitando cura e
'      teto. O espacamento por cura emerge naturalmente.
'   2. FASE 2 (PenteFino): distribui as sobras em rodadas, sempre
'      no dia mais baixo primeiro, subindo ate o teto.
'
'   Resultado: curva plana ~alvoDia, sem cascata.
'   Regra dos 10 ESTRITA: excecao SO se plano<10 OU qtdMoldes<10.
'   Cura medida em COLUNAS (k + diasCura >= diaIdx).
' ================================================================

Option Explicit

' ================================================================
' [1] CONSTANTES DE CONFIGURAÇÃO
' ================================================================
Private Const COL_SKU           As Integer = 1
Private Const COL_GRUPO         As Integer = 4
Private Const COL_MOLDE         As Integer = 5
Private Const COL_TIPO_CAP      As Integer = 9
Private Const COL_URGENCIA      As Integer = 16
Private Const COL_CAP           As Integer = 20
Private Const COL_PLANO         As Integer = 21
Private Const COL_QTD_MOLDES    As Integer = 24
Private Const COL_LEADTIME      As Integer = 32
Private Const COL_INICIO        As Integer = 33
Private Const COL_DIA_INI       As Integer = 34
Private Const COL_DIA_FIM       As Integer = 64
Private Const COL_LOG           As Integer = 70
Private Const LIN_CABEC         As Integer = 16
Private Const LIN_DADOS         As Integer = 17

Private Const TETO_PREF         As Double = 1400
Private Const TETO_GERAL        As Double = 1500
Private Const TETO_MAQ          As Double = 200
Private Const TETO_CAT_A        As Double = 700
Private Const TETO_CAT_B         As Double = 600
Private Const TETO_CAT_C         As Double = 150
Private Const TETO_CAT_D         As Double = 120

Private Const CAP_MAXIMA        As Double = 50
Private Const CAP_DEFAULT       As Double = 50
Private Const QTD_MOLDES_MIN    As Integer = 1
Private Const LT_DEFAULT        As Integer = 3
Private Const LOTE_MIN_PCP      As Double = 10

' ================================================================
' [2] TIPOS DE DADOS
' ================================================================
Private Type SKU_Data
    idx          As Long
    plano        As Double
    urgencia     As Double
    capMolde     As Double
    diasCura     As Integer
    leadtime     As Integer
    qtdMoldes    As Integer
    inicioTec    As Date
    molde        As String
    tipoCap      As String
    grupo        As String
    fatorGargalo As Double
    score        As Double
End Type

Private Type Dia_Data
    dt          As Date
    col         As Integer
    acGeral     As Double
    acMaq       As Double
    acCatA     As Double
    acCatB      As Double
    acCatC      As Double
    acCatD      As Double
End Type

' ================================================================
' [3] FUNÇÕES AUXILIARES
' ================================================================
Private Function NormStr(s As String) As String
    Dim r As String: r = LCase(s)
    r = Replace(r, Chr(225), "a"): r = Replace(r, Chr(224), "a")
    r = Replace(r, Chr(226), "a"): r = Replace(r, Chr(227), "a")
    r = Replace(r, Chr(233), "e"): r = Replace(r, Chr(234), "e")
    r = Replace(r, Chr(243), "o"): r = Replace(r, Chr(244), "o")
    r = Replace(r, Chr(250), "u"): r = Replace(r, Chr(237), "i")
    NormStr = Trim(r)
End Function

Private Function CeilLng(v As Double) As Long
    If v = Int(v) Then CeilLng = CLng(v) Else CeilLng = CLng(Int(v) + 1)
End Function

Private Function MinDbl(a As Double, B As Double) As Double
    If a <= B Then MinDbl = a Else MinDbl = B
End Function

Private Function EspacoDia(ByRef d As Dia_Data, grp As String, _
                            tip As String, limGeral As Double) As Double
    Dim esp As Double
    esp = limGeral - d.acGeral
    If tip = "maquina" Then esp = MinDbl(esp, TETO_MAQ - d.acMaq)
    If InStr(grp, "categoria a") > 0 Then esp = MinDbl(esp, TETO_CAT_A - d.acCatA)
    If InStr(grp, "categoria b") > 0 Then esp = MinDbl(esp, TETO_CAT_B - d.acCatB)
    If InStr(grp, "categoria c") > 0 Then esp = MinDbl(esp, TETO_CAT_C - d.acCatC)
    If InStr(grp, "categoria d") > 0 Then esp = MinDbl(esp, TETO_CAT_D - d.acCatD)
    If esp < 0 Then EspacoDia = 0 Else EspacoDia = esp
End Function

Private Function EhFundoVerde(cel As Range) As Boolean
    If cel.Interior.ColorIndex = -4142 Then EhFundoVerde = False: Exit Function
    Dim cIdx As Long: cIdx = cel.Interior.ColorIndex
    If cIdx = 4 Or cIdx = 35 Or cIdx = 43 Or cIdx = 14 Then
        EhFundoVerde = True: Exit Function
    End If
    Dim cor As Long: cor = cel.Interior.Color
    Dim r As Long, G As Long, B As Long
    r = cor Mod 256
    G = (cor \ 256) Mod 256
    B = (cor \ 65536) Mod 256
    EhFundoVerde = (G > r + 10 And G > B + 10)
End Function

Private Function MoldeLivreDoCicloAnterior(ByRef sk As SKU_Data, _
                                            diaIdx As Integer, _
                                            ByRef dias() As Dia_Data, _
                                            ByRef mat() As Double, _
                                            ByRef skus() As SKU_Data, _
                                            totalSKUs As Long) As Boolean
    Dim k As Integer, i As Long
    MoldeLivreDoCicloAnterior = False
    For k = 1 To diaIdx - 1
        For i = 1 To totalSKUs
            If skus(i).molde = sk.molde Then
                If mat(i, dias(k).col) > 0.001 Then
                    If k + skus(i).diasCura < diaIdx Then
                        MoldeLivreDoCicloAnterior = True: Exit Function
                    End If
                End If
            End If
        Next i
    Next k
End Function

' ----------------------------------------------------------------
' Ordena os indices dos dias do MAIS VAZIO ao mais cheio (acGeral).
' Insertion sort estavel. diaOrdem(1) = dia com menor carga.
' ----------------------------------------------------------------
Private Sub OrdenaDiasPorCarga(ByRef dias() As Dia_Data, nDias As Integer, _
                                ByRef diaOrdem() As Long)
    Dim i As Integer, j As Integer, tmp As Long
    For i = 1 To nDias: diaOrdem(i) = i: Next i
    For i = 2 To nDias
        tmp = diaOrdem(i)
        j = i - 1
        Do While j >= 1
            If dias(diaOrdem(j)).acGeral <= dias(tmp).acGeral Then Exit Do
            diaOrdem(j + 1) = diaOrdem(j)
            j = j - 1
        Loop
        diaOrdem(j + 1) = tmp
    Next i
End Sub
' ----------------------------------------------------------------
' [NOVO] MOTOR DE CAPACIDADE LIVRE (FORWARD WINDOW CHECKING)
' ----------------------------------------------------------------
Private Function CapacidadeLivrePecas(ByRef sk As SKU_Data, _
                                       diaIdx As Integer, _
                                       ByRef dias() As Dia_Data, _
                                       nDias As Integer, _
                                       ByRef mat() As Double, _
                                       ByRef skus() As SKU_Data, _
                                       totalSKUs As Long) As Double
    Dim i As Long, k_past As Integer, d_alvo As Integer, fimJanela As Integer
    Dim ocupados_d_alvo As Double, livres_d_alvo As Double
    Dim moldesLivresHoje As Double, capFisicaSobra As Double, poolMoldes As Double

    ' ----------------------------------------------------------------
    ' 1. POOL ÚNICO COMPARTILHADO DA FAMÍLIA
    ' ----------------------------------------------------------------
    poolMoldes = 0
    For i = 1 To totalSKUs
        If skus(i).molde = sk.molde Then
            If CDbl(skus(i).qtdMoldes) > poolMoldes Then poolMoldes = CDbl(skus(i).qtdMoldes)
        End If
    Next i
    If poolMoldes < 1 Then poolMoldes = CDbl(sk.qtdMoldes)

    ' ----------------------------------------------------------------
    ' 2. FORWARD WINDOW CHECKING (Blindagem Anti-Cascata no Valley-Fill)
    ' ----------------------------------------------------------------
    ' Como a alocação pula dias (não-cronológica), temos que olhar para o
    ' FUTURO do lote que estamos tentando alocar, cobrindo toda a sua cura.
    fimJanela = diaIdx + sk.diasCura
    If fimJanela > nDias Then fimJanela = nDias
    
    moldesLivresHoje = poolMoldes ' Assume o teto máximo e reduz no pior caso
    
    For d_alvo = diaIdx To fimJanela
        ocupados_d_alvo = 0
        
        ' Varre do passado até o d_alvo para ver quem está bloqueando moldes
        For k_past = 1 To d_alvo
            For i = 1 To totalSKUs
                If skus(i).molde = sk.molde Then
                    If mat(i, dias(k_past).col) > 0.000001 Then
                        ' Se a peça "i" foi produzida em "k_past", sua cura é k_past + diasCura.
                        ' Se isso alcançar ou invadir o d_alvo, o molde físico está ocupado.
                        If k_past + skus(i).diasCura >= d_alvo Then
                            ocupados_d_alvo = ocupados_d_alvo + mat(i, dias(k_past).col)
                        End If
                    End If
                End If
            Next i
        Next k_past
        
        livres_d_alvo = poolMoldes - ocupados_d_alvo
        If livres_d_alvo < 0 Then livres_d_alvo = 0
        
        ' A capacidade teórica é ditada pelo pior dia (maior gargalo) da janela
        If livres_d_alvo < moldesLivresHoje Then
            moldesLivresHoje = livres_d_alvo
        End If
    Next d_alvo

    ' ----------------------------------------------------------------
    ' 3. TRAVA DE CAPACIDADE DO DIA ALVO (Gargalo produtivo apenas para o dia alvo inicial)
    ' ----------------------------------------------------------------
    capFisicaSobra = sk.capMolde - mat(sk.idx, dias(diaIdx).col)
    If capFisicaSobra < 0 Then capFisicaSobra = 0

    CapacidadeLivrePecas = MinDbl(moldesLivresHoje, capFisicaSobra)
End Function

Private Sub Alocar(lote As Double, ByRef d As Dia_Data, ByRef mat() As Double, _
                   idxSKU As Long, grp As String, tip As String)
    mat(idxSKU, d.col) = mat(idxSKU, d.col) + lote
    d.acGeral = d.acGeral + lote
    If tip = "maquina" Then d.acMaq = d.acMaq + lote
    If InStr(grp, "categoria a") > 0 Then d.acCatA = d.acCatA + lote
    If InStr(grp, "categoria b") > 0 Then d.acCatB = d.acCatB + lote
    If InStr(grp, "categoria c") > 0 Then d.acCatC = d.acCatC + lote
    If InStr(grp, "categoria d") > 0 Then d.acCatD = d.acCatD + lote
End Sub

Private Function MotivoBloqueio(ByRef dias() As Dia_Data, nDias As Integer, _
                                 ByRef sk As SKU_Data, dataFim As Date, _
                                 ehUrgente As Boolean) As String
    Dim k As Integer, semDias As Boolean
    Dim bloqGeral As Boolean, bloqGeralPref As Boolean
    Dim bloqMaq As Boolean, bloqCatA As Boolean
    Dim bloqCatB As Boolean, bloqCatC As Boolean, bloqCatD As Boolean
    Dim motivos As String

    semDias = True
    For k = 1 To nDias
        If dias(k).dt >= sk.inicioTec And dias(k).dt <= dataFim Then
            semDias = False
            If dias(k).acGeral >= TETO_GERAL - 0.01 Then bloqGeral = True
            If dias(k).acGeral >= TETO_PREF - 0.01 Then bloqGeralPref = True
            If sk.tipoCap = "maquina" And dias(k).acMaq >= TETO_MAQ - 0.01 Then bloqMaq = True
            If InStr(sk.grupo, "categoria a") > 0 And dias(k).acCatA >= TETO_CAT_A - 0.01 Then bloqCatA = True
            If InStr(sk.grupo, "categoria b") > 0 And dias(k).acCatB >= TETO_CAT_B - 0.01 Then bloqCatB = True
            If InStr(sk.grupo, "categoria c") > 0 And dias(k).acCatC >= TETO_CAT_C - 0.01 Then bloqCatC = True
            If InStr(sk.grupo, "categoria d") > 0 And dias(k).acCatD >= TETO_CAT_D - 0.01 Then bloqCatD = True
        End If
    Next k

    If semDias Then
        MotivoBloqueio = "Sem dias disponiveis a partir do Inicio Tecnico"
        Exit Function
    End If
    motivos = ""
    If bloqGeral Or bloqGeralPref Then motivos = motivos & "Teto Fabrica Atingido; "
    If bloqMaq Then motivos = motivos & "Teto Maquina Atingido; "
    If bloqCatA Or bloqCatB Or bloqCatC Or bloqCatD Then motivos = motivos & "Teto Setor Atingido; "
    If motivos = "" Then motivos = "Falta Molde (cura) - plano alto p/ poucos moldes; "
    MotivoBloqueio = Left(motivos, Len(motivos) - 2)
End Function

' ================================================================
' [4] ROTINAS DE ETAPA
' ================================================================
Private Function CarregarCalendario(ws As Worksheet, dias() As Dia_Data) As Integer
    Dim k As Integer, vData As Variant, n As Integer: n = 0
    Dim i As Integer, j As Integer, tmp As Dia_Data
    For k = COL_DIA_INI To COL_DIA_FIM
        vData = ws.Cells(LIN_CABEC, k).Value
        If IsDate(vData) Then
            n = n + 1
            dias(n).dt = CDate(vData)
            dias(n).col = k
        End If
    Next k
    For i = 2 To n
        tmp = dias(i): j = i - 1
        Do While j >= 1 And dias(j).dt > tmp.dt
            dias(j + 1) = dias(j): j = j - 1
        Loop
        dias(j + 1) = tmp
    Next i
    CarregarCalendario = n
End Function

Private Sub CarregarSKUs(ws As Worksheet, totalSKUs As Long, dataIni As Date, _
                          skus() As SKU_Data, saldo() As Double)
    Dim i As Long, r As Long, valVar As Variant
    Dim baseScore As Double

    For i = 1 To totalSKUs
        r = LIN_DADOS + i - 1
        With skus(i)
            .idx = i
            valVar = ws.Cells(r, COL_PLANO).Value
            .plano = IIf(IsNumeric(valVar), CDbl(valVar), 0)
            valVar = ws.Cells(r, COL_URGENCIA).Value
            .urgencia = IIf(IsNumeric(valVar), CDbl(valVar), 0)
            valVar = ws.Cells(r, COL_CAP).Value
            .capMolde = IIf(IsNumeric(valVar) And CDbl(valVar) > 0, CDbl(valVar), CAP_DEFAULT)
            If .capMolde > CAP_MAXIMA Then .capMolde = CAP_MAXIMA
            valVar = ws.Cells(r, COL_LEADTIME).Value
            .diasCura = IIf(IsNumeric(valVar) And CInt(valVar) > 0, CInt(valVar), LT_DEFAULT - 1)
            If .diasCura <= 0 Then .diasCura = 2
            .leadtime = .diasCura + 1
            valVar = ws.Cells(r, COL_QTD_MOLDES).Value
            .qtdMoldes = IIf(IsNumeric(valVar) And CInt(valVar) > 0, CInt(valVar), QTD_MOLDES_MIN)
            If .qtdMoldes <= 0 Then .qtdMoldes = QTD_MOLDES_MIN
            .molde = Trim(UCase(CStr(ws.Cells(r, COL_MOLDE).Value)))
            .tipoCap = NormStr(CStr(ws.Cells(r, COL_TIPO_CAP).Value))
            .grupo = NormStr(CStr(ws.Cells(r, COL_GRUPO).Value))
            valVar = ws.Cells(r, COL_INICIO).Value
            .inicioTec = IIf(IsDate(valVar), CDate(valVar), dataIni)
            If .inicioTec < dataIni Then .inicioTec = dataIni
            If .plano > 0 Then
                .fatorGargalo = (.plano / .qtdMoldes) * .leadtime
            Else
                .fatorGargalo = 0
            End If
            baseScore = 0
            If .urgencia > 0 Then baseScore = 1000000000 + (.urgencia * 1000000)
            If .tipoCap = "maquina" Then baseScore = baseScore + 5000000
            baseScore = baseScore + (.plano * 10) + .fatorGargalo
            .score = baseScore
        End With
        saldo(i) = skus(i).plano
    Next i
End Sub

Private Sub OrdenagemScore(skus() As SKU_Data, totalSKUs As Long, ordem() As Long)
    Dim i As Long, j As Long, tmpOrd As Long
    Dim scoreA As Double, scoreB As Double
    For i = 1 To totalSKUs: ordem(i) = i: Next i
    For i = 2 To totalSKUs
        tmpOrd = ordem(i): scoreA = skus(tmpOrd).score: j = i - 1
        Do While j >= 1
            scoreB = skus(ordem(j)).score
            If scoreA <= scoreB Then Exit Do
            ordem(j + 1) = ordem(j): j = j - 1
        Loop
        ordem(j + 1) = tmpOrd
    Next i
End Sub

' ================================================================
' [FASE 1 - v9.5] DISTRIBUIR NIVELADO (valley-fill ate alvoDia)
' ================================================================
'
' Para cada peca (por prioridade), coloca seus lotes cheios SEMPRE
' no dia mais vazio disponivel (sob o alvoDia), respeitando cura,
' teto de setor e a regra dos 10. Re-ordena os dias por carga apos
' CADA lote -> a peca sempre busca o vale atual.
'
' O espacamento por cura emerge sozinho: se o vale escolhido fica
' com o molde curando, CapacidadeLivrePecas zera aquele dia nas
' proximas tentativas, e a peca migra para o proximo vale livre.
'
' Quando nao ha mais vale sob o alvoDia para esta peca, o saldo
' restante e deixado para a FASE 2.
'
Private Sub DistribuirNivelado(skus() As SKU_Data, totalSKUs As Long, _
                                dias() As Dia_Data, nDias As Integer, _
                                ordem() As Long, saldo() As Double, _
                                ByRef mat() As Double, _
                                ByRef travado() As Boolean, _
                                alvoDia As Double)
    Dim i As Long, idxSKU As Long
    Dim loteFisico As Double
    Dim diaOrdem() As Long
    Dim kk As Integer, k As Integer
    Dim espacoDisp As Double, capLivre As Double, lote As Double
    Dim ehUrgente As Boolean, colocou As Boolean
    Dim isOrderEx As Boolean, isMoldEx As Boolean

    ReDim diaOrdem(1 To nDias)

    For i = 1 To totalSKUs
        idxSKU = ordem(i)
        If saldo(idxSKU) < 0.01 Then GoTo ProxSKU

        ehUrgente = (skus(idxSKU).urgencia > 0)
        isOrderEx = (skus(idxSKU).plano < LOTE_MIN_PCP)
        isMoldEx = (skus(idxSKU).qtdMoldes < LOTE_MIN_PCP)

        loteFisico = MinDbl(skus(idxSKU).capMolde, CDbl(skus(idxSKU).qtdMoldes))
        If loteFisico < 1 Then loteFisico = 1

        ' Coloca um lote por vez, sempre no vale atual, ate o saldo
        ' acabar ou nao haver mais vale sob o alvoDia.
        Do While saldo(idxSKU) >= 0.01
            Call OrdenaDiasPorCarga(dias, nDias, diaOrdem)
            colocou = False

            For kk = 1 To nDias
                k = diaOrdem(kk)

                ' Dias ordenados por carga crescente: ao achar o
                ' primeiro >= alvoDia, todos os demais tambem estao
                ' -> nao ha mais vale para nivelar.
                If dias(k).acGeral >= alvoDia - 0.01 Then Exit For

                If travado(idxSKU, dias(k).col) Then GoTo ProxDia

                espacoDisp = EspacoDia(dias(k), skus(idxSKU).grupo, skus(idxSKU).tipoCap, _
                                       IIf(ehUrgente, TETO_GERAL, TETO_PREF))
                ' Nesta fase nao passa do alvoDia (so nivela o vale)
                espacoDisp = MinDbl(espacoDisp, alvoDia - dias(k).acGeral)
                If espacoDisp < 0.01 Then GoTo ProxDia

                capLivre = CapacidadeLivrePecas(skus(idxSKU), k, dias, nDias, _
                                                mat, skus, totalSKUs)
                If capLivre < 0.01 Then GoTo ProxDia

                lote = MinDbl(loteFisico, MinDbl(espacoDisp, MinDbl(capLivre, saldo(idxSKU))))

                ' REGRA DOS 10 ESTRITA (lote novo)
                If Not (isOrderEx Or isMoldEx) Then
                    If lote > 0 And lote < LOTE_MIN_PCP Then lote = 0
                End If

                If lote >= 0.01 Then
                    Call Alocar(lote, dias(k), mat, idxSKU, _
                                skus(idxSKU).grupo, skus(idxSKU).tipoCap)
                    saldo(idxSKU) = saldo(idxSKU) - lote
                    colocou = True
                    Exit For   ' coloca UM lote e re-ordena (vai pro proximo vale)
                End If
ProxDia:
            Next kk

            If Not colocou Then Exit Do   ' sem vale sob alvoDia -> FASE 2
        Loop
ProxSKU:
    Next i
End Sub

' ================================================================
' [FASE 2 - v9.5] PENTE-FINO / SOBRAS (valley-fill ate o teto)
' ================================================================
'
' Distribui o saldo residual em RODADAS. A cada rodada, ordena os
' dias por carga e coloca no maximo UM lote por dia (do vale ao
' pico), crescendo lotes existentes ou abrindo dia novo >=10.
' Re-ordena a cada rodada -> os dias mais baixos sobem primeiro.
' SOMENTE onde ha molde livre.
'
Private Function PenteFino(skus() As SKU_Data, totalSKUs As Long, _
                            dias() As Dia_Data, nDias As Integer, _
                            ordem() As Long, saldo() As Double, _
                            ByRef mat() As Double, _
                            ByRef travado() As Boolean) As Long
    Dim qtdPreench As Long
    Dim diaOrdem() As Long
    Dim rodada As Integer, maxRodadas As Integer, colocouRodada As Boolean
    Dim kk As Integer, k As Integer, i As Long, idxSKU As Long
    Dim espacoDisp As Double, capLivre As Double, room As Double, lote As Double
    Dim loteFisico As Double
    Dim ehUrgente As Boolean, jaTemLote As Boolean, colocouDia As Boolean
    Dim isOrderEx As Boolean, isMoldEx As Boolean

    qtdPreench = 0
    ReDim diaOrdem(1 To nDias)
    maxRodadas = nDias * 4   ' teto de seguranca contra loop infinito

    rodada = 0
    Do
        rodada = rodada + 1
        colocouRodada = False
        Call OrdenaDiasPorCarga(dias, nDias, diaOrdem)

        For kk = 1 To nDias
            k = diaOrdem(kk)
            colocouDia = False

            For i = 1 To totalSKUs
                If colocouDia Then Exit For
                idxSKU = ordem(i)
                If saldo(idxSKU) < 0.01 Then GoTo ProxSKUPF
                If travado(idxSKU, dias(k).col) Then GoTo ProxSKUPF

                ehUrgente = (skus(idxSKU).urgencia > 0)

                espacoDisp = EspacoDia(dias(k), skus(idxSKU).grupo, skus(idxSKU).tipoCap, _
                                       IIf(ehUrgente, TETO_GERAL, TETO_PREF))
                If espacoDisp < 0.01 Then GoTo ProxSKUPF

                capLivre = CapacidadeLivrePecas(skus(idxSKU), k, dias, nDias, _
                                                mat, skus, totalSKUs)
                If capLivre < 0.01 Then GoTo ProxSKUPF

                ' Limita o incremento a UM lote fisico por visita
                ' (leveling suave: o dia sobe aos poucos por rodada)
                loteFisico = MinDbl(skus(idxSKU).capMolde, CDbl(skus(idxSKU).qtdMoldes))
                If loteFisico < 1 Then loteFisico = 1

                room = MinDbl(loteFisico, MinDbl(espacoDisp, MinDbl(saldo(idxSKU), capLivre)))
                If room < 0.01 Then GoTo ProxSKUPF

                jaTemLote = (mat(idxSKU, dias(k).col) > 0.001)
                isOrderEx = (skus(idxSKU).plano < LOTE_MIN_PCP)
                isMoldEx = (skus(idxSKU).qtdMoldes < LOTE_MIN_PCP)

                lote = 0
                If jaTemLote Then
                    lote = room                       ' (A) cresce lote existente
                Else
                    If isOrderEx Or isMoldEx Then
                        lote = room
                    ElseIf room >= LOTE_MIN_PCP Then
                        lote = room                   ' (B) abre dia novo >=10
                    Else
                        lote = 0
                    End If
                End If

                If lote >= 0.01 Then
                    Call Alocar(lote, dias(k), mat, idxSKU, _
                                skus(idxSKU).grupo, skus(idxSKU).tipoCap)
                    saldo(idxSKU) = saldo(idxSKU) - lote
                    qtdPreench = qtdPreench + 1
                    colocouRodada = True
                    colocouDia = True
                End If
ProxSKUPF:
            Next i
        Next kk
    Loop While colocouRodada And rodada < maxRodadas

    PenteFino = qtdPreench
End Function

' ================================================================
' [4b] GRAVAÇÃO E SUMÁRIO
' ================================================================
Private Sub GravarResultado(ws As Worksheet, totalSKUs As Long, _
                              ultimaLinha As Long, ByRef mat() As Double)
    Dim i As Long, k As Integer, nCols As Integer
    Dim bloco() As Variant
    nCols = COL_DIA_FIM - COL_DIA_INI + 1
    ReDim bloco(1 To totalSKUs, 1 To nCols)
    For i = 1 To totalSKUs
        For k = COL_DIA_INI To COL_DIA_FIM
            If mat(i, k) > 0.001 Then
                bloco(i, k - COL_DIA_INI + 1) = Round(mat(i, k), 2)
            End If
        Next k
    Next i
    ws.Cells(LIN_DADOS, COL_DIA_INI).Resize(totalSKUs, nCols).Value = bloco
End Sub

Private Sub GravarLogMotivos(ws As Worksheet, totalSKUs As Long, _
                               ultimaLinha As Long, skus() As SKU_Data, _
                               saldo() As Double, dias() As Dia_Data, _
                               nDias As Integer, dataFim As Date)
    Dim i As Long, r As Long, ehUrgente As Boolean
    ws.Cells(LIN_CABEC, COL_LOG).Value = "Motivo Residual"
    ws.Range(ws.Cells(LIN_DADOS, COL_LOG), ws.Cells(ultimaLinha, COL_LOG)).ClearContents
    For i = 1 To totalSKUs
        If saldo(i) > 0.01 Then
            r = LIN_DADOS + i - 1
            ehUrgente = (skus(i).urgencia > 0)
            ws.Cells(r, COL_LOG).Value = MotivoBloqueio(dias, nDias, skus(i), dataFim, ehUrgente)
        End If
    Next i
End Sub

Private Sub MontarSumario(skus() As SKU_Data, totalSKUs As Long, _
                           dias() As Dia_Data, nDias As Integer, _
                           ByRef mat() As Double, saldo() As Double, _
                           preench As Long, alvoDia As Double)
    Dim i As Long, j As Integer, k As Integer
    Dim somaSKU As Double, totalAloc As Double, totalRes As Double
    Dim skusRes As Long, diasAcima As Long, diasAcimaPref As Long, mediaGeral As Double
    Dim diasAbaixo1000 As Long, diaMin As Double, diaMax As Double

    diaMin = 1E+18: diaMax = 0
    For i = 1 To totalSKUs
        somaSKU = 0
        For k = COL_DIA_INI To COL_DIA_FIM: somaSKU = somaSKU + mat(i, k): Next k
        totalAloc = totalAloc + somaSKU
        If saldo(i) > 0.01 Then
            totalRes = totalRes + saldo(i)
            skusRes = skusRes + 1
        End If
    Next i
    For j = 1 To nDias
        If dias(j).acGeral > TETO_GERAL Then diasAcima = diasAcima + 1
        If dias(j).acGeral > TETO_PREF Then diasAcimaPref = diasAcimaPref + 1
        If dias(j).acGeral < 1000 Then diasAbaixo1000 = diasAbaixo1000 + 1
        If dias(j).acGeral < diaMin Then diaMin = dias(j).acGeral
        If dias(j).acGeral > diaMax Then diaMax = dias(j).acGeral
    Next j
    If nDias > 0 Then mediaGeral = totalAloc / nDias Else mediaGeral = 0

    MsgBox "Diarizacao v9.5 PRO concluida! (Valley-Fill anti-cascata)" & vbCrLf & vbCrLf & _
           "  Alvo nivelado/dia     : " & Format(alvoDia, "#,##0") & vbCrLf & _
           "  Total alocado         : " & Format(totalAloc, "#,##0") & " unidades" & vbCrLf & _
           "  Residual de meta      : " & Format(totalRes, "#,##0") & " unidades" & vbCrLf & _
           "  SKUs com residual     : " & skusRes & vbCrLf & _
           "  Media diaria          : " & Format(mediaGeral, "#,##0.0") & vbCrLf & _
           "  Dia mais BAIXO        : " & Format(diaMin, "#,##0") & vbCrLf & _
           "  Dia mais ALTO         : " & Format(diaMax, "#,##0") & vbCrLf & _
           "  Dias abaixo de 1000   : " & diasAbaixo1000 & vbCrLf & _
           "  Dias acima de " & TETO_PREF & "    : " & diasAcimaPref & vbCrLf & _
           "  Dias acima de " & TETO_GERAL & "   : " & diasAcima & vbCrLf & _
           "  Aloc. sobras (FASE 2) : " & preench & vbCrLf & _
           "  Log de motivos em BP  : " & skusRes & " linhas preenchidas", _
           vbInformation, "Diarizacao v9.5"
End Sub

' ================================================================
' [5] ROTINA PRINCIPAL
' ================================================================
Sub Diarização()
    Dim ws As Worksheet, ultimaLinha As Long, totalSKUs As Long, nDias As Integer
    Dim dataIni As Date, dataFim As Date, i As Long, preench As Long
    Dim skus() As SKU_Data, dias() As Dia_Data, ordem() As Long
    Dim alocMat() As Double, saldo() As Double, travado() As Boolean
    Dim cel As Range, valMan As Double, dIdx As Integer, d As Integer
    Dim k As Integer
    Dim somaPlano As Double, alvoDia As Double

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets("Mini APS")
    ultimaLinha = ws.Cells(ws.Rows.Count, COL_SKU).End(xlUp).Row
    If ultimaLinha < LIN_DADOS Then GoTo Finalizar

    totalSKUs = ultimaLinha - LIN_DADOS + 1
    ReDim dias(1 To COL_DIA_FIM - COL_DIA_INI + 1)
    nDias = CarregarCalendario(ws, dias)
    If nDias = 0 Then GoTo Finalizar

    dataIni = dias(1).dt
    dataFim = dias(nDias).dt

    ReDim skus(1 To totalSKUs)
    ReDim saldo(1 To totalSKUs)
    ReDim alocMat(1 To totalSKUs, COL_DIA_INI To COL_DIA_FIM)
    ReDim travado(1 To totalSKUs, COL_DIA_INI To COL_DIA_FIM)

    Call CarregarSKUs(ws, totalSKUs, dataIni, skus, saldo)
    ReDim ordem(1 To totalSKUs)
    Call OrdenagemScore(skus, totalSKUs, ordem)

    ' ----------------------------------------------------------------
    ' ALVO NIVELADO = PlanoTotal / nDias (a "linha media" da fabrica)
    ' E o limite SOFT da FASE 1. A FASE 2 sobe acima dele se sobrar.
    ' ----------------------------------------------------------------
    somaPlano = 0
    For i = 1 To totalSKUs: somaPlano = somaPlano + skus(i).plano: Next i
    If nDias > 0 Then alvoDia = somaPlano / CDbl(nDias) Else alvoDia = somaPlano
    alvoDia = Round(alvoDia, 0)
    If alvoDia > TETO_PREF Then alvoDia = TETO_PREF
    If alvoDia < LOTE_MIN_PCP Then alvoDia = LOTE_MIN_PCP

    ' ----------------------------------------------------------------
    ' FASE 0: Scanner de Verde (Ordens Firmes Manuais)
    ' ----------------------------------------------------------------
    For i = 1 To totalSKUs
        For k = COL_DIA_INI To COL_DIA_FIM
            Set cel = ws.Cells(LIN_DADOS + i - 1, k)
            If EhFundoVerde(cel) Then
                travado(i, k) = True
                valMan = IIf(IsNumeric(cel.Value), CDbl(cel.Value), 0)
                If valMan > 0 Then
                    dIdx = 0
                    For d = 1 To nDias
                        If dias(d).col = k Then dIdx = d: Exit For
                    Next d
                    If dIdx > 0 Then
                        Call Alocar(valMan, dias(dIdx), alocMat, i, skus(i).grupo, skus(i).tipoCap)
                        saldo(i) = saldo(i) - valMan
                    End If
                End If
            Else
                cel.ClearContents
            End If
        Next k
    Next i
    DoEvents

    ' ----------------------------------------------------------------
    ' FASE 1: NIVELA enchendo o vale ate o alvoDia
    ' ----------------------------------------------------------------
    Call DistribuirNivelado(skus, totalSKUs, dias, nDias, ordem, saldo, _
                            alocMat, travado, alvoDia)

    ' ----------------------------------------------------------------
    ' FASE 2: SOBRAS — sobe os vales restantes ate o teto
    ' ----------------------------------------------------------------
    preench = PenteFino(skus, totalSKUs, dias, nDias, ordem, saldo, alocMat, travado)

    Call GravarResultado(ws, totalSKUs, ultimaLinha, alocMat)
    Call GravarLogMotivos(ws, totalSKUs, ultimaLinha, skus, saldo, dias, nDias, dataFim)
    Call MontarSumario(skus, totalSKUs, dias, nDias, alocMat, saldo, preench, alvoDia)

Finalizar:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    ws.Calculate
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Diarizacao v9.5"
End Sub

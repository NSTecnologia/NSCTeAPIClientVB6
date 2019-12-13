Attribute VB_Name = "CTeAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references
Public responseText As String

'Atributo privado da classe
Private Const tempoResposta = 500
Private Const token = "SEU_TOKEN"

Function enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain;charset=utf-8"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml;charset=utf-8"
    Else
        contentType = "application/json;charset=utf-8"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP60
    Set obj = New MSXML2.ServerXMLHTTP60
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    
    Select Case obj.status
        'Se o token não for enviado ou for inválido
        Case 401
            MsgBox ("Token não enviado ou inválido")
        'Se o token informado for inválido 403
        Case 403
            MsgBox ("Token sem permissão")
    End Select
    
    enviaConteudoParaAPI = resposta
    Exit Function
SAI:
  enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Emitir NF-e Síncrono
Public Function emitirCTeSincrono(conteudo As String, tpConteudo As String, CNPJ As String, tpDown As String, tpAmb As String, modelo As String, caminho As String, exibeNaTela As Boolean) As String
    Dim retorno As String
    Dim resposta As String
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim motivo As String
    Dim erros As String
    Dim nsNRec As String
    Dim chCTe As String
    Dim cStat As String
    Dim nProt As String

    status = ""
    motivo = ""
    erros = ""
    nsNRec = ""
    chCTe = ""
    cStat = ""
    nProt = ""

    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = emitirCTe(conteudo, tpConteudo)
    statusEnvio = LerDadosJSON(resposta, "status", "", "")

    If (statusEnvio = "200") Or (statusEnvio = "-6") Then
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")

        Sleep (tempoResposta)

        resposta = consultarStatusProcessamento(CNPJ, nsNRec, tpAmb)
        statusConsulta = LerDadosJSON(resposta, "status", "", "")

        If (statusConsulta = "200") Then
            cStat = LerDadosJSON(resposta, "cStat", "", "")

            If (cStat = "100") Or (cStat = "150") Then
                chCTe = LerDadosJSON(resposta, "chCTe", "", "")
                nProt = LerDadosJSON(resposta, "nProt", "", "")
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")

                resposta = downloadCTeESalvar(chCTe, tpAmb, tpDown, caminho, exibeNaTela)
                statusDownload = LerDadosJSON(resposta, "status", "", "")

                If (statusDownload <> "200") Then
                    motivo = LerDadosJSON(resposta, "motivo", "", "")
                End If
            Else
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")
            End If
        Else
            motivo = LerDadosJSON(resposta, "motivo", "", "")
        End If
        
    ElseIf (statusEnvio = "-7") Then
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")
   
    ElseIf (statusEnvio = "-4") Then

        motivo = LerDadosJSON(resposta, "motivo", "", "")
        erros = LerDadosJSON(resposta, "erros", "", "")

    ElseIf (statusEnvio = "-9") Then
        'Lê o objeto erro
        erros = Split(resposta, """erro"":""")
        erros = LerDadosJSON(resposta, "erro", "", "")
        
        motivo = LerDadosJSON(erros, "xMotivo", "", "")
        
        cStat = LerDadosJSON(erros, "cStat", "", "")
    Else
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        
    End If
    
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusConsulta"":""" & statusConsulta & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chCTe"":""" & chCTe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """nsNRec"":""" & nsNRec & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    'Grava dados de retorno
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")
    gravaLinhaLog ("")

    emitirCTeSincrono = retorno
End Function


'Emitir NF-e
Public Function emitirCTe(conteudo As String, tpConteudo As String) As String
    Dim url As String
    Dim resposta As String
    
    url = "https://cte.ns.eti.br/cte/issue"

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
    
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirCTe = resposta
End Function

'Consultar Status de Processamento
Public Function consultarStatusProcessamento(CNPJ As String, nsNRec As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    json = "{"
    json = json & """CNPJ"":""" & CNPJ & ""","
    json = json & """nsNRec"":""" & nsNRec & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/cte/issueStatus/300"
    
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusProcessamento = resposta
End Function

'Download do CT-e
Public Function downloadCTe(chCTe As String, tpDown As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String

    json = "{"
    json = json & """chCTe"":""" & chCTe & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/cte/get/300"

    gravaLinhaLog ("[DOWNLOAD_CTe_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status <> "200") Then
        gravaLinhaLog ("[DOWNLOAD_CTE_RESPOSTA]")
        gravaLinhaLog (resposta)
    Else
        gravaLinhaLog ("[DOWNLOAD_CTE_RESPOSTA]")
        gravaLinhaLog (status)
    End If

    downloadCTe = resposta
End Function

'Download do CT-e e Salvar
Public Function downloadCTeESalvar(chCTe As String, tpAmb As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadCTe(chCTe, tpDown, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
    
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
    
        'Checa se deve baixar XML
        If InStr(1, tpDown, "X") Then
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chCTe, "", "")
        End If
        'Checa se deve baixar JSON
        If InStr(1, tpDown, "J") Then
            Dim conteudoJSON() As String
            'Separa o JSON da CTe
            conteudoJSON = Split(resposta, """cteProc"":{")
            json = "{""cteProc"":{" & conteudoJSON(1)
            Call salvarJSON(json, caminho, chCTe, "", "")
        End If
        'Checa se deve baixar PDF
        If InStr(1, tpDown, "P") Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chCTe, "", "")
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & chCTe & "-procCTe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadCTeESalvar = resposta
End Function

'Download do Evento do CT-e
Public Function downloadEventoCTe(chCTe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    json = "{"
    json = json & """chCTe"":""" & chCTe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpEvento"":""" & tpEvento & ""","
    json = json & """nSeqEvento"":""" & nSeqEvento & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/cte/get/event/300"
    
    gravaLinhaLog ("[DOWNLOAD_EVENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status <> "200") Then
        gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
        gravaLinhaLog (resposta)
    Else
        gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
        gravaLinhaLog (status)
    End If

    downloadEventoCTe = resposta
End Function

'Download do Evento do CT-e e Salvar
Public Function downloadEventoCTeESalvar(chCTe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String, caminho As String, exibeNaTela As Boolean) As String
    Dim baixarXML As Boolean
    Dim baixarPDF As Boolean
    Dim baixarJSON As Boolean
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    Dim tpEventoSalvar As String

    resposta = downloadEventoCTe(chCTe, tpAmb, tpDown, tpEvento, nSeqEvento)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
        
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        If (UCase(tpEvento) = "CANC") Then
          tpEventoSalvar = "110111"
        Else
          tpEventoSalvar = "110110"
        End If
        
        'Checa se deve baixar XML
        If InStr(1, tpDown, "X") Then
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chCTe, tpEventoSalvar, nSeqEvento)
        End If
        'Checa se deve baixar JSON
        If InStr(1, tpDown, "J") Then
            json = LerDadosJSON(resposta, "json", "", "")
            Call salvarJSON(json, caminho, chCTe, tpEventoSalvar, nSeqEvento)
        End If
        'Checa se deve baixar PDF
        If InStr(1, tpDown, "P") Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chCTe, tpEventoSalvar, nSeqEvento)
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & tpEventoSalvar & chCTe & nSeqEvento & "-procEvenCTe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadEventoCTeESalvar = resposta
End Function

'Realizar o cancelamento do CT-e
Public Function cancelarCTe(chCTe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String

    json = "{"
    json = json & """chCTe"":""" & chCTe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"
    
    url = "https://cte.ns.eti.br/cte/cancel/300"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")

    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
        respostaDownload = downloadEventoCTeESalvar(chCTe, tpAmb, tpDown, "CANC", "1", caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "200") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    cancelarCTe = resposta
End Function

'Realizar a correção de um CT-e
Public Function corrigirCTe(chCTe As String, tpAmb As String, dhEvento As String, nSeqEvento As String, grupoAlterado As String, campoAlterado As String, valorAlterado As String, nroItemAlterado As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim infCorrecao As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    infCorrecao = "{"
    infCorrecao = infCorrecao & """grupoAlterado""" & grupoAlterado & ""","
    infCorrecao = infCorrecao & """campoAlterado""" & campoAlterado & ""","
    infCorrecao = infCorrecao & """valorAlterado""" & valorAlterado & ""","
    infCorrecao = infCorrecao & """nroItemAlterado""" & nroItemAlterado & ""
    infCorrecao = infCorrecao & "}"
    
    json = "{"
    json = json & """chCTe"":""" & chCTe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nSeqEvento"":""" & nSeqEvento & ""","
    json = json & """infCorrecao"":" & infCorrecao & ""
    json = json & "}"
    
    url = "https://cte.ns.eti.br/cte/cce"
    
    gravaLinhaLog ("[CCE_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")

    gravaLinhaLog ("[CCE_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
        respostaDownload = downloadEventoCTeESalvar(chCTe, tpAmb, tpDown, "CCE", nSeqEvento, caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "200") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    corrigirCTe = resposta
End Function

'Esta função realiza a consulta de cadastro de contribuinte
Public Function consultarCadastroContribuinte(CNPJCont As String, UF As String, documentoConsulta As String, tpConsulta As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """CNPJCont"":""" & CNPJCont & ""","
    json = json & """UF"":""" & UF & ""","
    json = json & """" & tpConsulta & """:""" & documentoConsulta & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/util/conscad"
    
    gravaLinhaLog ("[CONSULTA_CADASTRO_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_CADASTRO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarCadastroContribuinte = resposta
End Function

'Esta função realiza a consulta de situação de um CT-e
Public Function consultarSituacao(licencaCnpj As String, chCTe As String, tpAmb As String, versao As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """licencaCnpj"":""" & licencaCnpj & ""","
    json = json & """chCTe"":""" & chCTe & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/cte/stats/300"
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (json)

    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarSituacao = resposta
End Function

'Esta função realiza a inutilização de um intervalo de numeração de CT-e
Public Function inutilizar(cUF As String, tpAmb As String, ano As String, CNPJ As String, serie As String, nNFIni As String, nNFFin As String, xJust As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """cUF"":""" & cUF & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """ano"":""" & ano & ""","
    json = json & """CNPJ"":""" & CNPJ & ""","
    json = json & """serie"":""" & serie & ""","
    json = json & """nCTIni"":""" & nNFIni & ""","
    json = json & """nCTFin"":""" & nNFFin & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/cte/inut"
    
    gravaLinhaLog ("[INUTILIZACAO_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[INUTILIZACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    inutilizar = resposta
End Function

'Esta função faz a listagem de nsNRec vinculados a uma chave de CT-e
Public Function listarNSNRecs(chCTe As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chCTe"":""" & chCTe & """"
    json = json & "}"

    url = "https://cte.ns.eti.br/util/list/nsnrecs"
    
    gravaLinhaLog ("[LISTA_NSNRECS_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[LISTA_NSNRECS_RESPOSTA]")
    gravaLinhaLog (resposta)

    listarNSNRecs = resposta
End Function

'Salvar XML
Public Sub salvarXML(xml As String, caminho As String, chCTe As String, tpEvento As String, nSeqEvento As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo XML
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chCTe & nSeqEvento & "-procCTe.xml"
    Else
        localParaSalvar = caminho & tpEvento & chCTe & nSeqEvento & "-procEvenCTe.xml"
    End If

    'Remove as contrabarras
    conteudoSalvar = Replace(xml, "\", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Salvar JSON
Public Sub salvarJSON(json As String, caminho As String, chCTe As String, tpEvento As String, nSeqEvento As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo JSON
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chCTe & nSeqEvento & "-procCTe.json"
    Else
        localParaSalvar = caminho & tpEvento & chCTe & nSeqEvento & "-procEvenCTe.json"
    End If

    conteudoSalvar = json

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Salvar PDF
Public Function salvarPDF(pdf As String, caminho As String, chCTe As String, tpEvento As String, nSeqEvento As String) As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo PDF
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chCTe & nSeqEvento & "-procCTe.pdf"
    Else
        localParaSalvar = caminho & tpEvento & chCTe & nSeqEvento & "-procEvenCTe.pdf"
    End If

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

'activate microsoft script control 1.0 in references
Public Function LerDadosJSON(sJsonString As String, Key1 As String, Key2 As String, key3 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If Key1 <> "" And Key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet), key3, VbGet)
    ElseIf Key1 <> "" And Key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet)
    ElseIf Key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, Key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function

Public Function LerDadosXML(sXml As String, Key1 As String, Key2 As String) As String
    On Error Resume Next
    LerDadosXML = ""
    
    Set xml = New DOMDocument60
    xml.async = False
    
    If xml.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = xml.getElementsByTagName(Key1 & "//" & Key2)
        Set objNode = objNodeList.nextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            LerDadosXML = valor
        End If
        Else
        MsgBox "Não foi possível ler o conteúdo do XML do CTe especificado para leitura.", vbCritical, "ERRO"
    End If
End Function

'Função genérica para gravação de log
Public Sub gravaLinhaLog(conteudoSalvar As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim data As String
    
    'Diretório para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    'Pega data atual
    data = Format(Date, "yyyyMMdd")
    
    'Diretório + nome do arquivo para salvar os logs
    localParaSalvar = App.Path & "\log\" & data & ".txt"
    
    'Pega data e hora atual
    data = DateTime.Now
    
    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Append Shared As #fnum
    Print #fnum, data & " - " & conteudoSalvar
    Close fnum
End Sub


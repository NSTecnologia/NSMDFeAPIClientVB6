Attribute VB_Name = "MDFeAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references
Public responseText As String

'Atributo privado da classe
Private Const tempoResposta = 500
Private Const token = "4EB15D6DEDAEBAE3FD0B7B5E5B0AD6D4"

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
  enviaConteudoParaAPI = Err.Number & " " & Err.Description
End Function

'Emitir MDF-e Síncrono
Public Function emitirMDFeSincrono(conteudo As String, tpConteudo As String, CNPJ As String, tpDown As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    
    Dim retorno As String
    Dim resposta As String
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim motivo As String
    Dim erros As String
    Dim nsNRec As String
    Dim chMDFe As String
    Dim cStat As String
    Dim nProt As String

    'Inicia as variáveis vazias
    statusEnvio = ""
    statusConsulta = ""
    statusDownload = ""
    motivo = ""
    erros = ""
    nsNRec = ""
    chMDFe = ""
    cStat = ""
    nProt = ""

    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = emitirMDFe(conteudo, tpConteudo)
    statusEnvio = LerDadosJSON(resposta, "status", "", "")

    If (statusEnvio = "200") Or (statusEnvio = "-6") Then
    
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")

        Sleep (tempoEspera)

        resposta = consultarStatusProcessamento(CNPJ, nsNRec, tpAmb)
        statusConsulta = LerDadosJSON(resposta, "status", "", "")

        If (statusConsulta = "200") Then

            cStat = LerDadosJSON(resposta, "cStat", "", "")

            If (cStat = "100") Then

                chMDFe = LerDadosJSON(resposta, "chMDFe", "", "")
                nProt = LerDadosJSON(resposta, "nProt", "", "")
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")
                
                resposta = downloadMDFeESalvar(chMDFe, tpDown, tpAmb, caminho, exibeNaTela)
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

    ElseIf (statusEnvio = "-4") Or (statusEnvio = "-2") Then
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        erros = LerDadosJSON(resposta, "erros", "", "")

    ElseIf (statusEnvio = "-5") Then
    
        erros = Split(resposta, """erro"":""")
        erros = LerDadosJSON(resposta, "erro", "", "")
        erros = LerDadosJSON(erros, "xMotivo", "", "")
        cStat = LerDadosJSON(erros, "cStat", "", "")
        
    ElseIf (statusEnvio = "-999") Then
    
        erros = Split(resposta, """erro"":""")
        erros = LerDadosJSON(resposta, "erro", "", "")
        erros = LerDadosJSON(erros, "xMotivo", "", "")
    Else
 
        motivo = LerDadosJSON(resposta, "motivo", "", "")
    End If
    
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusConsulta"":""" & statusConsulta & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chMDFe"":""" & chMDFe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """nsNRec"":""" & nsNRec & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")

    emitirMDFeSincrono = retorno
End Function


'Emitir MDF-e
Public Function emitirMDFe(conteudo As String, tpConteudo As String) As String
    Dim urlEnvio As String
    Dim resposta As String

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
    
    urlEnvio = "https://mdfe.ns.eti.br/mdfe/issue"

    resposta = enviaConteudoParaAPI(conteudo, urlEnvio, tpConteudo)

    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirMDFe = resposta
End Function

'Consultar Status de Processamento
Public Function consultarStatusProcessamento(CNPJ As String, nsNRec As String, tpAmb As String) As String
    Dim json As String
    Dim urlEnvio As String
    Dim resposta As String

    json = "{"
    json = json & """CNPJ"":""" & CNPJ & ""","
    json = json & """nsNRec"":""" & nsNRec & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"
    
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (json)

    urlEnvio = "https://mdfe.ns.eti.br/mdfe/issue/status"

    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
    
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusProcessamento = resposta
End Function

'Download do MDF-e
Public Function downloadMDFe(chMDFe As String, tpDown As String, tpAmb As String) As String
    Dim json As String
    Dim urlEnvio As String
    Dim resposta As String

    json = "{"
    json = json & """chMDFe"":""" & chMDFe & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"
    
    gravaLinhaLog ("[DOWNLOAD_MDFE_DADOS]")
    gravaLinhaLog (json)

    urlEnvio = "https://mdfe.ns.eti.br/mdfe/get"

    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
    
    gravaLinhaLog ("[DOWNLOAD_MDFE_RESPOSTA]")
    gravaLinhaLog (resposta)

    downloadMDFe = resposta
End Function

'Download do MDF-e e Salvar
Public Function downloadMDFeESalvar(chMDFe As String, tpDown As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadMDFe(chMDFe, tpDown, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
    
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        'Checa se deve baixar XML
        If InStr(1, tpDown, "X") Then
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chMDFe, "", False)
        End If
        'Checa se deve baixar JSON
        If InStr(1, tpDown, "J") Then
            Dim conteudoJSON() As String
            'Separa o JSON da MDFe
            conteudoJSON = Split(resposta, """MDFeProc"":{")
            json = "{""MDFeProc"":{" & conteudoJSON(1)
            Call salvarJSON(json, caminho, chMDFe, "", False)
        End If
        'Checa se deve baixar PDF
        If InStr(1, tpDown, "P") Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chMDFe, "", False)
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & chMDFe & "-procMDFe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadMDFeESalvar = resposta
End Function

'Download do Evento do MDF-e
Public Function downloadEventoMDFe(chMDFe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String) As String
    Dim json As String
    Dim urlEnvio As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chMDFe"":""" & chMDFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpEvento"":""" & tpEvento & ""","
    json = json & """nSeqEvento"":""" & nSeqEvento & """"
    json = json & "}"
    
    gravaLinhaLog ("[DOWNLOAD_EVENTO_DADOS]")
    gravaLinhaLog (json)

    urlEnvio = "https://mdfe.ns.eti.br/mdfe/get/event"

    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
    
    gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
    gravaLinhaLog (resposta)

    downloadEventoMDFe = resposta
End Function

'Download do Evento do MDF-e e Salvar
Public Function downloadEventoMDFeESalvar(chMDFe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String, caminho As String, exibeNaTela As Boolean) As String
    Dim baixarXML As Boolean
    Dim baixarPDF As Boolean
    Dim baixarJSON As Boolean
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadEventoMDFe(chMDFe, tpAmb, tpDown, tpEvento, nSeqEvento)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
    
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        'Checa se deve baixar XML
        If InStr(1, tpDown, "X") Then
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chMDFe, nSeqEvento, True)
        End If
        'Checa se deve baixar JSON
        If InStr(1, tpDown, "J") Then
            json = LerDadosJSON(resposta, "json", "", "")
            Call salvarJSON(json, caminho, chMDFe, nSeqEvento, True)
        End If
        'Checa se deve baixar PDF
        If InStr(1, tpDown, "P") Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chMDFe, nSeqEvento, True)
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & chMDFe & nSeqEvento & "-procEvenMDFe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadEventoMDFeESalvar = resposta
End Function

'Realizar a Carta de Correção do MDF-e
Public Function encerrarMDFe(chMDFe As String, nProt As String, dhEvento As String, dtEnc As String, cUF As String, cMun As String, tpAmb As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim urlEnvio As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    json = "{"
    json = json & """chMDFe"":""" & chMDFe & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """dtEnc"":""" & dtEnc & ""","
    json = json & """cUF"":""" & cUF & ""","
    json = json & """cMun"":""" & cMun & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"
    
    gravaLinhaLog ("[ENCERRAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    urlEnvio = "https://mdfe.ns.eti.br/mdfe/closure"
    
    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
        
    gravaLinhaLog ("[ENCERRAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
    
        respostaDownload = downloadEventoMDFeESalvar(chMDFe, tpAmb, tpDown, "ENC", "1", caminho, exibeNaTela)
        status = LerDadosJSON(resposta, "status", "", "")
        
        If (status <> "200") Then
            resposta = resposta + Chr(13) + Chr(13) + respostaDownload
        End If
    End If
    
    encerrarMDFe = resposta
End Function

'Realizar o cancelamento do MDF-e
Public Function cancelarMDFe(chMDFe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    json = "{"
    json = json & """chMDFe"":""" & chMDFe & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """xJust"":""" & xJust & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    url = "https://mdfe.ns.eti.br/mdfe/cancel"
    
    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
        
    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
    
        respostaDownload = downloadEventoMDFeESalvar(chMDFe, tpAmb, tpDown, "CANC", "1", caminho, exibeNaTela)
        status = LerDadosJSON(resposta, "status", "", "")
        
        If (status <> "200") Then
            resposta = resposta + Chr(13) + Chr(13) + respostaDownload
        End If
    End If
    
    cancelarMDFe = resposta
End Function

'Incluir um novo condutor no MDF-e
Public Function incluirCondutorMDFe(chMDFe As String, tpDown As String, tpAmb As String, dhEvento As String, xNome As String, CPF As String, nSeqEvento As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    json = "{"
    json = json & """chMDFe"":""" & chMDFe & ""","
    json = json & """xNome"":""" & xNome & ""","
    json = json & """CPF"":""" & CPF & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"
    
    gravaLinhaLog ("[INCLUSAO_CONDUTOR_DADOS]")
    gravaLinhaLog (json)
    
    url = "https://mdfe.ns.eti.br/mdfe/adddriver"
    
    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
        
    gravaLinhaLog ("[INCLUSAO_CONDUTOR_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
    
        respostaDownload = downloadEventoMDFeESalvar(chMDFe, tpAmb, tpDown, "INCCOND", nSeqEvento, caminho, exibeNaTela)
        status = LerDadosJSON(resposta, "status", "", "")
        
        If (status <> "200") Then
            resposta = resposta + Chr(13) + Chr(13) + respostaDownload
        End If
    End If
    
    incluirCondutorMDFe = resposta
End Function

'Realiza a conutla de MDF-e que não foram encerrados
Public Function consultarMDFeNaoEcerrada(tpAmb As String, cUF As String, CNPJ As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
   
    json = "{"
    json = json & """CNPJ"":""" & CNPJ & ""","
    json = json & """cUF"":""" & cUF & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"
    
    gravaLinhaLog ("[CONSULTA_MDFE_NAO_ENCERRADA_DADOS]")
    gravaLinhaLog (json)
    
    url = "https://mdfe.ns.eti.br/util/consnotclosed"
    
    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
        
    gravaLinhaLog ("[CONSULTA_MDFE_NAO_ENCERRADA_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarMDFeNaoEcerrada = resposta
End Function

'Lista todos os NSNRecs de um MDF-e
Public Function listarNSNRecs(chMDFe As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    
    json = "{"
    json = json & """chMDFe"":""" & chMDFe & """"
    json = json & "}"
    
    gravaLinhaLog ("[LISTA_NSNRECS_DADOS]")
    gravaLinhaLog (json)
    
    url = "https://mdfe.ns.eti.br/util/list/nsnrecs"
    
    resposta = enviaConteudoParaAPI(json, urlEnvio, "json")
        
    gravaLinhaLog ("[LISTA_NSNRECS_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    listarNSNRecs = resposta
End Function

'Salvar XML
Public Sub salvarXML(xml As String, caminho As String, chMDFe As String, nSeqEvento As String, evento As Boolean)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo XML
    If (evento = False) Then
        localParaSalvar = caminho & chMDFe & nSeqEvento & "-procMDFe.xml"
    Else
        localParaSalvar = caminho & chMDFe & nSeqEvento & "-procEvenMDFe.xml"
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
Public Sub salvarJSON(json As String, caminho As String, chMDFe As String, nSeqEvento As String, evento As Boolean)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo JSON
    If (evento = False) Then
        localParaSalvar = caminho & chMDFe & nSeqEvento & "-procMDFe.json"
    Else
        localParaSalvar = caminho & chMDFe & nSeqEvento & "-procEvenMDFe.json"
    End If

    conteudoSalvar = json

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Salvar PDF
Public Function salvarPDF(pdf As String, caminho As String, chMDFe As String, nSeqEvento As String, evento As Boolean) As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo PDF
    If (evento = False) Then
        localParaSalvar = caminho & chMDFe & nSeqEvento & "-procMDFe.pdf"
    Else
        localParaSalvar = caminho & chMDFe & nSeqEvento & "-procEvenMDFe.pdf"
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
    
    Set xml = New DOMDocument
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
        MsgBox "Não foi possível ler o conteúdo do XML da MDFe especificado para leitura.", vbCritical, "ERRO"
    End If
End Function

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

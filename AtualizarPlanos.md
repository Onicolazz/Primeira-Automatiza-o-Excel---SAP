Sub AtualizarPlanosSAP()

    ' ==========================================
    ' DECLAÇÃO DE VARIÁVEIS e CONFIGURAÇÃO INICIAL
    ' ==========================================
    Dim SAP As Object
    Dim session As Object
    Dim ws As Worksheet
    Dim grid As Object
    Dim i As Long
    Dim countOrdens As Long
    Dim resposta As VbMsgBoxResult
    Dim LinhasQnt As Long
    
    ' Define a planilha "PLANOS" dentro do arquivo atual para manipulação dos dados
    Set ws = ThisWorkbook.Sheets("PLANOS") 
    
    ' Conexão com o sistema SAP GUI ativo e abertura do motor de Scripting
    Set SAP = GetObject("SAPGUI")
    Set session = SAP.GetScriptingEngine.Children(0).Children(0)

    ' Solicita ao usuário a quantidade limite de ordens para cada pausa de controle
    LinhasQnt = InputBox("Digite a quantidade de ordens:")
    
    countOrdens = 0
    
    ' ==========================================
    ' LOOP PRINCIPAL: PROCESSAMENTO DE PLANOS
    ' ==========================================
    ' Varre a planilha da linha 3 até a última linha preenchida na Coluna A
    For i = 3 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        Dim plano As String
        Dim dt1 As String, dt2 As String, dt3 As String
        
        ' Captura o código do plano na Coluna A
        plano = ws.Cells(i, 1).Value 
        
        ' Abre a transação IP24 no SAP (Exibição de Itens de Manutenção)
        session.StartTransaction "IP24"
        
        ' Preenche o campo do plano no SAP e executa a busca (botão Executar / F8)
        session.findById("wnd[0]/usr/ctxtWARPL-LOW").Text = plano
        session.findById("wnd[0]/tbar[1]/btn[8]").Press
        
        ' ==========================================
        ' VALIDAÇÃO DA GRADE DE DADOS DO SAP
        ' ==========================================
        ' Captura o componente de tabela/grade (Grid Control) gerado pelo SAP
        On Error Resume Next
        Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        On Error GoTo 0
        
        ' Valida se o plano existe, se trouxe registros e se possui o mínimo de 3 linhas necessárias
        If grid Is Nothing Then GoTo ProximaLinha
        If grid.RowCount = 0 Then GoTo ProximaLinha
        If grid.RowCount < 3 Then GoTo ProximaLinha
        
        ' ==========================================
        ' EXTRAÇÃO DE DADOS E GRAVAÇÃO NA PLANILHA
        ' ==========================================
        ' Extrai as datas planejadas das 3 primeiras linhas da tabela SAP (coluna ADDAT)
        dt1 = grid.GetCellValue(0, "ADDAT")
        dt2 = grid.GetCellValue(1, "ADDAT")
        dt3 = grid.GetCellValue(2, "ADDAT")
        
        ' Escreve as datas capturadas nas Colunas E, F e G da linha atual
        ws.Cells(i, 5).Value = dt3
        ws.Cells(i, 6).Value = dt2
        ws.Cells(i, 7).Value = dt1
        
        ' Retorna à tela anterior do SAP (atalho F3) para liberar a transação
        session.findById("wnd[0]").sendVKey 3
        
        ' Incrementa o contador de ordens processadas com sucesso
        countOrdens = countOrdens + 1
        
        ' ==========================================
        ' CONTROLE DE EXECUÇÃO (INTERRUPÇÃO DO USUÁRIO)
        ' ==========================================
        ' Verifica se atingiu o bloco de ordens definido no InputBox
        If countOrdens Mod LinhasQnt = 0 Then
            resposta = MsgBox("Foram processadas " & countOrdens & " ordens." & vbCrLf & _
                              "Deseja continuar?", vbYesNo + vbQuestion, "Controle de Execução")
            
            ' Se o usuário escolher "Não", formata os dados já extraídos e encerra
            If resposta = vbNo Then
                ' Trata o padrão de data do SAP substituindo pontos por barras nas colunas E, F e G
                ws.Columns("E:G").Replace _
                What:=".", _
                Replacement:="/", _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False

                ' Aplica a formatação regional de data nas colunas
                ws.Columns("E:G").NumberFormat = "dd/mm/yyyy"
                
                ' Fecha a sessão atual do SAP por segurança e encerra a execução
                session.findById("wnd[0]").CloseSession
                Exit Sub
            End If
        End If

    ProximaLinha:
        ' Limpa o objeto da grade para evitar erros de cache no próximo loop
        Set grid = Nothing
        
    Next i
    
    ' ==========================================
    ' TRATAMENTO FINAL DOS DADOS (FIM DO LOOP)
    ' ==========================================
    ' Substitui os pontos por barras em todas as datas extraídas no bloco final
    ws.Columns("E:G").Replace _
    What:=".", _
    Replacement:="/", _
    LookAt:=xlPart, _
    SearchOrder:=xlByRows, _
    MatchCase:=False

    ' Formata visualmente as colunas E, F e G como Data curta
    ws.Columns("E:G").NumberFormat = "dd/mm/yyyy"

    ' Avisa ao usuário que toda a lista de planos foi perfeitamente atualizada
    MsgBox "Atualização concluída!"

End Sub

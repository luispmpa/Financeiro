#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
;  SIGEF - Salvar e-mail do Outlook como PDF nas etapas do fluxo
;  Atalho: Ctrl + Alt + S (com Outlook em foco)
;  Pasta base: \\pmpa-fs3\smf_sti$\...\Numeração das Demandas\<NUM>\<ETAPA>.pdf
;
;  Etapas suportadas:
;    1_Abertura  → "Demanda Cadastrada" (cria pasta se não existir)
;    2_Aprovação → "Aprovação Demanda Cadastrada" (pasta deve existir)
;    3, 4, 5 → reservadas (adicionar quando padrão for definido)
;
;  Correção definitiva (v2):
;    - Não usa mais espera fixa nem digitação "às cegas".
;    - Espera ATIVAMENTE o diálogo de salvar abrir.
;    - Escreve o caminho DIRETO no controle (ControlSetText), o que
;      não depende de foco, timing ou da janela estar ativa.
;    - Clica no botão "Salvar" pelo controle.
;    - Garante que a impressora padrão seja "Microsoft Print to PDF"
;      durante a impressão e restaura a anterior em seguida.
; ============================================================

PASTA_BASE := "\\pmpa-fs3\smf_sti$\Pasta publica GTI\SIGEF 2021\Implantação SIGEF\Demandas SIGEF\Demandas Novo Contrato\Numeração das Demandas"
ATALHO_SALVAR := "^!s"
IMPRESSORA_PDF := "Microsoft Print to PDF"

; Tempo máximo (segundos) para o diálogo de salvar aparecer
TIMEOUT_DIALOGO := 45
; Tempo máximo (segundos) para o PDF aparecer no disco (rede pode ser lenta)
TIMEOUT_ARQUIVO := 90

Hotkey(ATALHO_SALVAR, SalvarDemandaPDF)

; ============================================================
;  Catálogo de etapas (ordem = prioridade de match)
;  Etapa 1 fica POR ÚLTIMO porque "Demanda Cadastrada" também
;  aparece dentro de "Aprovação Demanda Cadastrada".
; ============================================================
GetEtapas() {
    etapas := []

    ; --- Etapa 2: Aprovação ---
    etapa2 := Map()
    etapa2["nome"] := "2_Aprovação"
    etapa2["criarPasta"] := false
    etapa2["padroes"] := [
        "i)Aprova[çc][aã]o\s+Demanda\s+Cadastrada\s*:?\s*(\d+)",
        "i)Demanda\s+(\d+)\s+aprovada"
    ]
    etapas.Push(etapa2)

    ; --- Etapa 1: Abertura ---
    etapa1 := Map()
    etapa1["nome"] := "1_Abertura"
    etapa1["criarPasta"] := true
    etapa1["padroes"] := [
        "i)Demanda\s*Cadastrada\s*:?\s*(\d+)",
        "i)Demanda\s*:\s*SIGEF\/POA\s*[-–]\s*(\d+)",
        "i)SIGEF\/POA\s*[-–]\s*(\d+)"
    ]
    etapas.Push(etapa1)

    ; Futuro: adicionar etapas 3, 4, 5 aqui (ANTES da etapa 1)

    return etapas
}

; ============================================================
;  Função principal
; ============================================================
SalvarDemandaPDF(*) {
    global PASTA_BASE

    if !WinActive("ahk_exe OUTLOOK.EXE") {
        MsgBox("O Outlook precisa estar em foco.", "SIGEF - Aviso", "Iconx")
        return
    }

    dadosEmail := LerEmailOutlook()
    if !IsObject(dadosEmail) {
        MsgBox("Não foi possível ler o e-mail selecionado no Outlook.",
               "SIGEF - Erro", "Iconx")
        return
    }

    resultado := IdentificarEtapa(dadosEmail["assunto"], dadosEmail["corpo"])
    if !IsObject(resultado) {
        MsgBox("Não foi possível identificar a etapa do fluxo neste e-mail.`n`n"
             . "Etapas suportadas atualmente:`n"
             . '  • 1_Abertura → "Demanda Cadastrada"`n'
             . '  • 2_Aprovação → "Aprovação Demanda Cadastrada"',
               "SIGEF - Erro", "Iconx")
        return
    }

    numeroDemanda := resultado["numero"]
    etapa := resultado["etapa"]
    nomeEtapa := etapa["nome"]
    criarPasta := etapa["criarPasta"]

    if !DirExist(PASTA_BASE) {
        MsgBox("A pasta base não foi encontrada:`n" PASTA_BASE,
               "SIGEF - Erro", "Iconx")
        return
    }

    pastaDemanda := PASTA_BASE "\" numeroDemanda
    if !DirExist(pastaDemanda) {
        if (criarPasta) {
            try {
                DirCreate(pastaDemanda)
            } catch as err {
                MsgBox("Falha ao criar a pasta:`n" pastaDemanda "`n`n" err.Message,
                       "SIGEF - Erro", "Iconx")
                return
            }
        } else {
            MsgBox("A pasta da demanda " numeroDemanda " não existe:`n" pastaDemanda
                 . "`n`nEsta etapa (" nomeEtapa ") não cria pasta automaticamente.`n"
                 . "Salve a etapa 1_Abertura desta demanda primeiro.",
                   "SIGEF - Erro", "Iconx")
            return
        }
    }

    caminhoCompleto := pastaDemanda "\" nomeEtapa ".pdf"
    if FileExist(caminhoCompleto) {
        ToolTip("Demanda " numeroDemanda ": " nomeEtapa ".pdf já existe — ignorado.")
        SetTimer(() => ToolTip(), -2500)
        return
    }

    if !ImprimirComoPDF(caminhoCompleto) {
        return
    }

    ToolTip("✅ Demanda " numeroDemanda " — " nomeEtapa ".pdf salvo.")
    SetTimer(() => ToolTip(), -3500)
}

; ============================================================
;  Lê assunto e corpo do e-mail selecionado via COM
; ============================================================
LerEmailOutlook() {
    try {
        outlook := ComObjActive("Outlook.Application")
        explorer := outlook.ActiveExplorer
        if !IsObject(explorer)
            return ""
        selecao := explorer.Selection
        if (selecao.Count = 0)
            return ""

        item := selecao.Item(1)
        assunto := ""
        corpo := ""
        try assunto := item.Subject
        try corpo := item.Body

        dados := Map()
        dados["assunto"] := assunto
        dados["corpo"] := corpo
        return dados
    } catch as err {
        MsgBox("Erro ao acessar o Outlook via COM:`n" err.Message,
               "SIGEF - Erro", "Iconx")
        return ""
    }
}

; ============================================================
;  Identifica a etapa e extrai o número da demanda
; ============================================================
IdentificarEtapa(assunto, corpo) {
    textoCombinado := assunto . "`n" . corpo
    etapas := GetEtapas()

    for etapa in etapas {
        for padrao in etapa["padroes"] {
            if RegExMatch(textoCombinado, padrao, &m) {
                resultado := Map()
                resultado["etapa"] := etapa
                resultado["numero"] := m[1]
                return resultado
            }
        }
    }
    return ""
}

; ============================================================
;  Imprime o e-mail selecionado como PDF via COM do Outlook
;
;  Estratégia DEFINITIVA (não depende de timing nem digitação):
;   1) Garante "Microsoft Print to PDF" como impressora padrão
;   2) Dispara PrintOut() via COM
;   3) Espera ATIVAMENTE o diálogo de salvar aparecer
;   4) Restaura a impressora padrão original
;   5) Escreve o caminho DIRETO no controle Edit (ControlSetText)
;   6) Clica no botão "Salvar" pelo controle
;   7) Aguarda o arquivo aparecer no disco
; ============================================================
ImprimirComoPDF(caminhoCompleto) {
    global IMPRESSORA_PDF, TIMEOUT_DIALOGO, TIMEOUT_ARQUIVO

    ; Fecha qualquer Backstage/painel aberto antes de começar
    Loop 3 {
        Send("{Esc}")
        Sleep(150)
    }

    ; --- Garante a impressora correta (para o diálogo de salvar aparecer) ---
    impressoraOriginal := ObterImpressoraPadrao()
    trocouImpressora := false
    if (impressoraOriginal != "" && impressoraOriginal != IMPRESSORA_PDF) {
        if DefinirImpressoraPadrao(IMPRESSORA_PDF)
            trocouImpressora := true
    }

    ; --- Dispara impressão direto via COM (ignora a interface) ---
    try {
        outlook := ComObjActive("Outlook.Application")
        explorer := outlook.ActiveExplorer
        selecao := explorer.Selection
        if (selecao.Count = 0) {
            if (trocouImpressora)
                DefinirImpressoraPadrao(impressoraOriginal)
            MsgBox("Nenhum e-mail selecionado para imprimir.",
                   "SIGEF - Erro", "Iconx")
            return false
        }
        item := selecao.Item(1)
        item.PrintOut()
    } catch as err {
        if (trocouImpressora)
            DefinirImpressoraPadrao(impressoraOriginal)
        MsgBox("Falha ao disparar impressão via COM:`n" err.Message,
               "SIGEF - Erro", "Iconx")
        return false
    }

    ToolTip("Aguardando o diálogo 'Salvar Saída de Impressão' abrir...")

    ; --- Espera ATIVA: aguarda o diálogo de salvar realmente aparecer ---
    dialogoHwnd := EsperarDialogoSalvar(TIMEOUT_DIALOGO)

    ; Já pode restaurar a impressora: o driver de PDF já foi escolhido
    if (trocouImpressora)
        DefinirImpressoraPadrao(impressoraOriginal)

    if (!dialogoHwnd) {
        ToolTip()
        MsgBox("O diálogo de salvar PDF não abriu em " TIMEOUT_DIALOGO "s.`n`n"
             . "Verifique se a impressora '" IMPRESSORA_PDF "' está"
             . " instalada no Windows.",
               "SIGEF - Erro", "Iconx")
        return false
    }

    ; Garante que o diálogo está ativo
    try WinActivate("ahk_id " dialogoHwnd)
    WinWaitActive("ahk_id " dialogoHwnd, , 5)
    Sleep(300)

    ToolTip("Preenchendo o nome do arquivo...")

    ; --- Escreve o caminho DIRETO no controle (não depende de digitação) ---
    if !PreencherNomeArquivo(dialogoHwnd, caminhoCompleto) {
        ToolTip()
        MsgBox("Não foi possível preencher o caminho no diálogo de salvar.`n`n"
             . "Caminho:`n" caminhoCompleto,
               "SIGEF - Erro", "Iconx")
        return false
    }

    ToolTip("Salvando...")

    ; --- Clica no botão Salvar pelo controle ---
    if !ClicarSalvar(dialogoHwnd) {
        ToolTip()
        MsgBox("Não foi possível acionar o botão Salvar.",
               "SIGEF - Erro", "Iconx")
        return false
    }

    ; Espera o diálogo de salvar fechar
    WinWaitClose("ahk_id " dialogoHwnd, , 15)

    ; Confirma sobrescrita, caso apareça (não deveria, pois já checamos antes)
    if WinWait("Confirmar Salvar como ahk_class #32770", , 2) {
        Send("{Enter}")
    }

    ToolTip("Aguardando o arquivo no disco...")

    ; --- Aguarda o arquivo aparecer (rede pode ser lenta) ---
    inicio := A_TickCount
    Loop {
        Sleep(500)
        if FileExist(caminhoCompleto) {
            ToolTip()
            return true
        }
        if (A_TickCount - inicio > TIMEOUT_ARQUIVO * 1000) {
            ToolTip()
            MsgBox("O PDF não foi gerado em " TIMEOUT_ARQUIVO "s.`n`n"
                 . "Verifique manualmente:`n" caminhoCompleto,
                   "SIGEF - Aviso", "Icon!")
            return false
        }
    }
}

; ============================================================
;  Espera ativamente o diálogo "Salvar Saída de Impressão como"
;  Retorna o HWND da janela, ou 0 se estourar o timeout.
; ============================================================
EsperarDialogoSalvar(timeoutSeg) {
    inicio := A_TickCount
    Loop {
        ; Diálogos comuns do Windows usam a classe #32770
        try {
            for hwnd in WinGetList("ahk_class #32770") {
                titulo := ""
                try titulo := WinGetTitle("ahk_id " hwnd)
                ; Título pt-BR: "Salvar Saída de Impressão como"
                ; Título en-US: "Save Print Output As"
                if RegExMatch(titulo, "i)(Impress|Print)") {
                    ; Confirma que é o diálogo de salvar (tem campo Edit)
                    if TemControleEdit(hwnd)
                        return hwnd
                }
            }
        }
        if (A_TickCount - inicio > timeoutSeg * 1000)
            return 0
        Sleep(250)
    }
}

; Verifica se a janela possui pelo menos um controle Edit
TemControleEdit(hwnd) {
    try {
        for ctrl in WinGetControls("ahk_id " hwnd) {
            if (SubStr(ctrl, 1, 4) = "Edit")
                return true
        }
    }
    return false
}

; ============================================================
;  Preenche o campo "Nome" do diálogo escrevendo direto no
;  controle. Não depende de foco, timing ou digitação.
; ============================================================
PreencherNomeArquivo(hwnd, caminho) {
    ; O campo "Nome" recebe foco automaticamente quando o diálogo abre
    editAlvo := ""
    try editAlvo := ControlGetFocus("ahk_id " hwnd)

    ; Se o foco não estiver em um Edit, procura o primeiro Edit do diálogo
    if (editAlvo = "" || SubStr(editAlvo, 1, 4) != "Edit") {
        try {
            for ctrl in WinGetControls("ahk_id " hwnd) {
                if (SubStr(ctrl, 1, 4) = "Edit") {
                    editAlvo := ctrl
                    break
                }
            }
        }
    }
    if (editAlvo = "" || SubStr(editAlvo, 1, 4) != "Edit")
        return false

    ; Método principal: escreve direto no controle
    try {
        ControlFocus(editAlvo, "ahk_id " hwnd)
        ControlSetText(caminho, editAlvo, "ahk_id " hwnd)
        Sleep(250)
        if (ControlGetText(editAlvo, "ahk_id " hwnd) = caminho)
            return true
    }

    ; Fallback: cola via clipboard no controle
    try {
        clipAntigo := ClipboardAll()
        A_Clipboard := caminho
        if ClipWait(2) {
            ControlFocus(editAlvo, "ahk_id " hwnd)
            Sleep(150)
            ControlSend("^a", editAlvo, "ahk_id " hwnd)
            Sleep(120)
            ControlSend("^v", editAlvo, "ahk_id " hwnd)
            Sleep(400)
            SetTimer(() => (A_Clipboard := clipAntigo), -2000)
            if (ControlGetText(editAlvo, "ahk_id " hwnd) != "")
                return true
        }
    }
    return false
}

; ============================================================
;  Clica no botão "Salvar" do diálogo (pelo controle)
; ============================================================
ClicarSalvar(hwnd) {
    botaoSalvar := ""
    try {
        for ctrl in WinGetControls("ahk_id " hwnd) {
            if (SubStr(ctrl, 1, 6) = "Button") {
                txt := ""
                try txt := ControlGetText(ctrl, "ahk_id " hwnd)
                if RegExMatch(txt, "i)Salvar|Save") {
                    botaoSalvar := ctrl
                    break
                }
            }
        }
    }

    if (botaoSalvar != "") {
        try {
            ControlClick(botaoSalvar, "ahk_id " hwnd)
            return true
        }
    }

    ; Fallback: Enter no diálogo ativo (Salvar é o botão padrão)
    try {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 3)
        Send("{Enter}")
        return true
    }
    return false
}

; ============================================================
;  Gerenciamento da impressora padrão (via WMI)
; ============================================================
ObterImpressoraPadrao() {
    try {
        for prn in ComObjGet("winmgmts:").ExecQuery(
            "SELECT Name FROM Win32_Printer WHERE Default=TRUE")
            return prn.Name
    }
    return ""
}

DefinirImpressoraPadrao(nomeImpressora) {
    try {
        nomeEscapado := StrReplace(nomeImpressora, "'", "''")
        for prn in ComObjGet("winmgmts:").ExecQuery(
            "SELECT * FROM Win32_Printer WHERE Name='" nomeEscapado "'") {
            prn.SetDefaultPrinter()
            return true
        }
    }
    return false
}

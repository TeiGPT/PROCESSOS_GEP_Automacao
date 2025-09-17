<#!
processar_processo.ps1
Pipeline completo para tratamento de processos de peritagem GEP.
Execute com PowerShell:  .\processar_processo.ps1
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# === Configuração ===
$Config = [ordered]@{
    BaseOutputRoot = "H:\PROCESSOS_GEP"
    PopplerPath    = "H:\Programas instalados\Ferramentas\Poppler\bin\pdftotext.exe"
    TesseractPath  = "H:\Programas instalados\Ferramentas\Tesseract\tesseract.exe"
    EmailTemplate  = "H:\PROJECTOS\PROCESSOS_GEP_Automacao\templates\email_template.html"
    Locale         = "pt-PT"
    ConfigFile     = "H:\PROJECTOS\PROCESSOS_GEP_Automacao\_config\config.ini"
}

function Read-ConfigFile {
    param([string]$ConfigPath)
    $config = @{}
    if (Test-Path -LiteralPath $ConfigPath) {
        $content = Get-Content -LiteralPath $ConfigPath -Raw
        $currentSection = ""
        foreach ($line in ($content -split "`n")) {
            $line = $line.Trim()
            if ($line -match '^\[(.+)\]$') {
                $currentSection = $matches[1]
                $config[$currentSection] = @{}
            } elseif ($line -match '^(.+?)\s*=\s*(.+)$' -and $currentSection) {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $config[$currentSection][$key] = $value
            }
        }
    }
    return $config
}

$script:LogPath = $null

function As-Text([object]$v) {
  if ($null -eq $v) { return "" }
  if ($v -is [System.Management.Automation.ErrorRecord]) {
    $m = if ($v.Exception) { $v.Exception.Message } else { $v.ToString() }
    return ($m | Out-String).Trim()
  }
  if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
    return ( ($v | Out-String) ).Trim()
  }
  return ( ($v | Out-String) ).Trim()
}

function Initialize-Log {
    param(
        [Parameter(Mandatory)] [string] $LogFile
    )
    $script:LogPath = $LogFile
    if (-not (Test-Path -LiteralPath $LogFile)) {
        New-Item -ItemType File -Path $LogFile -Force | Out-Null
    }
    Write-Log "Log inicializado." "INFO"
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet("INFO","WARN","ERRO","OK")] [string] $Level = "INFO"
    )
    $Message = As-Text $Message
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "{0} [{1}] {2}" -f $timestamp, $Level, $Message
    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line
    }
    Write-Host $line
}

function Show-ErrorMessage {
    param([string]$Message)
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show($Message, "PROCESSOS GEP", 'OK', 'Error') | Out-Null
    } catch {
        Write-Warning $Message
    }
}

function Select-InputPdf {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Ficheiros PDF (*.pdf)|*.pdf"
    $dialog.Multiselect = $false
    $dialog.Title = "Seleccione o ficheiro PDF do processo"
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Operação cancelada pelo utilizador."
    }
    return $dialog.FileName
}

function Get-ProcessIdFromFileName {
    param([string]$FilePath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $match = [regex]::Match($name, "apn[\s_-]*(?<id>[0-9A-Za-z]+)", 'IgnoreCase')
    if ($match.Success) {
        return $match.Groups['id'].Value.ToUpper()
    }
    return $name.ToUpper()
}

function Ensure-ProcessFolders {
    param(
        [string]$ProcessId
    )
    $base = Join-Path $Config.BaseOutputRoot $ProcessId
    $folders = @(
        $base,
        (Join-Path $base "origem"),
        (Join-Path $base "trabalho"),
        (Join-Path $base "output")
    )
    foreach ($folder in $folders) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            Write-Log "Criada pasta: $folder"
        }
    }
    return [ordered]@{
        Base     = $base
        Origem   = Join-Path $base "origem"
        Trabalho = Join-Path $base "trabalho"
        Output   = Join-Path $base "output"
    }
}

function Copy-SourcePdf {
    param(
        [string]$InputPdf,
        [string]$OrigemFolder
    )
    $destPdf = Join-Path $OrigemFolder ([System.IO.Path]::GetFileName($InputPdf))
    Copy-Item -LiteralPath $InputPdf -Destination $destPdf -Force
    Write-Log "PDF copiado para $destPdf"
    return $destPdf
}

function PdfParaTxt_Temp {
    param([string]$PdfPath)
    try {
        if (-not (Test-Path -LiteralPath $Config.PopplerPath)) {
            Write-Log "pdftotext não encontrado em $($Config.PopplerPath)" "ERRO"
            return ""
        }
        $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + ".txt")
        $arguments = "-layout -nopgbrk -enc UTF-8 `"$PdfPath`" `"$tempPath`""
        Write-Log "A executar pdftotext: $arguments"
        $proc = Start-Process -FilePath $Config.PopplerPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru
        if ($proc.ExitCode -ne 0) {
            Write-Log "pdftotext terminou com código $($proc.ExitCode)" "WARN"
            if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
            return ""
        }
        if (-not (Test-Path -LiteralPath $tempPath)) {
            Write-Log "pdftotext não gerou ficheiro." "WARN"
            return ""
        }
        return $tempPath
    } catch {
        Write-Log "Falha em PdfParaTxt_Temp: $(As-Text $_)" "ERRO"
        return ""
    }
}

function ConvertWithTesseract {
    param(
        [string]$PdfPath,
        [string]$TargetTxt
    )
    if (-not (Test-Path -LiteralPath $Config.TesseractPath)) {
        Write-Log "tesseract.exe não encontrado em $($Config.TesseractPath)" "ERRO"
        return $false
    }
    $dir = Split-Path -Path $TargetTxt -Parent
    $base = [System.IO.Path]::GetFileNameWithoutExtension($TargetTxt)
    $intermediateBase = Join-Path $dir ($base + "_ocr")
    $arguments = "`"$PdfPath`" `"$intermediateBase`" -l por --psm 3"
    Write-Log "A executar Tesseract: $arguments"
    $proc = Start-Process -FilePath $Config.TesseractPath -ArgumentList $arguments -NoNewWindow -Wait -Pass-Thru
    if ($proc.ExitCode -ne 0) {
        Write-Log "Tesseract terminou com código $($proc.ExitCode)" "ERRO"
        return $false
    }
    $generated = $intermediateBase + ".txt"
    if (-not (Test-Path -LiteralPath $generated)) {
        Write-Log "Tesseract não gerou ficheiro de texto." "ERRO"
        return $false
    }
    Move-Item -LiteralPath $generated -Destination $TargetTxt -Force
    return $true
}

function Ensure-DocTxt {
    param(
        [string]$SourcePdf,
        [string]$WorkFolder
    )
    $docTxtPath = Join-Path $WorkFolder "doc.txt"
    $popplerTxt = PdfParaTxt_Temp -PdfPath $SourcePdf
    if (Test-Path -LiteralPath $popplerTxt) {
        Move-Item -LiteralPath $popplerTxt -Destination $docTxtPath -Force
        Write-Log "Texto gerado com pdftotext."
        return $docTxtPath
    }
    Write-Log "A tentar OCR com Tesseract." "WARN"
    if (ConvertWithTesseract -PdfPath $SourcePdf -TargetTxt $docTxtPath) {
        Write-Log "Texto gerado via Tesseract OCR."
        return $docTxtPath
    }
    throw "Falha na conversão PDF→TXT com pdftotext e Tesseract."
}

function Read-TextContent {
    param([string]$TextPath)
    return [System.IO.File]::ReadAllText($TextPath, [System.Text.Encoding]::UTF8)
}

function Extract-Fields {
    param([string]$Text)
    $fields = [ordered]@{}

    $fields['nome'] = (Select-RegexFirst $Text "Nome[:\s]*([^\r\n]+)")
    $fields['morada'] = (Select-RegexFirst $Text "Morada[:\s]*([^\r\n]+)")
    $fields['localidade'] = (Select-RegexFirst $Text "Localidade[:\s]*([^\r\n]+)")
    $fields['codigo_postal'] = Validate-OrEmpty (Select-RegexFirst $Text "C[oó]digo\s+Postal[:\s]*([0-9]{4}-[0-9]{3})") -Pattern "^[0-9]{4}-[0-9]{3}$"
    $fields['telefone'] = Validate-OrEmpty (Select-RegexFirst $Text "Telefone[:\s]*([0-9 \+/()-]{6,})") -Pattern "[0-9]{6,}"
    $fields['email'] = Validate-OrEmpty (Select-RegexFirst $Text "Email[:\s]*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})" 'IgnoreCase') -Pattern "^[^@\s]+@[^@\s]+\.[^@\s]+$"
    $fields['nif'] = Validate-OrEmpty (Select-RegexFirst $Text "N[º°\.\s]*de\s+Contribuinte\s*/?\s*NIPC[:\s]*([0-9]{9})") -Pattern "^[0-9]{9}$"
    $fields['num_apolice'] = (Select-RegexFirst $Text "N[º°\.\s]*de\s+Ap[oó]lice[:\s]*([^\r\n]+)")

    $datas = [regex]::Matches($Text, "\b[0-3]?\d/[0-1]?\d/\d{4}\b") | ForEach-Object { $_.Value }
    $fields['datas'] = $datas | Select-Object -Unique
    return $fields
}

function Select-RegexFirst {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Options = "IgnoreCase"
    )
    $regex = New-Object System.Text.RegularExpressions.Regex($Pattern, $Options)
    $match = $regex.Match($Text)
    if ($match.Success) {
        if ($match.Groups.Count > 1) {
            return $match.Groups[1].Value.Trim()
        }
        return $match.Value.Trim()
    }
    return ""
}

function Validate-OrEmpty {
    param(
        [string]$Value,
        [string]$Pattern
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    if ($Value -match $Pattern) {
        return $Value.Trim()
    }
    return ""
}

function Write-DataJson {
    param(
        [string]$JsonPath,
        [hashtable]$Fields,
        [string]$ProcessId,
        [string]$SourcePdf,
        [string]$CopiedPdf
    )
    $payload = [ordered]@{
        process_id = $ProcessId
        timestamp = (Get-Date).ToString("s")
        source_pdf = $SourcePdf
        copied_pdf = $CopiedPdf
        fields = $Fields
    }
    $json = $payload | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $JsonPath -Value $json -Encoding UTF8
    Write-Log "data.json criado em $JsonPath"
}

function Read-WordSections {
    param([string]$DocxPath)
    $sections = [ordered]@{
        Descricao = ""
        Causas    = ""
        Conclusoes= ""
    }
    if (-not (Test-Path -LiteralPath $DocxPath)) {
        Write-Log "DOCX de descrição não encontrado (opcional)." "WARN"
        return $sections
    }
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($DocxPath, $false, $true)
        $text = $doc.Content.Text
        $sections['Descricao']  = Get-SectionText $text "Descri[çc][aã]o" "Causas"
        $sections['Causas']     = Get-SectionText $text "Causas" "Conclus[õo]es"
        $sections['Conclusoes'] = Get-SectionText $text "Conclus[õo]es" $null
    } catch {
        Write-Log "Erro a ler DOCX: $(As-Text $_)" "WARN"
    } finally {
        if ($doc) { $doc.Close([ref]0) }
        if ($word) { $word.Quit() }
        if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    }
    return $sections
}

function Get-SectionText {
    param(
        [string]$FullText,
        [string]$HeaderPattern,
        [string]$NextHeaderPattern
    )
    $regexPattern = if ($NextHeaderPattern) {
        "(?is)$HeaderPattern\s*[:\-]*\s*(.*?)\s*(?=$NextHeaderPattern\s*[:\-]|\Z)"
    } else {
        "(?is)$HeaderPattern\s*[:\-]*\s*(.*)"
    }
    $match = [regex]::Match($FullText, $regexPattern)
    if ($match.Success) {
        return ($match.Groups[1].Value.Trim())
    }
    return ""
}

function Generate-ReportPdf {
    param(
        [string]$OutputPdf,
        [hashtable]$Fields,
        [hashtable]$Sections,
        [string]$ProcessId
    )
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Add()
        $selection = $word.Selection

        $selection.Style = "Title"
        $selection.TypeText("Relatório de Processo GEP " + $ProcessId)
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Style = "Heading 1"
        $selection.TypeText("Dados do Segurado")
        $selection.TypeParagraph()

        $selection.Style = "Normal"
        $selection.TypeText("Nome: " + (if ($Fields['nome']) { $Fields['nome'] } else { "(não encontrado)" }))
        $selection.TypeParagraph()
        $selection.TypeText("Morada: " + (if ($Fields['morada']) { $Fields['morada'] } else { "" }))
        $selection.TypeParagraph()
        $selection.TypeText("Localidade: " + (if ($Fields['localidade']) { $Fields['localidade'] } else { "" }))
        $selection.TypeParagraph()
        $selection.TypeText("Código Postal: " + (if ($Fields['codigo_postal']) { $Fields['codigo_postal'] } else { "" }))
        $selection.TypeParagraph()
        $selection.TypeText("Telefone: " + (if ($Fields['telefone']) { $Fields['telefone'] } else { "" }))
        $selection.TypeParagraph()
        $selection.TypeText("Email: " + (if ($Fields['email']) { $Fields['email'] } else { "" }))
        $selection.TypeParagraph()
        $selection.TypeText("NIF/NIPC: " + (if ($Fields['nif']) { $Fields['nif'] } else { "" }))
        $selection.TypeParagraph()
        $selection.TypeText("Nº Apólice: " + (if ($Fields['num_apolice']) { $Fields['num_apolice'] } else { "" }))
        $selection.TypeParagraph()
        if ($Fields['datas']) {
            $selection.TypeText("Datas relevantes: " + ($Fields['datas'] -join ", "))
            $selection.TypeParagraph()
        }
        $selection.TypeParagraph()

        foreach ($sectionName in @("Descricao","Causas","Conclusoes")) {
            $selection.Style = "Heading 1"
            switch ($sectionName) {
                "Descricao"   { $label = "Descrição" }
                "Causas"      { $label = "Causas" }
                "Conclusoes"  { $label = "Conclusões" }
            }
            $selection.TypeText($label)
            $selection.TypeParagraph()
            $selection.Style = "Normal"
            $selection.TypeText(($Sections[$sectionName] -replace '\s+$','').Trim())
            $selection.TypeParagraph()
            $selection.TypeParagraph()
        }

        $doc.ExportAsFixedFormat($OutputPdf, 17) # 17 = wdExportFormatPDF
        Write-Log "relatorio_fillform.pdf criado."
    } catch {
        throw "Erro ao gerar relatório PDF: $(As-Text $_)"
    } finally {
        if ($doc) { $doc.Close([ref]0) }
        if ($word) { $word.Quit() }
        if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    }
}

function Create-OutlookDraft {
    param(
        [hashtable]$Fields,
        [string]$ProcessId,
        [string]$TemplatePath
    )
    if (-not $Fields['email']) {
        Write-Log "Email do segurado ausente. Rascunho não criado." "WARN"
        return
    }
    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        Write-Log "Template de email não encontrado em $TemplatePath." "WARN"
        return
    }
    $html = Get-Content -LiteralPath $TemplatePath -Raw
    $placeholders = @{
        "{{NOME}}" = $Fields['nome']
        "{{NUM_APOLICE}}" = $Fields['num_apolice']
        "{{DATA_PROCESSO}}" = (Get-Date).ToString("dd/MM/yyyy")
        "{{NIF}}" = $Fields['nif']
        "{{MORADA}}" = $Fields['morada']
        "{{CODIGO_POSTAL}}" = $Fields['codigo_postal']
        "{{LOCALIDADE}}" = $Fields['localidade']
    }
    foreach ($key in $placeholders.Keys) {
        $html = $html.Replace($key, $placeholders[$key])
    }
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = $Fields['email']
        $mail.Subject = "Processo GEP – " + (if ($Fields['nome']) { $Fields['nome'] } else { $ProcessId }) + " / " + (if ($Fields['num_apolice']) { $Fields['num_apolice'] } else { "(sem apólice)" })
        $mail.HTMLBody = $html
        $mail.Display()
        Write-Log "Rascunho de email criado no Outlook."
    } catch {
        Write-Log "Falha ao criar rascunho no Outlook: $(As-Text $_)" "ERRO"
    }
}

function Create-FillformJson {
    param(
        [string]$JsonPath,
        [hashtable]$Fields,
        [hashtable]$Sections
    )
    
    # Validações e normalizações
    $segurado_cp = ""
    if ($Fields['codigo_postal'] -and $Fields['codigo_postal'] -match '^\d{4}-\d{3}$') {
        $segurado_cp = $Fields['codigo_postal']
    } elseif ($Fields['codigo_postal']) {
        Write-Log "Código postal inválido: $($Fields['codigo_postal'])" "WARN"
    }
    
    $nif_nipc = ""
    if ($Fields['nif'] -and $Fields['nif'] -match '^\d{9}$') {
        $nif_nipc = $Fields['nif']
    } elseif ($Fields['nif']) {
        Write-Log "NIF/NIPC inválido: $($Fields['nif'])" "WARN"
    }
    
    $segurado_email = ""
    if ($Fields['email'] -and $Fields['email'] -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $segurado_email = $Fields['email']
    } elseif ($Fields['email']) {
        Write-Log "Email inválido: $($Fields['email'])" "WARN"
    }
    
    # Extrair primeira data como data_ocorrencia se disponível
    $data_ocorrencia = ""
    if ($Fields['datas'] -and $Fields['datas'].Count -gt 0) {
        $data_ocorrencia = $Fields['datas'][0]
    }
    
    $fillform = [ordered]@{
        segurado_nome = if ($Fields['nome']) { $Fields['nome'] } else { "" }
        segurado_morada = if ($Fields['morada']) { $Fields['morada'] } else { "" }
        segurado_localidade = if ($Fields['localidade']) { $Fields['localidade'] } else { "" }
        segurado_cp = $segurado_cp
        segurado_telefone = if ($Fields['telefone']) { $Fields['telefone'] } else { "" }
        segurado_email = $segurado_email
        nif_nipc = $nif_nipc
        apolice_numero = if ($Fields['num_apolice']) { $Fields['num_apolice'] } else { "" }
        apolice_companhia = "GEP"
        data_ocorrencia = $data_ocorrencia
        data_emissao = (Get-Date).ToString("dd/MM/yyyy")
        descricao = if ($Sections['Descricao']) { $Sections['Descricao'] } else { "" }
        causas = if ($Sections['Causas']) { $Sections['Causas'] } else { "" }
        conclusoes = if ($Sections['Conclusoes']) { $Sections['Conclusoes'] } else { "" }
    }
    
    $json = $fillform | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $JsonPath -Value $json -Encoding UTF8
    Write-Log "fillform.json criado" "INFO"
}

try {
    Write-Log "Início do pipeline PROCESSOS GEP" "INFO"
    
    # Ler configuração de features
    $iniConfig = Read-ConfigFile -ConfigPath $Config.ConfigFile
    $FeaturePdf   = (($iniConfig.features.generate_pdf   -as [string]) -match '^(1|true|on|yes)$')
    $FeatureEmail = (($iniConfig.features.generate_email -as [string]) -match '^(1|true|on|yes)$')
    Write-Log "Config PDF: $($iniConfig.features.generate_pdf) -> FeaturePdf: $FeaturePdf" "INFO"
    Write-Log "Config Email: $($iniConfig.features.generate_email) -> FeatureEmail: $FeatureEmail" "INFO"
    Write-Log "PowerShell: $($PSVersionTable.PSVersion)" "INFO"
    
    # Aviso sobre versão do PowerShell
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Log "A correr em Windows PowerShell 5.1; recomenda-se PowerShell 7." "WARN"
    }
    
    
    # Pré-flight check
    $PdfToTextOK = $false
    $TesseractOK = $false
    
    Write-Log "pdftotext: `"$($Config.PopplerPath)`"" "INFO"
    try {
        $pdfResult = & $Config.PopplerPath -v 2>&1
        $pdfOutput = As-Text $pdfResult
        if ($pdfOutput -imatch "pdftotext.*version" -or $pdfOutput -imatch "version.*pdftotext" -or $pdfOutput -imatch "pdftotext") {
            Write-Log "pdftotext disponível: $($pdfOutput -split "`n" | Select-Object -First 1)" "OK"
            $PdfToTextOK = $true
        } else {
            Write-Log "pdftotext indisponível: resposta inesperada - $pdfOutput" "ERRO"
        }
    } catch {
        Write-Log "pdftotext indisponível: $(As-Text $_)" "ERRO"
    }
    
    Write-Log "tesseract: `"$($Config.TesseractPath)`"" "INFO"
    try {
        $tessResult = & $Config.TesseractPath --version 2>&1
        $tessOutput = As-Text $tessResult
        if ($tessOutput -imatch "tesseract" -and $tessOutput -imatch "version") {
            Write-Log "tesseract disponível: $($tessOutput -split "`n" | Select-Object -First 1)" "OK"
            $TesseractOK = $true
        } else {
            if ($PdfToTextOK) {
                Write-Log "Tesseract indisponível (OCR desactivado)" "WARN"
            } else {
                Write-Log "tesseract indisponível: resposta inesperada" "ERRO"
            }
        }
    } catch {
        if ($PdfToTextOK) {
            Write-Log "Tesseract indisponível (OCR desactivado)" "WARN"
        } else {
            Write-Log "tesseract indisponível: $(As-Text $_)" "ERRO"
        }
    }
    
    # Gate do arranque
    if (-not ($PdfToTextOK -or $TesseractOK)) {
        Write-Log "Nenhuma ferramenta disponível (pdftotext e tesseract falharam)" "ERRO"
        throw "Nenhuma ferramenta de extração disponível"
    }

    $inputPdf = Select-InputPdf
    Write-Log "Ficheiro seleccionado: $inputPdf" "INFO"

    $processId = Get-ProcessIdFromFileName -FilePath $inputPdf
    Write-Log "Processo identificado: $processId" "INFO"

    $folders = Ensure-ProcessFolders -ProcessId $processId
    $logFile = Join-Path $folders.Trabalho "log.txt"
    Initialize-Log -LogFile $logFile

    $copiedPdf = Copy-SourcePdf -InputPdf $inputPdf -OrigemFolder $folders.Origem

    $docTxt = Ensure-DocTxt -SourcePdf $copiedPdf -WorkFolder $folders.Trabalho

    $textContent = Read-TextContent -TextPath $docTxt
    $fields = Extract-Fields -Text $textContent
    $dataJsonPath = Join-Path $folders.Trabalho "data.json"
    Write-DataJson -JsonPath $dataJsonPath -Fields $fields -ProcessId $processId -SourcePdf $inputPdf -CopiedPdf $copiedPdf

    $docxPath = Join-Path $folders.Origem "peritagem_descricao_causas_conclusoes.docx"
    $sections = Read-WordSections -DocxPath $docxPath
    
    # Criar fillform.json sempre
    $fillformJsonPath = Join-Path $folders.Trabalho "fillform.json"
    Create-FillformJson -JsonPath $fillformJsonPath -Fields $fields -Sections $sections
    
    # PDF opcional
    if ($FeaturePdf) {
        $pdfReport = Join-Path $folders.Trabalho "relatorio_fillform.pdf"
        Generate-ReportPdf -OutputPdf $pdfReport -Fields $fields -Sections $sections -ProcessId $processId
    } else {
        Write-Log "PDF desactivado por configuração." "INFO"
    }

    # Email opcional
    if ($FeatureEmail) {
        Create-OutlookDraft -Fields $fields -ProcessId $processId -TemplatePath $Config.EmailTemplate
    } else {
        Write-Log "Email desactivado por configuração." "INFO"
    }

    Write-Log "Processo $processId concluído com sucesso." "OK"

} catch {
    $message = "Erro geral: $(As-Text $_)"
    Write-Log $message "ERRO"
    Show-ErrorMessage $message | Out-Null
    exit 1
}

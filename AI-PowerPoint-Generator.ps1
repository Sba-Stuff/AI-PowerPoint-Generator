#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ---------- Configuration ----------
$apiBase = "http://192.168.10.4:1234/v1"
$model   = "liquid/lfm2-1.2b"

# API endpoint (chat completions)
$chatEndpoint = "$apiBase/chat/completions"

# ---------- GUI Setup ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "AI Presentation Generator"
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = "CenterScreen"

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Presentation Title:"
$lblTitle.Location = New-Object System.Drawing.Point(20,30)
$lblTitle.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($lblTitle)

$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Location = New-Object System.Drawing.Point(150,28)
$txtTitle.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($txtTitle)

$lblAuthor = New-Object System.Windows.Forms.Label
$lblAuthor.Text = "Author Name:"
$lblAuthor.Location = New-Object System.Drawing.Point(20,70)
$lblAuthor.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($lblAuthor)

$txtAuthor = New-Object System.Windows.Forms.TextBox
$txtAuthor.Location = New-Object System.Drawing.Point(150,68)
$txtAuthor.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($txtAuthor)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Presentation"
$btnGenerate.Location = New-Object System.Drawing.Point(150,120)
$btnGenerate.Size = New-Object System.Drawing.Size(150,30)
$btnGenerate.Add_Click({
    $btnGenerate.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Generate-Presentation -Title $txtTitle.Text -Author $txtAuthor.Text
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    } finally {
        $btnGenerate.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$form.Controls.Add($btnGenerate)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# ---------- Helper Functions ----------
function Update-Status {
    param([string]$message)
    $statusLabel.Text = $message
    $statusStrip.Refresh()
    Start-Sleep -Milliseconds 100
}

function Invoke-AI {
    param([string]$prompt, [int]$maxTokens = 600)

    $body = @{
        model    = $model
        messages = @(
            @{ role = "system"; content = "You are a helpful assistant that creates presentation outlines and bullet points. Use clear, concise language." }
            @{ role = "user";    content = $prompt }
        )
        temperature = 0.7
        max_tokens  = $maxTokens
        stream      = $false
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri $chatEndpoint -Method Post -Body $body -ContentType "application/json" -ErrorAction Stop
        return $response.choices[0].message.content.Trim()
    } catch {
        Write-Warning "AI call failed: $_"
        return $null
    }
}

function Generate-Outline {
    param([string]$title)

    Update-Status "Generating outline for '$title'..."
    $prompt = "Generate an outline for a presentation about '$title'. Provide 5 to 7 main sections or topics. List each topic on a new line, without numbers or bullet symbols. Do not include extra commentary."
    $outlineText = Invoke-AI -prompt $prompt -maxTokens 400

    if (-not $outlineText) {
        # Fallback outline if AI fails
        Write-Warning "AI outline generation failed, using fallback."
        $outlineText = "Introduction`nKey Concepts`nBenefits`nChallenges`nCase Studies`nConclusion"
    }

    # Split into lines and clean up
    $topics = $outlineText -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    Write-Host "Generated topics: $($topics -join ', ')"
    return $topics
}

function Generate-Bullets {
    param([string]$title, [string]$topic)

    Update-Status "Generating bullet points for '$topic'..."
    $prompt = "For a presentation titled '$title', write 5 to 7 bullet points for the section '$topic'. Each bullet should be a complete, informative sentence. Start each bullet with a bullet symbol (•) on its own line. No extra text."
    $bulletText = Invoke-AI -prompt $prompt -maxTokens 600

    if (-not $bulletText) {
        Write-Warning "AI bullet generation failed, using placeholder."
        $bulletText = "• Example bullet 1 for $topic`n• Example bullet 2 for $topic`n• Example bullet 3 for $topic"
    }

    # Ensure bullets start with •; if not, add one
    $lines = $bulletText -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $bullets = @()
    foreach ($line in $lines) {
        if ($line -notlike "•*") {
            $line = "• $line"
        }
        $bullets += $line
    }
    return $bullets -join "`r"
}

function New-PowerPoint {
    param([string]$title, [string]$author, [array]$topics, [hashtable]$bulletsPerTopic)

    Update-Status "Creating PowerPoint presentation..."
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    $pres = $ppt.Presentations.Add()
    $slideCount = 1

    # --- Title Slide ---
    $titleSlide = $pres.Slides.Add($slideCount, 1)   # ppLayoutTitle
    $titleSlide.Shapes.Item(1).TextFrame.TextRange.Text = $title
    $titleSlide.Shapes.Item(2).TextFrame.TextRange.Text = $author
    $slideCount++

    # --- Content Slides ---
    foreach ($topic in $topics) {
        $slide = $pres.Slides.Add($slideCount, 2)   # ppLayoutText
        $slide.Shapes.Item(1).TextFrame.TextRange.Text = $topic
        $slide.Shapes.Item(2).TextFrame.TextRange.Text = $bulletsPerTopic[$topic]
        $slideCount++
    }

    # --- Save to Desktop ---
    $desktop = [Environment]::GetFolderPath("Desktop")
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $fileName = "$desktop\$title - $timestamp.pptx"
    $pres.SaveAs($fileName)
    Write-Host "Presentation saved: $fileName" -ForegroundColor Green

    # Optional: Clean up COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    return $fileName
}

function Generate-Presentation {
    param([string]$Title, [string]$Author)

    if ([string]::IsNullOrWhiteSpace($Title)) {
        throw "Title cannot be empty."
    }
    if ([string]::IsNullOrWhiteSpace($Author)) {
        $Author = "Generated by AI"
    }

    # 1. Generate outline topics
    $topics = Generate-Outline -title $Title
    if ($topics.Count -eq 0) {
        throw "No topics generated. Aborting."
    }

    # 2. For each topic, generate bullets
    $bulletsPerTopic = @{}
    foreach ($topic in $topics) {
        $bullets = Generate-Bullets -title $Title -topic $topic
        $bulletsPerTopic[$topic] = $bullets
        Start-Sleep -Milliseconds 200   # small delay to avoid overwhelming the API
    }

    # 3. Build PowerPoint
    $outputFile = New-PowerPoint -title $Title -author $Author -topics $topics -bulletsPerTopic $bulletsPerTopic

    [System.Windows.Forms.MessageBox]::Show("Presentation created successfully!`n$outputFile", "Success", "OK", "Information")
    Update-Status "Done!"
}

# ---------- Run the GUI ----------
$form.ShowDialog() | Out-Null
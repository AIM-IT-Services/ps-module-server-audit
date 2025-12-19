<#
.SYNOPSIS
Generate an HTML file/folder audit report with TreeView, Dark Mode toggle,
and click-to-open file details modal. Compatible with Windows PowerShell 5.1
and servers without System.Web.
#>

function New-FileDirectoryAuditReport {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]
        [string]$DrivePath,

        [string]$ClientName  = "Client",
        [string]$CompanyName = "AIM IT Services",

        [string]$OutputFile,
        [switch]$OpenInBrowser,

        # Files created OR modified within the last N years
        [int]$LastYears = 5,

        [switch]$TreeView,
        [switch]$DarkMode
    )

    # ---------- Safe string helpers (no System.Web dependency) ----------
    function To-SafeString {
        param($Value)
        if ($null -eq $Value) { return '' }
        return [string]$Value
    }

    function Html-Encode {
        param([string]$Text)
        if ($null -eq $Text) { return '' }
        $t = [string]$Text
        $t = $t -replace '&', '&amp;'
        $t = $t -replace '<', '&lt;'
        $t = $t -replace '>', '&gt;'
        $t = $t -replace '"', '&quot;'
        $t = $t -replace "'", '&#39;'
        return $t
    }

    function Html-AttrEncode {
        param([string]$Text)
        # Same as encode, but also protect newlines/tabs so attributes don't break
        $t = Html-Encode $Text
        $t = $t -replace "`r", '&#13;'
        $t = $t -replace "`n", '&#10;'
        $t = $t -replace "`t", '&#9;'
        return $t
    }

    # ---------- Validate path ----------
    $DrivePath = (Resolve-Path -Path $DrivePath).ProviderPath.TrimEnd('\')
    if (-not (Test-Path $DrivePath)) { throw "DrivePath '$DrivePath' was not found." }

    # ---------- Output file ----------
    if (-not $OutputFile) {
        $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeClient = ($ClientName -replace '[^\w\-]', '_')
        $OutputFile = Join-Path (Get-Location) "$safeClient-FileDirectoryAudit-$timestamp.html"
    } elseif ((Test-Path $OutputFile) -and (Get-Item $OutputFile).PSIsContainer) {
        $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeClient = ($ClientName -replace '[^\w\-]', '_')
        $OutputFile = Join-Path $OutputFile "$safeClient-FileDirectoryAudit-$timestamp.html"
    }

    # ---------- Date window ----------
    $endDate   = Get-Date
    $startDate = $endDate.AddYears(-$LastYears)
    Write-Verbose "Scanning files created OR modified between $startDate and $endDate"

    # ---------- Collect files ----------
    $files = Get-ChildItem -Path $DrivePath -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.CreationTime -ge $startDate -and $_.CreationTime -le $endDate) -or
            ($_.LastWriteTime -ge $startDate -and $_.LastWriteTime -le $endDate)
        } |
        ForEach-Object {
            $owner = try { (Get-Acl $_.FullName -ErrorAction Stop).Owner } catch { 'Unknown' }

            [PSCustomObject]@{
                Name         = $_.Name
                FullName     = $_.FullName
                SizeKB       = if ($null -ne $_.Length) { [math]::Round($_.Length / 1KB, 2) } else { 0 }
                CreatedTime  = $_.CreationTime
                ModifiedTime = $_.LastWriteTime
                Owner        = $owner
            }
        }

    # ---------- Report strings ----------
    $reportDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $windowText = "$($startDate.ToString('yyyy-MM-dd')) → $($endDate.ToString('yyyy-MM-dd'))"
    $summary    = "Path: $DrivePath | Items: $($files.Count) | Window: $windowText (created OR modified)"

    # ---------- Build directory tree ----------
    function New-DirectoryTree {
        param([string]$BasePath, $FileList)

        # Root node supports Files as well
        $tree = @{ Files = @() }

        foreach ($f in $FileList) {
            if (-not $f.FullName) { continue }

            $relative = $f.FullName.Substring($BasePath.Length).TrimStart('\')

            # file directly in root (no backslash in relative path)
            if ($relative -notmatch '\\') {
                $tree.Files += $f
                continue
            }

            $parts   = $relative -split '\\'
            $current = $tree

            for ($i = 0; $i -lt $parts.Length; $i++) {
                $part = $parts[$i]

                if ($i -eq $parts.Length - 1) {
                    if (-not $current.ContainsKey('Files')) { $current['Files'] = @() }
                    $current['Files'] += $f
                } else {
                    if (-not $current.ContainsKey($part)) { $current[$part] = @{ Files = @() } }
                    $current = $current[$part]
                }
            }
        }

        return $tree
    }

    # ---------- Render files list for a node ----------
    function Render-FilesHtml {
        param($Files, [int]$Level)

        $indent = '    ' * $Level
        if (-not $Files -or $Files.Count -eq 0) { return "" }

        $out = "$indent<ul class='files'>`n"
        foreach ($file in $Files) {
            if (-not $file.FullName) { continue }

            $createdText  = if ($file.CreatedTime)  { $file.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
            $modifiedText = if ($file.ModifiedTime) { $file.ModifiedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
            $sizeText     = if ($null -ne $file.SizeKB) { ([string]$file.SizeKB) } else { "0" }

            $ownerText = To-SafeString $file.Owner
            if ([string]::IsNullOrWhiteSpace($ownerText)) { $ownerText = "Unknown" }

            $nameVisible  = Html-Encode (To-SafeString $file.Name)

            $attrName     = Html-AttrEncode (To-SafeString $file.Name)
            $attrFullPath = Html-AttrEncode (To-SafeString $file.FullName)
            $attrOwner    = Html-AttrEncode $ownerText
            $attrCreated  = Html-AttrEncode $createdText
            $attrModified = Html-AttrEncode $modifiedText
            $attrSizeKB   = Html-AttrEncode $sizeText

            $out += "$indent  <li class='file'>" +
                    "<a href='#' class='file-link' " +
                    "data-name='$attrName' " +
                    "data-fullpath='$attrFullPath' " +
                    "data-owner='$attrOwner' " +
                    "data-created='$attrCreated' " +
                    "data-modified='$attrModified' " +
                    "data-sizekb='$attrSizeKB'>$nameVisible</a> " +
                    "<span class='meta'>($sizeText KB)</span></li>`n"
        }
        $out += "$indent</ul>`n"
        return $out
    }

    # ---------- Tree HTML ----------
    function ConvertTo-TreeHtml {
        param([string]$Name, $Node, [int]$Level = 0)

        $indent = '    ' * $Level
        $safeFolder = Html-Encode (To-SafeString $Name)

        $html = "$indent<li class='folder'><span>$safeFolder</span>`n"

        if ($Node.ContainsKey('Files')) {
            $html += (Render-FilesHtml -Files $Node['Files'] -Level ($Level + 1))
        }

        $subfolders = @($Node.Keys | Where-Object { $_ -ne 'Files' } | Sort-Object)
        if ($subfolders.Count -gt 0) {
            $html += "$indent  <ul class='folders'>`n"
            foreach ($k in $subfolders) {
                $html += (ConvertTo-TreeHtml -Name $k -Node $Node[$k] -Level ($Level + 2))
            }
            $html += "$indent  </ul>`n"
        }

        $html += "$indent</li>`n"
        return $html
    }

    # ---------- Body content ----------
    if ($TreeView) {
        $tree = New-DirectoryTree -BasePath $DrivePath -FileList $files

        $bodyContent  = "<ul class='folders-root'>`n"

        # root-level files first
        $bodyContent += "<li class='folder expanded'><span>(Root)</span>`n"
        $bodyContent += (Render-FilesHtml -Files $tree['Files'] -Level 1)
        $bodyContent += "</li>`n"

        # then top-level folders
        foreach ($k in ($tree.Keys | Where-Object { $_ -ne 'Files' } | Sort-Object)) {
            $bodyContent += (ConvertTo-TreeHtml -Name $k -Node $tree[$k] -Level 0)
        }

        $bodyContent += "</ul>`n"
    } else {
        $bodyContent = "<p>Run with <b>-TreeView</b> to see results.</p>"
    }

    # ---------- Styles ----------
    $darkClass = if ($DarkMode) { 'dark' } else { '' }

    $styles = @'
<style>
:root { --bg:#fff; --fg:#111; --muted:#666; --border:#ddd; --card:#fff; --shadow: rgba(0,0,0,0.15); }
body.dark { --bg:#111; --fg:#e8e8e8; --muted:#aaa; --border:#333; --card:#1b1b1b; --shadow: rgba(0,0,0,0.55); }

body { font-family: Segoe UI, Arial, sans-serif; margin: 16px; background: var(--bg); color: var(--fg); }

.header {
  display:flex; align-items:flex-start; justify-content:space-between; gap: 16px;
  padding: 14px 16px; border:1px solid var(--border); border-radius: 10px;
  background: var(--card); box-shadow: 0 10px 24px var(--shadow); margin-bottom: 14px;
}
.header h1 { margin: 0; font-size: 20px; }
.header .sub { margin-top: 6px; color: var(--muted); font-size: 13px; line-height: 1.35; }

.btn { padding: 6px 10px; border-radius: 6px; border: 1px solid #bbb; background: #f6f6f6; cursor: pointer; }
body.dark .btn { background: #222; border-color:#444; color:#eee; }

.folder > span { cursor:pointer; font-weight:600; }
.files, .folders { display:none; margin-left:16px; padding-left: 12px; }
.folder.expanded > .files, .folder.expanded > .folders { display:block; }

.file { margin: 3px 0; }
.meta { color: var(--muted); font-size: 0.9em; margin-left: 8px; }
.file-link { color: inherit; text-decoration: underline; cursor: pointer; }

.modal-overlay { position:fixed; inset:0; background:rgba(0,0,0,.6); display:flex; align-items:center; justify-content:center; padding: 24px; z-index:9999; }
.modal-overlay.hidden { display:none; }
.modal { background: var(--card); color: var(--fg); padding: 16px; width: min(900px, 95vw); border-radius: 10px; box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
.modal-header { display:flex; justify-content:space-between; align-items:center; margin-bottom: 10px; }
.modal-title { font-weight: 700; }
.modal-close { border:1px solid #bbb; background:#f6f6f6; border-radius:6px; padding:4px 10px; cursor:pointer; }
body.dark .modal-close { background:#222; border-color:#444; color:#eee; }

.kv { display:grid; grid-template-columns: 160px 1fr; gap: 6px 12px; }
.kv .key { color: var(--muted); }
.code { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; white-space: pre-wrap; word-break: break-word; }
</style>
'@

    # ---------- Scripts ----------
    $scripts = @'
<script>
function escapeHtml(s) {
  return (s ?? '').toString()
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'","&#039;");
}

function toggleDarkMode() {
  document.body.classList.toggle('dark');
}

function showModal(d) {
  const overlay = document.getElementById('fileModalOverlay');
  const body = document.getElementById('modalBody');
  const title = document.getElementById('modalTitle');
  if (!overlay || !body || !title) return;

  title.textContent = d.name || 'File Details';

  body.innerHTML = `
    <div class="kv">
      <div class="key"><b>Name</b></div><div class="code">${escapeHtml(d.name)}</div>
      <div class="key"><b>Full path</b></div><div class="code">${escapeHtml(d.fullpath)}</div>
      <div class="key"><b>Owner</b></div><div class="code">${escapeHtml(d.owner)}</div>
      <div class="key"><b>Created</b></div><div class="code">${escapeHtml(d.created)}</div>
      <div class="key"><b>Modified</b></div><div class="code">${escapeHtml(d.modified)}</div>
      <div class="key"><b>Size (KB)</b></div><div class="code">${escapeHtml(d.sizekb)}</div>
    </div>
  `;

  overlay.classList.remove('hidden');
  overlay.setAttribute('aria-hidden','false');
}

function hideModal() {
  const overlay = document.getElementById('fileModalOverlay');
  if (!overlay) return;
  overlay.classList.add('hidden');
  overlay.setAttribute('aria-hidden','true');
}

window.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('li.folder > span').forEach(s => {
    s.addEventListener('click', (e) => {
      e.stopPropagation();
      s.parentElement.classList.toggle('expanded');
    });
  });

  document.querySelectorAll('a.file-link').forEach(a => {
    a.addEventListener('click', (e) => {
      e.preventDefault();
      showModal({
        name: a.dataset.name || '',
        fullpath: a.dataset.fullpath || '',
        owner: a.dataset.owner || '',
        created: a.dataset.created || '',
        modified: a.dataset.modified || '',
        sizekb: a.dataset.sizekb || ''
      });
    });
  });

  document.getElementById('modalCloseBtn')?.addEventListener('click', hideModal);

  document.getElementById('fileModalOverlay')?.addEventListener('click', (e) => {
    if (e.target.id === 'fileModalOverlay') hideModal();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') hideModal();
  });
});
</script>
'@

    # ---------- HTML ----------
    $darkToggleButton = "<button class='btn' type='button' onclick='toggleDarkMode()'>Toggle Dark Mode</button>"

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>$CompanyName - $ClientName - Audit Report</title>
$styles
</head>
<body class='$darkClass'>

<div class="header">
  <div>
    <h1>$CompanyName</h1>
    <div class="sub">
      <div><b>Client:</b> $(Html-Encode $ClientName)</div>
      <div><b>Generated:</b> $(Html-Encode $reportDate)</div>
      <div><b>$(Html-Encode $summary)</b></div>
    </div>
  </div>
  <div>$darkToggleButton</div>
</div>

<div id="fileModalOverlay" class="modal-overlay hidden" aria-hidden="true">
  <div class="modal" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
    <div class="modal-header">
      <div class="modal-title" id="modalTitle">File Details</div>
      <button class="modal-close" id="modalCloseBtn" type="button" aria-label="Close">X</button>
    </div>
    <div id="modalBody"></div>
  </div>
</div>

$bodyContent

$scripts
</body>
</html>
"@

    if ($PSCmdlet.ShouldProcess($OutputFile, "Write HTML report")) {
        Set-Content -Path $OutputFile -Value $html -Encoding UTF8
        Write-Output "Report written to $OutputFile"
    }

    if ($OpenInBrowser) {
        Start-Process $OutputFile
    }
}

<# 
.SYNOPSIS
  Find-Text — Search files and optionally contents across drives or folders.

.DESCRIPTION
  • Name search (plain, wildcard -like, regex) with optional fuzzy scoring.
  • Content search for: txt, md, csv, log, ini, docx, xlsx, pptx, pdf (via pdftotext.exe),
    and **code/source files**: cpp, c, cc, h, hpp, inl, cs, java, py, js, ts, rs, go, swift,
    kt, m, mm, ps1, psm1, bat, cmd, sh, yaml, yml, json, xml.
  • NEW: -WholeWord for whole-word matching in names/paths AND content.
  • Snippet preview around the first content hit.
  • Filters: roots, include/exclude extensions, date since, size range, max results.
  • Traversal controls: exclude directories, max depth, folder trace (-ShowFolders).
  • Output: table/list, JSON, or export to CSV.
  • Works in Windows PowerShell 5.1 and PowerShell 7+.

.NOTES
  Save as:  find_text.ps1
  Usage:    .\find_text.ps1 -Query "foo" -Content -Snippet -ShowFolders
#>

[CmdletBinding()]
param(
  [string] $Query,
  [string[]] $Queries,
  [string[]] $Roots,
  [string[]] $IncludeExt,
  [string[]] $ExcludeExt,
  [switch] $NameOnly,
  [switch] $Content,
  [ValidateSet('Auto','Indexed','Brute')]
  [string] $Mode = 'Auto',
  [Nullable[datetime]] $Since,
  [int] $MinSizeKB,
  [int] $MaxSizeKB,
  [int] $MaxResults = 500,
  [switch] $Fuzzy,
  [switch] $Regex,
  [switch] $Wildcard,
  [switch] $WholeWord,      # <— NEW
  [switch] $Snippet,
  [switch] $EnablePdf,

  # Usability options
  [switch] $ShowFolders,
  [string[]] $ExcludeDir = @('bin','obj','.git','node_modules','__pycache__','.idea','.vs'),
  [int] $MaxDepth = 0,                  # 0 = unlimited
  [string] $ExportCsv,                  # path to write CSV
  [switch] $AsJson                      # emit JSON instead of formatted view
)

# Default fuzzy ON if not specified
if (-not $PSBoundParameters.ContainsKey('Fuzzy')) { $Fuzzy = $true }

# Default extensions (you can override with -IncludeExt)
$DefaultIncludeExt = @('pdf','docx','xlsx','pptx','txt','md','csv','log','ini','cpp','h','hpp','inl')

# -------------------------------
# Levenshtein (PS5 safe)
# -------------------------------
function Get-Levenshtein([string]$a, [string]$b) {
  if ($a -eq $b) { return 0 }
  if ([string]::IsNullOrEmpty($a)) { return $b.Length }
  if ([string]::IsNullOrEmpty($b)) { return $a.Length }
  $lenA=[int]$a.Length; $lenB=[int]$b.Length
  $prev = New-Object int[] ($lenB + 1)
  $curr = New-Object int[] ($lenB + 1)
  for ($j=0;$j -le $lenB;$j++){ $prev[$j]=$j }
  for ($i=1;$i -le $lenA;$i++){
    $curr[0]=$i; $ai=$a[$i-1]
    for ($j=1;$j -le $lenB;$j++){
      $cost = if ($ai -ceq $b[$j-1]) {0} else {1}
      $del=$prev[$j]+1; $ins=$curr[$j-1]+1; $sub=$prev[$j-1]+$cost
      $m=$del; if ($ins -lt $m){$m=$ins}; if ($sub -lt $m){$m=$sub}; $curr[$j]=$m
    }
    $tmp=$prev; $prev=$curr; $curr=$tmp
  }
  return $prev[$lenB]
}

# -------------------------------
# Name/path match (plain/wild/regex + fuzzy score)
# -------------------------------
function Test-MatchName {
  param(
    [string]$text,
    [string[]]$terms,
    [switch]$Regex,
    [switch]$Wildcard,
    [switch]$Fuzzy,
    [switch]$WholeWord
  )

  $textSafe = if ($null -eq $text) { '' } else { [string]$text }
  $textLC = $textSafe.ToLowerInvariant()
  $score = 0; $hits = 0

  foreach ($t in $terms) {
    if ([string]::IsNullOrWhiteSpace($t)) { continue }
    $tLC = $t.ToLowerInvariant(); $matched = $false

    if ($Regex) {
      try { if ($textSafe -match $t) { $matched = $true } } catch {}
    } elseif ($Wildcard) {
      if ($textLC -like $tLC) { $matched = $true }
    } elseif ($WholeWord) {
      $pat = '(?i)\b' + [regex]::Escape($t) + '\b'
      if ($textSafe -match $pat) { $matched = $true }
    } else {
      if ($textLC.Contains($tLC)) { $matched = $true }
    }

    if (-not $matched -and $Fuzzy) {
      $dist = Get-Levenshtein $textLC $tLC
      if ($dist -le [Math]::Max(1,[int]([double]$tLC.Length*0.15))) { $matched = $true }
      $score += [Math]::Max(0,10-$dist)
    }

    if ($matched) { $hits++; $score+=20 }
  }

  return @{ Match=($hits -gt 0); Score=$score+($hits*5) }
}

# -------------------------------
# PDF text support (optional)
# -------------------------------
function Find-Pdftotext {
  $paths = @(
    'C:\Tools\pdftotext.exe',
    'C:\Program Files\Poppler\bin\pdftotext.exe'
  )
  foreach ($p in $paths) { if (Test-Path $p) { return $p } }
  $cmd = Get-Command pdftotext.exe -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  return $null
}

function Invoke-Pdftotext([string]$PdfPath,[string]$Exe) {
  try {
    if (-not (Test-Path $PdfPath)) { return '' }
    $tmp = [IO.Path]::GetTempFileName()
    $out = "$tmp.txt"
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = "-layout -nopgbrk -q `"$PdfPath`" `"$out`""
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    [void]$p.WaitForExit()
    $t = if (Test-Path $out) { Get-Content $out -Raw -ErrorAction SilentlyContinue } else { '' }
    if (Test-Path $out) { Remove-Item $out -ErrorAction SilentlyContinue }
    if (Test-Path $tmp) { Remove-Item $tmp -ErrorAction SilentlyContinue }
    if ($t) { return $t } else { return '' }
  } catch { return '' }
}

# -------------------------------
# Office text extractors (docx/xlsx/pptx)
# -------------------------------
try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch {}

function Read-ZipText($zip,$path) {
  $e = $zip.GetEntry($path); if (-not $e){return ''}
  $r = New-Object IO.StreamReader($e.Open())
  $t = $r.ReadToEnd()
  $r.Close()
  return $t
}

function Get-DocxText($path) {
  try {
    $fs=[IO.File]::OpenRead($path); $z=New-Object IO.Compression.ZipArchive($fs)
    $x=Read-ZipText $z 'word/document.xml'
    $z.Dispose();$fs.Dispose()
    $plain = ([regex]::Replace($x,'<[^>]+>',' ')) -replace '\s+',' '
    return $plain
  } catch { return '' }
}

function Get-XlsxText($path) {
  try {
    $fs=[IO.File]::OpenRead($path);$z=New-Object IO.Compression.ZipArchive($fs)
    $s=Read-ZipText $z 'xl/sharedStrings.xml'
    $z.Dispose();$fs.Dispose()
    $plain = ([regex]::Replace($s,'<[^>]+>',' ')) -replace '\s+',' '
    return $plain
  } catch { return '' }
}

function Get-PptxText($path) {
  try {
    $fs=[IO.File]::OpenRead($path);$z=New-Object IO.Compression.ZipArchive($fs)
    $slides=$z.Entries|Where-Object{$_.FullName -like 'ppt/slides/slide*.xml'}
    $sb=New-Object Text.StringBuilder
    foreach($s in $slides){
      $x=Read-ZipText $z $s.FullName
      if($x){$null=$sb.Append(' '+$x)}
    }
    $z.Dispose();$fs.Dispose()
    $plain = ([regex]::Replace($sb.ToString(),'<[^>]+>',' ')) -replace '\s+',' '
    return $plain
  } catch { return '' }
}

# -------------------------------
# Snippet helpers
# -------------------------------
function Get-FirstSnippet {
  param([string] $Text,[string[]] $Terms,[int] $Radius = 60)
  if ([string]::IsNullOrEmpty($Text)) { return $null }
  $lc = $Text.ToLowerInvariant()
  foreach ($t in $Terms) {
    if ([string]::IsNullOrWhiteSpace($t)) { continue }
    $i = $lc.IndexOf($t.ToLowerInvariant())
    if ($i -ge 0) {
      $start = [Math]::Max(0, $i - $Radius)
      $len   = [Math]::Min(($Radius * 2), $Text.Length - $start)
      return ("… " + $Text.Substring($start, $len).Trim() + " …")
    }
  }
  return $null
}

function New-WordRegexFromTerms {
  param([string[]]$Terms)
  $escaped = $Terms | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_) }
  if (-not $escaped -or $escaped.Count -eq 0) { return $null }
  return '(?i)\b(' + ($escaped -join '|') + ')\b'
}

function Get-FirstSnippetRegex {
  param([string]$Text,[string]$Pattern,[int]$Radius=60)
  if ([string]::IsNullOrEmpty($Text) -or [string]::IsNullOrEmpty($Pattern)) { return $null }
  $m = [regex]::Match($Text,$Pattern)
  if (-not $m.Success) { return $null }
  $i = $m.Index
  $start = [Math]::Max(0,$i-$Radius)
  $len = [Math]::Min(($Radius*2), $Text.Length-$start)
  return ("… " + $Text.Substring($start,$len).Trim() + " …")
}

# -------------------------------
# File enumerator with dir excludes, depth, and folder tracing
# -------------------------------
function Get-Files {
  param(
    [string[]] $roots,
    [string[]] $inc,
    [string[]] $exc,
    [Nullable[datetime]] $since,
    [int] $minKB,
    [int] $maxKB,
    [string[]] $ExcludeDir,
    [int] $MaxDepth,
    [switch] $ShowFolders
  )

  $out = @()

  foreach ($root in $roots) {
    if (-not (Test-Path $root)) { continue }

    $startPath = (Resolve-Path $root).Path
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push(@{ Path = $startPath; Depth = 0 })

    while ($stack.Count -gt 0) {
      $d = $stack.Pop()
      $path = $d.Path
      $depth = [int]$d.Depth

      if ($ShowFolders) { Write-Host "[scan] $path" }

      # subdirs
      if ($MaxDepth -eq 0 -or $depth -lt $MaxDepth) {
        Get-ChildItem -Path $path -Directory -Force -ErrorAction SilentlyContinue |
          Where-Object { 
            $name = $_.Name
            -not ($ExcludeDir | Where-Object { $name -ieq $_ })
          } |
          ForEach-Object {
            $stack.Push(@{ Path = $_.FullName; Depth = $depth + 1 })
          }
      }

      # files
      Get-ChildItem -Path $path -File -Force -ErrorAction SilentlyContinue |
        ForEach-Object {
          $e = $_.Extension.TrimStart('.').ToLowerInvariant()
          if ($inc -and $inc.Count -and ($inc -notcontains $e)) { return }
          if ($exc -and $exc.Count -and ($exc -contains $e)) { return }
          if ($since -and $_.LastWriteTime -lt $since.Value) { return }
          $kb = [int]([double]$_.Length/1024)
          if ($minKB -and $kb -lt $minKB) { return }
          if ($maxKB -and $kb -gt $maxKB) { return }

          $out += [pscustomobject]@{
            Name      = $_.Name
            Path      = $_.FullName
            Ext       = $e
            SizeKB    = $kb
            LastWrite = $_.LastWriteTime
          }
        }
    }
  }
  return $out
}

# -------------------------------
# Main
# -------------------------------
if (-not $Roots) { $Roots = @("$PWD") }
$inc = if ($IncludeExt) { $IncludeExt } else { $DefaultIncludeExt }

$terms = @()
if ($Query)   { $terms += $Query }
if ($Queries) { $terms += $Queries }
$terms = $terms | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

if (-not $terms.Count) {
  Write-Error "Provide -Query or -Queries"
  exit 1
}

if ($NameOnly) { $Content = $false }

# Build a single regex for content if WholeWord or Regex
$ContentPattern = $null
if ($Regex) {
  $ContentPattern = '(?i)(' + ($terms -join '|') + ')'
} elseif ($WholeWord) {
  $ContentPattern = New-WordRegexFromTerms -Terms $terms
}

$files = Get-Files -roots $Roots -inc $inc -exc $ExcludeExt -since $Since `
                   -minKB $MinSizeKB -maxKB $MaxSizeKB `
                   -ExcludeDir $ExcludeDir -MaxDepth $MaxDepth -ShowFolders:$ShowFolders

$results = @()

# Extensions we will treat as code/source (content-searched)
$CodeExt = @('cpp','c','cc','h','hpp','inl','cs','java','py','js','ts','rs','go','swift','kt','m','mm','ps1','psm1','bat','cmd','sh','yaml','yml','json','xml')

foreach ($f in $files) {
  # Name/path scoring (now aware of -WholeWord)
  $nm = Test-MatchName -text ($f.Name + ' ' + $f.Path) -terms $terms -Regex:$Regex -Wildcard:$Wildcard -Fuzzy:$Fuzzy -WholeWord:$WholeWord
  $hadNameMatch = [bool]$nm.Match
  $score = [int]$nm.Score
  $hitType = if ($hadNameMatch) { 'Name' } else { '' }
  $snippetText = $null
  $contentHit = $false

  if ($Content) {
    switch ($f.Ext) {
      # Plain text-ish
      { $_ -in @('txt','md','csv','log','ini','json','xml','yaml','yml','ps1','psm1','bat','cmd','sh') } {
        $raw = Get-Content $f.Path -Raw -ErrorAction SilentlyContinue
        if ($raw) {
          if ($raw.Length -gt 2000000) { $raw = $raw.Substring(0,2000000) }
          if ($Regex -or $WholeWord) {
            try {
              if ($ContentPattern -and ($raw -match $ContentPattern)) {
                $snippetText = Get-FirstSnippetRegex -Text $raw -Pattern $ContentPattern
              }
            } catch {}
          } else {
            $snippetText = Get-FirstSnippet -Text $raw -Terms $terms
          }
          if ($snippetText) { $hitType = 'Content'; $contentHit = $true; $score += 80 }
        }
        break
      }

      # Code/source files
      { $_ -in $CodeExt } {
        $raw = Get-Content $f.Path -Raw -ErrorAction SilentlyContinue
        if ($raw) {
          if ($raw.Length -gt 2000000) { $raw = $raw.Substring(0,2000000) }
          if ($Regex -or $WholeWord) {
            try {
              if ($ContentPattern -and ($raw -match $ContentPattern)) {
                $snippetText = Get-FirstSnippetRegex -Text $raw -Pattern $ContentPattern
              }
            } catch {}
          } else {
            $snippetText = Get-FirstSnippet -Text $raw -Terms $terms
          }
          if ($snippetText) { $hitType = 'Content'; $contentHit = $true; $score += 80 }
        }
        break
      }

      # Office docs
      'docx' {
        $t = Get-DocxText $f.Path
        if ($t) {
          if ($Regex -or $WholeWord) {
            try { if ($ContentPattern -and ($t -match $ContentPattern)) { $snippetText = Get-FirstSnippetRegex -Text $t -Pattern $ContentPattern } } catch {}
          } else {
            $snippetText = Get-FirstSnippet -Text $t -Terms $terms
          }
        }
        if ($snippetText) { $hitType = 'Content'; $contentHit = $true; $score += 80 }
        break
      }
      'xlsx' {
        $t = Get-XlsxText $f.Path
        if ($t) {
          if ($Regex -or $WholeWord) {
            try { if ($ContentPattern -and ($t -match $ContentPattern)) { $snippetText = Get-FirstSnippetRegex -Text $t -Pattern $ContentPattern } } catch {}
          } else {
            $snippetText = Get-FirstSnippet -Text $t -Terms $terms
          }
        }
        if ($snippetText) { $hitType = 'Content'; $contentHit = $true; $score += 80 }
        break
      }
      'pptx' {
        $t = Get-PptxText $f.Path
        if ($t) {
          if ($Regex -or $WholeWord) {
            try { if ($ContentPattern -and ($t -match $ContentPattern)) { $snippetText = Get-FirstSnippetRegex -Text $t -Pattern $ContentPattern } } catch {}
          } else {
            $snippetText = Get-FirstSnippet -Text $t -Terms $terms
          }
        }
        if ($snippetText) { $hitType = 'Content'; $contentHit = $true; $score += 80 }
        break
      }

      # PDFs
      'pdf' {
        if ($EnablePdf) {
          $exe = Find-Pdftotext
          if ($exe) {
            $t = Invoke-Pdftotext $f.Path $exe
            if ($t) {
              if ($t.Length -gt 2000000) { $t = $t.Substring(0,2000000) }
              if ($Regex -or $WholeWord) {
                try { if ($ContentPattern -and ($t -match $ContentPattern)) { $snippetText = Get-FirstSnippetRegex -Text $t -Pattern $ContentPattern } } catch {}
              } else {
                $snippetText = Get-FirstSnippet -Text $t -Terms $terms
              }
            }
            if ($snippetText) { $hitType = 'Content'; $contentHit = $true; $score += 80 }
          }
        }
        break
      }
    }
  }

  # If we had NO name match and NO content hit, skip this file
  if (-not $hadNameMatch -and -not $contentHit) { continue }

  $results += [pscustomobject]@{
    Score     = $score
    Name      = $f.Name
    Ext       = $f.Ext
    SizeKB    = $f.SizeKB
    LastWrite = $f.LastWrite
    Hit       = if ($hitType) { $hitType } else { if ($hadNameMatch){'Name'} else {'Content'} }
    Path      = $f.Path
    Snippet   = if ($Snippet) { $snippetText } else { $null }
  }

  if ($results.Count -ge $MaxResults) { break }
}

if (-not $results -or $results.Count -eq 0) { Write-Host 'No matches.'; exit }

$results = $results |
  Sort-Object @{e='Score';Descending=$true}, @{e='LastWrite';Descending=$true}

# Export / JSON / default view
if ($ExportCsv) {
  $results | Export-Csv -Path $ExportCsv -NoTypeInformation
  Write-Host "Exported to $ExportCsv"
  exit
}

if ($AsJson) {
  $results | ConvertTo-Json -Depth 5
  exit
}

# Default: full details without truncation
$results |
  Select-Object Name,Ext,SizeKB,LastWrite,Hit,Path,Snippet |
  Format-List

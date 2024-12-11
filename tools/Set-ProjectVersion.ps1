<#PSScriptInfo .VERSION 1.0.1#>

[CmdletBinding()]
param ([string] $Root)

git add *.vbs *.ps*1

& {
  # Set the version of the executable
  $infoFilePath = "$Root\cvmd2html.vbs"
  $diffLines = git diff HEAD *.vbs
  if ($null -ne $diffLines) {
    $matchRegex = '>(?<Version>((\d+\.){3})\d+)</'
    $diff = [Math]::Abs(($diffLines | ForEach-Object { switch ($_[0]) { "+"{1} "-"{-1} } } | Measure-Object -Sum).Sum)
    $version = ($diff -eq 0 ? 1:$diff) + ([version](git cat-file -p HEAD:index.vbs | Where-Object { $_ -match $matchRegex } | Select-Object -Last 1 | ForEach-Object { [void]($_ -match $matchRegex); $Matches.Version })).Revision
    $infoContent = (Get-Content $infoFilePath | ForEach-Object { if ($_ -match $matchRegex) { $_ -replace $matchRegex,">`${1}$version</" } else { $_ } }) -join [Environment]::NewLine
    Set-Content $infoFilePath $infoContent -NoNewline
  }
}

function Set-SourceVersion([string] $ExtensionPattern, [string] $Filter, [scriptblock] $VersionGetter, [string] $VersionMatch, [string] $ReplacementFormat) {
  # Set the version of the source files
  git status -s $ExtensionPattern | ConvertFrom-StringData -Delimiter ' ' | Where-Object { $_.Keys[0].EndsWith('M') } | ForEach-Object { $_.Values } | Where-Object { $_ -ne $Filter } |
  ForEach-Object {
    $version = git cat-file -p HEAD:$_ 2>&1 | Select-Object -First 2 | Where-Object { $_ -match $VersionMatch } | ForEach-Object $VersionGetter | Select-Object -Last 1
    if (-not [String]::IsNullOrWhiteSpace($version)) {
      $content = (Get-Content $_ -Raw) -replace $VersionMatch,($ReplacementFormat -f $version)
      Set-Content "$Root\$_" $content -NoNewline
    }
  }
}

Set-SourceVersion *.ps*1 'rsc/*.ps1' { ([version]($_.TrimEnd().Substring('<#PSScriptInfo .VERSION '.Length) -split '#')[0]).Build + 1 } '<#PSScriptInfo .VERSION ((\d+\.){2})\d+#>' "<#PSScriptInfo .VERSION `${{1}}{0}#>"
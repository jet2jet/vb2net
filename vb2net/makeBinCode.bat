@setlocal enableextensions & cd /d "%~dp0" & PowerShell.exe -Command "& (iex -Command ('{#' + ((gc '%~nx0') -join \"`n\") + '}'))" %* & exit /b
#
# This generates Base64-encoded tar-gzip-ped vb2net dll binary as a VBA code.
# After executing this batch, ".work\vb2net.code.txt" will be generated, so
# we will replace 'GetVb2netBinary' code in vb2net.bas with generated code.
#

$DLL_PATH = "bin\Release\net6.0-windows\vb2net.dll"
$WORK_DIR = ".work"

New-Item -ItemType Directory -Force -Path $WORK_DIR | Out-Null

$tempTarGz = "$WORK_DIR\.temp.tar.gz"
tar czf $tempTarGz "$DLL_PATH"

$tarGzBin = Get-Content $tempTarGz -Encoding Byte
$tarGzBinB64 = [System.Convert]::ToBase64String($tarGzBin) -replace '.{256}', "`$&`n"
$source = ($tarGzBinB64.Split("`n") | ForEach-Object { "    bin = bin + `"$_`"" }) -join "`r`n"

Set-Content -Path "$WORK_DIR\vb2net.code.txt" -Value $source
Remove-Item $tempTarGz -Force

import base64
import json
import shlex
import subprocess
from abc import ABC, abstractmethod

CMD_TEMPLATE = '''\
chcp.com 65001 >nul 2>&1 && PowerShell -NoProfile -NonInteractive -ExecutionPolicy Unrestricted -EncodedCommand %s'''

POWERSHELL_TEMPLATE = '''\
chcp.com 65001 > $null; PowerShell -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand %s'''

ENV_PS1 = '''\
$acc = @{}
Get-ChildItem env: | ForEach-Object { $acc[$_.Name] = $_.Value }
ConvertTo-Json $acc -Compress'''

QUSER_PS1 = r'''
function Get-Quser
{
    [CmdletBinding()]
    Param (
    )

    $outputEncodingBackup = [System.Console]::OutputEncoding
    [System.Console]::OutputEncoding = [System.Text.Encoding]::Default
    $result = quser 2>&1
    [System.Console]::OutputEncoding = $outputEncodingBackup

    if ($LASTEXITCODE -ne 0)
    {
        return ,@()
    }

    $headerRow = $result[0]

    $match = [Regex]::Match($headerRow, '(\s{2,})')
    $usernameLength = $match.Index + $match.Length

    return ,@($result[1..($result.Length - 1)] | Foreach-Object {
        $username = $_.Substring(0, $usernameLength)
        $_ = $_.Substring($usernameLength).Trim()

        $parts = $_ -split '\s{2,}'

        if ($parts.Length -eq 4)
        {
            $parts.Insert(0, '')
        }

        $quser = @{
            IsCurrentSession = $username.StartsWith('>')
            UserName = $username.TrimStart('>').Trim()
            SessionName = $parts[0]
            Id = [int]$parts[1]
            State = $parts[2]
            LogonTime = (Get-Date $parts[4]).ToString()
        }

        return $quser
    })
}

ConvertTo-Json (Get-Quser) -Compress'''

SETUP_PS1 = '''\
$tempDir = Join-Path $env:PUBLIC 'Documents'

if (-not (Test-Path $tempDir -PathType Container)) {
    $null = New-Item -Path $tempDir -ItemType Directory
}

try {
    Get-Command 'PsExec.exe' -ErrorAction Stop > $null
} catch {
    [System.Console]::Error.WriteLine("Command not found: 'PsExec.exe'") 
    exit 127
}

$vbsPath = Join-Path $tempDir 'a.vbs'
$screenshotPath = Join-Path $tempDir 'b.jpg'

$command = @'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing

$screens = [System.Windows.Forms.Screen]::AllScreens

$top    = ($screens.Bounds.Top    | Measure-Object -Minimum).Minimum
$left   = ($screens.Bounds.Left   | Measure-Object -Minimum).Minimum
$width  = ($screens.Bounds.Right  | Measure-Object -Maximum).Maximum
$height = ($screens.Bounds.Bottom | Measure-Object -Maximum).Maximum

$bounds   = [System.Drawing.Rectangle]::FromLTRB($left, $top, $width, $height)
$bmp      = New-Object System.Drawing.Bitmap ([int]$bounds.Width), ([int]$bounds.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
$bmp.Save('{0}', [System.Drawing.Imaging.ImageFormat]::Jpeg)

$graphics.Dispose()
$bmp.Dispose()
'@ -f $screenshotPath

$bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
$encodedCommand = [System.Convert]::ToBase64String($bytes)

$vbs = @'
Dim shell,command
command = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {0}"
Set shell = CreateObject("WScript.Shell")
shell.Run command, 0, true
'@ -f $encodedCommand
$vbs | Out-File -FilePath $vbsPath

ConvertTo-Json @{
    VbsPath = $vbsPath
    ScreenshotPath = $screenshotPath
} -Compress'''

SCREENSHOT_PS1 = '''\
$ErrorActionPreference = 'Stop'
$vbsPath = '%s'
$screenshotPath = '%s'

PsExec.exe -accepteula -nobanner -s -i %d wscript.exe //B //Nologo $vbsPath > $null

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$lastWriteTime = (Get-Item $screenshotPath).LastWriteTime.ToString()
$bytes = [System.IO.File]::ReadAllBytes($screenshotPath)
$encodedBytes = [System.Convert]::ToBase64String($bytes)

ConvertTo-Json @{
    LastWriteTime = $lastWriteTime
    Data = $encodedBytes
} -Compress'''

CLEANUP_PS1 = '''\
$vbsPath = '%s'
$screenshotPath = '%s'
Remove-Item -Path $vbsPath,$screenshotPath -Force -ErrorAction SilentlyContinue'''


class BaseRPCStub(ABC):
    def __init__(self, ssh_executable: str, ssh_args: str, host: str):
        self._command = [ssh_executable, *shlex.split(ssh_args), host]
        self._configs = {}

    @abstractmethod
    def _get_command_template(self) -> str:
        pass

    def run_script(self, script: str, check_returncode=True, **kwargs):
        command_template = self._get_command_template()
        encoded_script = base64.b64encode(script.encode('utf-16le')).decode('ascii')
        command = self._command + [command_template % encoded_script]
        options = dict(capture_output=True, text=True)
        options.update(kwargs)
        p = subprocess.run(command, **options)
        if check_returncode:
            p.check_returncode()
        return p

    def gather_env(self):
        p = self.run_script(ENV_PS1)
        return json.loads(p.stdout)

    def get_quser(self):
        p = self.run_script(QUSER_PS1)
        return json.loads(p.stdout)

    def setup(self):
        p = self.run_script(SETUP_PS1)
        configs = json.loads(p.stdout)
        self._configs.update(configs)

    def cleanup(self, force: bool = False):
        if not force and not self._configs:
            return 0
        script = CLEANUP_PS1 % (self._configs['VbsPath'], self._configs['ScreenshotPath'])
        p = self.run_script(script, check_returncode=False)
        self._configs.clear()
        return p.returncode

    def take_screenshot(self, sid: int):
        if not self._configs:
            self.setup()
        script = SCREENSHOT_PS1 % (self._configs['VbsPath'], self._configs['ScreenshotPath'], sid)
        p = self.run_script(script)
        return json.loads(p.stdout)


class CMDStub(BaseRPCStub):
    def __init__(self, ssh_executable: str, ssh_args: str, host: str):
        super().__init__(ssh_executable, ssh_args, host)

    def _get_command_template(self) -> str:
        return CMD_TEMPLATE


class PowerShellStub(BaseRPCStub):
    def __init__(self, ssh_executable: str, ssh_args: str, host: str):
        super().__init__(ssh_executable, ssh_args, host)

    def _get_command_template(self) -> str:
        return POWERSHELL_TEMPLATE

<#
.SYNOPSIS
    docx-format-applier 一键环境检查与安装脚本
.DESCRIPTION
    自动检查 .NET SDK、Python 是否安装，还原 NuGet 包，编译 FormatFix 项目。
    适用于首次使用该技能的用户。
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Green  = [ConsoleColor]::Green
$Red    = [ConsoleColor]::Red
$Yellow = [ConsoleColor]::Yellow
$Cyan   = [ConsoleColor]::Cyan

function Write-Step {
    param([string]$Message)
    Write-Host "`n>>> $Message" -ForegroundColor $Cyan
}

function Write-OK {
    Write-Host "  [OK] $($args[0])" -ForegroundColor $Green
}

function Write-Fail {
    Write-Host "  [!!] $($args[0])" -ForegroundColor $Red
}

function Write-Warn {
    Write-Host "  [..] $($args[0])" -ForegroundColor $Yellow
}

Write-Host "========================================" -ForegroundColor $Cyan
Write-Host " docx-format-applier 环境检查脚本" -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "项目目录: $ScriptDir`n"

$allOk = $true

# 1. 检查 .NET SDK
Write-Step "1/4 检查 .NET SDK..."
try {
    $dotnetVersion = & dotnet --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK ".NET SDK: v$dotnetVersion"
    } else {
        throw ".NET SDK 未就绪"
    }
} catch {
    Write-Fail ".NET SDK 未安装或不在 PATH 中"
    Write-Warn "安装命令: winget install Microsoft.DotNet.SDK.8"
    Write-Warn "或手动下载: https://dotnet.microsoft.com/download/dotnet/8.0"
    $allOk = $false
}

# 2. 检查 Python
Write-Step "2/4 检查 Python..."
try {
    $pythonVersion = & python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK "Python: $pythonVersion"
    } else {
        throw "Python 未就绪"
    }
} catch {
    Write-Fail "Python 未安装或不在 PATH 中"
    Write-Warn "安装命令: winget install Python.Python.3"
    Write-Warn "或手动下载: https://www.python.org/downloads/"
    $allOk = $false
}

# 3. 还原并编译 FormatFix 项目
Write-Step "3/4 查找并还原 FormatFix 项目..."

$formatFixDirs = Get-ChildItem -Path $ScriptDir -Directory -Filter "FormatFix*"
if ($formatFixDirs.Count -eq 0) {
    Write-Fail "未找到 FormatFix* 项目目录"
    $allOk = $false
} else {
    foreach ($dir in $formatFixDirs) {
        $csprojFiles = Get-ChildItem -Path $dir.FullName -Filter "*.csproj" -ErrorAction SilentlyContinue
        $excludeDebug = $csprojFiles | Where-Object { $_.Name -notmatch '^Debug\.' }
        $targetCsproj = $excludeDebug | Select-Object -First 1

        if ($targetCsproj) {
            Write-Host "  发现项目: $($dir.Name)\$($targetCsproj.Name)"

            if (& { try { & dotnet --version 2>$null; return $true } catch { return $false } }) {
                Push-Location $dir.FullName
                try {
                    Write-Host "  正在还原 NuGet 包..."
                    & dotnet restore $targetCsproj.Name 2>&1 | Out-Null
                    if ($LASTEXITCODE -eq 0) {
                        Write-OK "$($dir.Name): 包还原成功"
                    } else {
                        Write-Warn "$($dir.Name): 包还原可能有警告，继续..."
                    }

                    Write-Host "  正在编译..."
                    $buildOutput = & dotnet build $targetCsproj.Name --no-incremental 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-OK "$($dir.Name): 编译成功 (Build succeeded)"
                    } else {
                        Write-Fail "$($dir.Name): 编译失败"
                        $allOk = $false
                    }
                } finally {
                    Pop-Location
                }
            } else {
                Write-Fail "  跳过 $($dir.Name): .NET SDK 不可用"
            }
        } else {
            Write-Warn "  跳过 $($dir.Name): 未找到 .csproj 文件"
        }
    }
}

# 4. 最终报告
Write-Step "4/4 环境就绪报告"
if ($allOk) {
    Write-Host "`n  All checks passed! 可以开始使用 docx-format-applier 技能了！" -ForegroundColor $Green
} else {
    Write-Host "`n  部分依赖未就绪，请根据上面的提示完成安装后重新运行。" -ForegroundColor $Red
}

Write-Host "`n按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

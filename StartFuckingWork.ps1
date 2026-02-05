# 自动查找并启动飞书、Teams（包括Windows App版）和微信多开
# 调整顺序：先执行微信多开，避免窗口焦点干扰
# 执行完毕后2秒自动退出，并关闭微信资源管理器窗口

# 1. 查找微信路径 - 优先执行微信多开
$wechatPath = $null
$possibleWechatPaths = @(
    # 标准英文名路径
    "${env:ProgramFiles(x86)}\Tencent\WeChat\WeChat.exe",
    "${env:ProgramFiles}\Tencent\WeChat\WeChat.exe",
    
    # 中文拼音名路径
    "${env:ProgramFiles}\Tencent\Weixin\Weixin.exe",
    "${env:ProgramFiles(x86)}\Tencent\Weixin\Weixin.exe",
    
    # 其他可能路径
    "$env:LOCALAPPDATA\Programs\Tencent\WeChat\WeChat.exe",
    "$env:APPDATA\Tencent\WeChat\WeChat.exe",
    "$env:LOCALAPPDATA\Programs\Tencent\Weixin\Weixin.exe",
    "$env:APPDATA\Tencent\Weixin\Weixin.exe"
)

foreach ($path in $possibleWechatPaths) {
    if (Test-Path $path) {
        $wechatPath = $path
        Write-Host "找到微信路径: $path" -ForegroundColor Cyan
        break
    }
}

# 2. 微信多开功能 - 最先执行
$wechatExplorerProcess = $null
if ($wechatPath) {
    # 获取微信所在目录
    $wechatDir = Split-Path -Path $wechatPath -Parent
    $wechatExe = Split-Path -Path $wechatPath -Leaf
    
    # 打开微信所在目录（确保没有选中任何文件）
    $wechatExplorerProcess = Start-Process explorer.exe -ArgumentList $wechatDir -PassThru
    
    # 等待目录窗口打开
    Write-Host "等待微信目录打开..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    # 使用更可靠的方法模拟按键
    Write-Host "模拟按住回车键并双击微信图标..." -ForegroundColor Yellow
    
    # 方法1：使用SendKeys的替代语法
    try {
        # 加载必要的程序集
        Add-Type -AssemblyName System.Windows.Forms
        
        # 发送回车键按下事件（使用正确的语法）
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    }
    catch {
        Write-Host "SendKeys方法失败: $($_.Exception.Message)" -ForegroundColor Red
    }

    # 方法2：使用WScript.Shell发送按键（更可靠）
    try {
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.SendKeys("{ENTER}")
        Start-Sleep -Milliseconds 100
        $wshell.SendKeys("{ENTER}")
    }
    catch {
        Write-Host "WScript.Shell方法失败: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # 方法3：直接多次启动微信（最可靠）
    Write-Host "尝试直接启动两个微信实例..." -ForegroundColor Yellow
    Start-Process $wechatPath
    Start-Sleep -Seconds 2
    Start-Process $wechatPath
    
    Write-Host "微信多开操作已完成" -ForegroundColor Green
} else {
    Write-Host "未找到微信路径，请手动安装或检查" -ForegroundColor Yellow
    Write-Host "已知微信可能路径:"
    Write-Host "1. C:\Program Files\Tencent\Weixin\Weixin.exe" -ForegroundColor White
    Write-Host "2. C:\Program Files (x86)\Tencent\WeChat\WeChat.exe" -ForegroundColor White
}

# 3. 查找飞书路径 - 在微信之后执行
$feishuPath = $null
$possibleFeishuPaths = @(
    "$env:LOCALAPPDATA\Feishu\Feishu.exe",
    "$env:APPDATA\Feishu\Feishu.exe",
    "${env:ProgramFiles(x86)}\Feishu\Feishu.exe",
    "$env:USERPROFILE\AppData\Local\Feishu\Feishu.exe"
)

foreach ($path in $possibleFeishuPaths) {
    if (Test-Path $path) {
        $feishuPath = $path
        break
    }
}

# 4. 查找并启动Teams（支持桌面版和Windows App版）
$teamsStarted = $false

# 方法1：尝试查找桌面版Teams
$desktopTeamsPaths = @(
    "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
    "$env:APPDATA\Microsoft\Teams\current\Teams.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Teams\current\Teams.exe",
    "${env:ProgramFiles}\Microsoft\Teams\current\Teams.exe"
)

foreach ($path in $desktopTeamsPaths) {
    if (Test-Path $path) {
        Start-Process $path
        $teamsStarted = $true
        Write-Host "Teams (桌面版) 已启动" -ForegroundColor Green
        break
    }
}

# 方法2：如果桌面版没找到，尝试启动Windows App版
if (-not $teamsStarted) {
    try {
        # 使用Teams的URI协议启动
        Start-Process "msteams:"
        $teamsStarted = $true
        Write-Host "Teams (Windows App版) 已启动" -ForegroundColor Green
    }
    catch {
        Write-Host "无法通过URI协议启动Teams" -ForegroundColor Yellow
    }
}

# 方法3：如果URI协议失败，尝试通过AppxPackage启动
if (-not $teamsStarted) {
    try {
        # 获取Teams的AppxPackage
        $teamsPackage = Get-AppxPackage -Name MicrosoftTeams
        
        if ($teamsPackage) {
            # 构造启动命令
            $teamsAppId = $teamsPackage.PackageFamilyName + "!MSTeams"
            Start-Process "explorer.exe" -ArgumentList "shell:AppsFolder\$teamsAppId"
            $teamsStarted = $true
            Write-Host "Teams (Windows App版) 已启动" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "无法通过AppxPackage启动Teams" -ForegroundColor Yellow
    }
}

# 如果所有方法都失败
if (-not $teamsStarted) {
    Write-Host "未找到Teams路径，请手动安装或检查" -ForegroundColor Yellow
}

# 5. 启动飞书 - 最后执行
if ($feishuPath) {
    Start-Process $feishuPath
    Write-Host "飞书已启动" -ForegroundColor Green
} else {
    Write-Host "未找到飞书路径，请手动安装或检查" -ForegroundColor Yellow
}

# 6. 关闭微信资源管理器窗口（强力版）
if ($wechatPath -and $wechatExplorerProcess) {
    try {
        Write-Host "正在尝试关闭微信资源管理器窗口..." -ForegroundColor Yellow
        
        # 方法1：通过进程ID直接关闭
        try {
            Stop-Process -Id $wechatExplorerProcess.Id -Force -ErrorAction Stop
            Write-Host "已通过进程ID关闭微信资源管理器窗口" -ForegroundColor Green
        }
        catch {
            Write-Host "进程关闭失败，尝试其他方法: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # 方法2：通过窗口标题关闭
        $windowClosed = $false
        $shell = New-Object -ComObject Shell.Application
        $windows = $shell.Windows()
        
        foreach ($window in $windows) {
            try {
                $path = $window.Document.Folder.Self.Path
                if ($path -eq $wechatDir) {
                    $window.Quit()
                    $windowClosed = $true
                    Write-Host "已通过COM对象关闭微信资源管理器窗口" -ForegroundColor Green
                    break
                }
            }
            catch {
                # 忽略无法访问路径的窗口
            }
        }
        
        # 方法3：通过窗口句柄关闭
        if (-not $windowClosed) {
            try {
                Add-Type @"
                    using System;
                    using System.Runtime.InteropServices;
                    public class WindowCloser {
                        [DllImport("user32.dll", SetLastError = true)]
                        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
                        
                        [DllImport("user32.dll", CharSet = CharSet.Auto)]
                        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
                        
                        const UInt32 WM_CLOSE = 0x0010;
                        
                        public static void CloseWindow(string windowTitle) {
                            IntPtr hWnd = FindWindow(null, windowTitle);
                            if (hWnd != IntPtr.Zero) {
                                SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                            }
                        }
                    }
"@ -ErrorAction Stop
                
                $windowTitle = Split-Path $wechatDir -Leaf
                [WindowCloser]::CloseWindow($windowTitle)
                Write-Host "已尝试通过窗口句柄关闭微信资源管理器窗口" -ForegroundColor Yellow
            }
            catch {
                Write-Host "窗口句柄关闭方法失败: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # 方法4：强制关闭所有资源管理器窗口（谨慎使用）
        if (-not $windowClosed) {
            try {
                # 获取所有资源管理器进程
                $explorerProcesses = Get-Process -Name explorer -ErrorAction SilentlyContinue
                
                foreach ($proc in $explorerProcesses) {
                    try {
                        # 检查进程是否是我们打开的
                        if ($proc.StartTime -ge $wechatExplorerProcess.StartTime) {
                            Stop-Process -Id $proc.Id -Force
                            Write-Host "已强制关闭资源管理器进程 (ID: $($proc.Id))" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        Write-Host "无法关闭进程 (ID: $($proc.Id)): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            catch {
                Write-Host "强制关闭资源管理器失败: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "关闭微信资源管理器窗口失败: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# 显示最终状态
Write-Host "`n启动状态汇总:" -ForegroundColor Cyan
Write-Host "微信: $(if ($wechatPath) { '已尝试多开' } else { '未启动' })"
Write-Host "Teams: $(if ($teamsStarted) { '已启动' } else { '未启动' })"
Write-Host "飞书: $(if ($feishuPath) { '已启动' } else { '未启动' })"

# 等待2秒后自动退出
Write-Host "`n脚本将在2秒后自动退出..."
Start-Sleep -Seconds 2
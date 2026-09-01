<#
.SYNOPSIS
    Microsoft Excelのダイアログを検索し、Enterキーを送信します。

.DESCRIPTION
    表示中のトップレベル・ウィンドウを列挙し、次の条件をすべて満たす
    ダイアログを検索します。

    ・指定したタイトルと完全に一致する
    ・標準ダイアログのウィンドウクラス（#32770）である
    ・EXCEL.exeが所有している

    対象が見つかった場合は、最初の1件にEnterキーを送信して終了します。
    -WaitSecondsで指定した時間内に見つからない場合は、警告を表示して終了します。

.PARAMETER Title
    検索するダイアログのタイトルを指定します。
    既定値は「Microsoft Excel」です。大文字と小文字も区別して完全一致します。

.PARAMETER ListOnly
    Enterキーを送信せず、該当したダイアログの情報だけを表示します。

.PARAMETER WaitSeconds
    ダイアログが見つかるまで待機する最大秒数を指定します。
    既定値は0秒で、待機せず1回だけ検索します。10分間待機する場合は600を指定します。

.PARAMETER PollIntervalMilliseconds
    待機中にダイアログを再検索する間隔をミリ秒で指定します。
    既定値は1000ミリ秒（1秒）です。

.EXAMPLE
    .\Send-EnterToExcelDialog.ps1

    待機せずに1回検索し、見つかった最初のダイアログにEnterキーを送信します。

.EXAMPLE
    .\Send-EnterToExcelDialog.ps1 -ListOnly

    待機せずに検索し、Enterキーを送信せずに該当ダイアログの情報を表示します。

.EXAMPLE
    .\Send-EnterToExcelDialog.ps1 -WaitSeconds 600

    最大10分間、1秒間隔で検索します。対象が見つかり次第Enterキーを送信して終了します。

.EXAMPLE
    .\Send-EnterToExcelDialog.ps1 -WaitSeconds 600 -ListOnly

    最大10分間検索し、対象が見つかったらキーを送信せずに情報だけを表示します。
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $Title = 'Microsoft Excel',

    [Parameter()]
    [switch] $ListOnly,

    [Parameter()]
    [ValidateRange(0, 86400)]
    [int] $WaitSeconds = 0,

    [Parameter()]
    [ValidateRange(100, 60000)]
    [int] $PollIntervalMilliseconds = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not ('ExcelDialogInputV2.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelDialogInputV2
{
    public sealed class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
        public string ClassName { get; set; }
        public uint ProcessId { get; set; }
    }

    public static class NativeMethods
    {
        [return: MarshalAs(UnmanagedType.Bool)]
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "EnumWindows", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindowsNative(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);

        public static WindowInfo[] GetVisibleWindows()
        {
            var windows = new List<WindowInfo>();
            EnumWindowsProc callback = delegate(IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                int textLength = GetWindowTextLength(hWnd);
                if (textLength == 0)
                    return true;

                var text = new StringBuilder(textLength + 1);
                GetWindowText(hWnd, text, text.Capacity);

                var className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);

                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);

                windows.Add(new WindowInfo {
                    Handle = hWnd,
                    Title = text.ToString(),
                    ClassName = className.ToString(),
                    ProcessId = processId
                });
                return true;
            };

            // EnumWindowsは、コールバックによって列挙が停止された場合もfalseを返す。
            // その時点までに取得できたウィンドウを使用することで、GetWindowTextが
            // 残した古いエラー値をWindows PowerShell 5.1で誤検出することも防ぐ。
            EnumWindowsNative(callback, IntPtr.Zero);

            return windows.ToArray();
        }
    }
}
'@
}

function Find-ExcelDialog {
    $windows = foreach ($window in [ExcelDialogInputV2.NativeMethods]::GetVisibleWindows()) {
        $processName = $null
        try {
            $processName = (Get-Process -Id $window.ProcessId -ErrorAction Stop).ProcessName
        }
        catch {
            # 列挙後、処理するまでの間にウィンドウが閉じられる場合がある。
        }

        [pscustomobject]@{
            Handle      = $window.Handle
            HandleHex   = ('0x{0:X}' -f $window.Handle.ToInt64())
            Title       = $window.Title
            ClassName   = $window.ClassName
            ProcessId   = $window.ProcessId
            ProcessName = $processName
        }
    }

    @(
        $windows | Where-Object {
            $_.Title -ceq $Title -and
            $_.ClassName -ceq '#32770' -and
            $_.ProcessName -ieq 'EXCEL'
        }
    )
}

$stopwatch = [Diagnostics.Stopwatch]::StartNew()

do {
    $matches = @(Find-ExcelDialog)
    if ($matches.Count -gt 0) {
        break
    }

    if ($stopwatch.Elapsed.TotalSeconds -ge $WaitSeconds) {
        break
    }

    $remainingMilliseconds = [math]::Max(
        0,
        ($WaitSeconds * 1000) - $stopwatch.ElapsedMilliseconds
    )
    $sleepMilliseconds = [math]::Min(
        $PollIntervalMilliseconds,
        $remainingMilliseconds
    )

    if ($sleepMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $sleepMilliseconds
    }
} while ($true)

$stopwatch.Stop()

if ($matches.Count -eq 0) {
    if ($WaitSeconds -gt 0) {
        Write-Warning "No visible Excel dialog with the exact title '$Title' was found within $WaitSeconds seconds."
    }
    else {
        Write-Warning "No visible Excel dialog with the exact title '$Title' was found."
    }
    return
}

if ($ListOnly) {
    $matches | Select-Object HandleHex, Title, ClassName, ProcessId, ProcessName
    return
}

$WM_KEYDOWN = 0x0100
$WM_KEYUP   = 0x0101
$VK_RETURN  = [IntPtr]::new(0x0D)
$ENTER_KEYDOWN_LPARAM = [IntPtr] [int64] 0x001C0001
$ENTER_KEYUP_LPARAM   = [IntPtr] [int64] 0xC01C0001

$window = $matches[0]

# ExcelのUIメッセージキューへキー操作を登録し、ダイアログ側で
# キーボードから入力されたEnterキーと同様に処理できるようにする。
$downSent = [ExcelDialogInputV2.NativeMethods]::PostMessage(
    $window.Handle,
    $WM_KEYDOWN,
    $VK_RETURN,
    $ENTER_KEYDOWN_LPARAM
)
$upSent = [ExcelDialogInputV2.NativeMethods]::PostMessage(
    $window.Handle,
    $WM_KEYUP,
    $VK_RETURN,
    $ENTER_KEYUP_LPARAM
)

if (-not ($downSent -and $upSent)) {
    $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    throw "Failed to send Enter to $($window.HandleHex). Win32 error: $errorCode"
}

Write-Output ([pscustomobject]@{
    HandleHex = $window.HandleHex
    Title     = $window.Title
    ProcessId = $window.ProcessId
    EnterSent = $true
})

return

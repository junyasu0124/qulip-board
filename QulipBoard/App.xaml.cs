using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Windows.ApplicationModel.DataTransfer;
using Clipboard = Windows.ApplicationModel.DataTransfer.Clipboard;

namespace QulipBoard;

public partial class App : System.Windows.Application
{
  private readonly ContextMenuStrip menu = new();
  private readonly NotifyIcon notifyIcon = new();

  private static bool IsEnabled { get; set; } = true;
  private static bool IsCtrlPressed { get; set; } = false;
  private static bool IsAltPressed { get; set; } = false;

  private static bool FirstOrSecond { get; set; } = false;

  private static readonly KeyboardHook hook = new();

  private static Timer? aliveTimer = null;


  protected override void OnStartup(StartupEventArgs e)
  {
    try // Subscribe to startup
    {
        var regkey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
        if (regkey != null)
        {
            regkey.SetValue(System.Windows.Forms.Application.ProductName, System.Windows.Forms.Application.ExecutablePath);
            regkey.Close();
        }
    }
    catch { }


    IsEnabled = Settings.Default.IsEnabled;

    notifyIcon.Visible = true;
    menu.Items.Clear();

    if (IsEnabled)
    {
      menu.Items.Add("Inactivate", null, (obj, e) =>
      {
        IsEnabled = !IsEnabled;
        ChangeState();
      });
      notifyIcon.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "NotifyIcon.ico");
      hook.Hook();
    }
    else
    {
      menu.Items.Add("Activate", null, (obj, e) =>
      {
        IsEnabled = !IsEnabled;
        ChangeState();
      });
      notifyIcon.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "NotifyIconBlack.ico");
      hook.UnHook();
    }

    menu.Items.Add("Exit", null, (obj, e) => { Shutdown(); });

    notifyIcon.Text = "QulipBoard";
    notifyIcon.ContextMenuStrip = menu;

    base.OnStartup(e);

    aliveTimer = new() { Interval = 1000 * 60 * 15 };
    aliveTimer.Start();
    aliveTimer.Tick += (obj, e) =>
    {
      var temp = hook;
      temp.Responding = true;
    };
  }

  protected override void OnExit(ExitEventArgs e)
  {
    menu.Dispose();
    notifyIcon.Dispose();
    hook.UnHook();
    Settings.Default.IsEnabled = IsEnabled;

    base.OnExit(e);
  }

  private void ChangeState()
  {
    if (IsEnabled)
    {
      hook.Hook();
      menu.Items[0].Text = "Inactivate";
      notifyIcon.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "NotifyIcon.ico");
    }
    else
    {
      hook.UnHook();
      menu.Items[0].Text = "Activate";
      notifyIcon.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "NotifyIconBlack.ico");
    }
  }

  private class KeyboardHook
  {
    protected const int WH_KEYBOARD_LL = 0x000D;
    protected const int WM_KEYDOWN = 0x0100;
    protected const int WM_KEYUP = 0x0101;
    protected const int WM_SYSKEYDOWN = 0x0104;
    protected const int WM_SYSKEYUP = 0x0105;

    private const int VK_LControl = 0xA2;

    [StructLayout(LayoutKind.Sequential)]
    private class KBDLLHOOKSTRUCT
    {
      internal uint vkCode;
      internal uint scanCode;
      internal KBDLLHOOKSTRUCTFlags flags;
      internal uint time;
      internal UIntPtr dwExtraInfo;
    }

    [Flags]
    private enum KBDLLHOOKSTRUCTFlags : uint
    {
      KEYEVENTF_EXTENDEDKEY = 0x0001,
      KEYEVENTF_KEYUP = 0x0002,
      KEYEVENTF_SCANCODE = 0x0008,
      KEYEVENTF_UNICODE = 0x0004,
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, KeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern Int16 GetAsyncKeyState(int vKey);

    private delegate IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private KeyboardProc? proc;
    private IntPtr hookId = IntPtr.Zero;

    public bool Responding { get; set; } = true;

    internal void Hook()
    {
      if (hookId == IntPtr.Zero)
      {
        proc = HookProcedure;
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;
        if (curModule != null && curModule.ModuleName != null)
          hookId = SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
      }
    }

    internal void UnHook()
    {
      UnhookWindowsHookEx(hookId);
      hookId = IntPtr.Zero;
    }

    private IntPtr HookProcedure(int nCode, IntPtr wParam, IntPtr lParam)
    {
      if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
      {
        var kb = (KBDLLHOOKSTRUCT?)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
        var vkCode = (int?)kb?.vkCode;
        if (vkCode.HasValue)
          OnKeyDownEvent(vkCode.Value);
      }
      else if (nCode >= 0 && (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP))
      {
        var kb = (KBDLLHOOKSTRUCT?)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
        var vkCode = (int?)kb?.vkCode;
        if (vkCode.HasValue)
          OnKeyUpEvent(vkCode.Value);
      }

      return CallNextHookEx(hookId, nCode, wParam, lParam);
    }

    protected static async void OnKeyDownEvent(int keyCode)
    {
      if (keyCode == (int)Keys.LControlKey)
      {
        IsCtrlPressed = true;
      }
      else if (keyCode == (int)Keys.LMenu)
      {
        IsAltPressed = true;
        if ((GetAsyncKeyState(VK_LControl) & 0x8000) != 0)
          IsCtrlPressed = true;
      }
      else if (keyCode == (int)Keys.Space && IsCtrlPressed && IsAltPressed)
      {
        await EditClipBoard.ChangeOrderSecondAsync();
        IsCtrlPressed = false;
        IsAltPressed = false;
      }
    }
    protected static void OnKeyUpEvent(int keyCode)
    {
      if (keyCode == (int)Keys.LControlKey)
      {
        IsCtrlPressed = false;
      }
      else if (keyCode == (int)Keys.LMenu)
      {
        IsAltPressed = false;
      }
    }

  }

  private class EditClipBoard
  {
    /// <returns>
    /// 0: Success
    /// 1: Clipboard history is not enabled
    /// 2: Failed to get clipboard history
    /// 3: Second clipboard history is not exist
    /// </returns>
    internal static async Task<int> ChangeOrderSecondAsync()
    {
      await Task.Delay(5);
      if (Clipboard.IsHistoryEnabled() == false)
      {
        return 1;
      }
      var result = await Clipboard.GetHistoryItemsAsync();
      if (result.Status != ClipboardHistoryItemsResultStatus.Success)
      {
        return 2;
      }

      IReadOnlyList<ClipboardHistoryItem> historyList = result.Items;

      if (historyList.Count <= 0)
      {
        return 3;
      }
      if (FirstOrSecond == false)
      {
        FirstOrSecond = true;
        Clipboard.SetHistoryItemAsContent(historyList[1]);
      }
      else
      {
        FirstOrSecond = false;
        Clipboard.SetHistoryItemAsContent(historyList[0]);
      }
      return 0;
    }
  }
}

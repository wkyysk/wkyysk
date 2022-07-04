using System;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using InterceptionDevice = System.Int32;
using InterceptionPrecedence = System.Int32;
using InterceptionFilter = System.UInt32;
using InterceptionContext = System.IntPtr;
using static RazerTartarusPro.Module.Interception;

using RazerTartarusPro.Class;
using RazerTartarusPro.Module;

namespace RazerTartarusPro.Form
{
   ///  <summary>メインルーチン</summary>
   /// <remarks>バッググラウンド起動</remarks>
   public partial class Main : System.Windows.Forms.Form
   {
       // ===== 定数 =====
       ///  <summary>キーボードのハードウェアID</summary>
       private const string KEYBD_HID = @"HID\VID_1532&PID_0244&REV_0200&MI_01&Col01";
       ///  <summary>パッドのハードウェアID</summary>
       private const string KEYBD_PAD = @"HID\VID_1532&PID_0244&REV_0200&MI_00";
       ///  <summary>マウスのハードウェアID</summary>
       private const string MOUSE_HID = @"HID\VID_1532&PID_0244&REV_0200&MI_02";
       ///  <summary>デバイス情報</summary>
       private KeyDevice Devices = new KeyDevice();
       
       ///  <summary>XTコードリスト</summary>
       private readonly Dictionary<string, ushort> XT_CODE = new Dictionary<string, ushort> {
           { "ESC", 0x01 }, { "TAB", 0x0F }, { "SHIFT", 0x2A }, { "CTRL", 0x1D }, { "SPACE", 0x39 }, { "ALT", 0x38 }, { "BACKSPACE", 0x0E }, { "ENTER", 0x1C },
           { "1", 0x02 }, { "2", 0x03 }, { "3", 0x04 }, { "4", 0x05 }, { "5", 0x06 }, { "6", 0x07 }, { "7", 0x08 }, { "8", 0x09 }, { "9", 0x0A }, { "0", 0x0B },
           { "F1", 0x3B }, { "F2", 0x3C }, { "F3", 0x3D }, { "F4", 0x3E }, { "F5", 0x3F }, { "F6", 0x40 }, { "F7", 0x41 }, { "F8", 0x42 }, { "F9", 0x43 }, { "F10", 0x44 },
           { "CAPSLOCK", 0x3A }, { "NUMLOCK", 0x45 }, { "SCROLLLOCK", 0x46 },
           { "Q", 0x10 }, { "W", 0x11 }, { "E", 0x12 }, { "R", 0x13 }, { "T", 0x14 }, { "Y", 0x15 }, { "U", 0x16 }, { "I", 0x17 }, { "O", 0x18 }, { "P", 0x19 },
           { "A", 0x1E }, { "S", 0x1F }, { "D", 0x20 }, { "F", 0x21 }, { "G", 0x22 }, { "H", 0x23 }, { "J", 0x24 }, { "K", 0x25 }, { "L", 0x26 },
           { "Z", 0x2C }, { "X", 0x2D }, { "C", 0x2E }, { "V", 0x2F }, { "B", 0x30 }, { "N", 0x31 }, { "M", 0x32 }
       };
       
       // ===== メンバー変数 =====
       ///  <summary>メインルーチン</summary>
       private ThreadTimer MainTimer;
       ///  <summary>スレッド重複回避用</summary>
       private object LockObject = new object();
       
       ///  <summary>設定フォーム</summary>
       private Config frmConfig = new Config();
       
       ///  <summary>ScanCode.ini</summary>
       public static ScanCode _ScanCode { private set; get; }
       ///  <summary>KeyConfig.ini</summary>
       public static KeyConfig _KeyConfig { private set; get; }
       
       ///  <summary>デバイス情報</summary>
       private class KeyDevice
       {
           public InterceptionContext Context;
           public InterceptionDevice Keyboard;
           public InterceptionDevice Pad;
           public InterceptionDevice Mouse;
           public InterceptionStroke Stroke;
           
           public KeyDevice() : this(Interception_create_context()) { }
           
           public KeyDevice(InterceptionContext _context)
           {
               Context = _context;
               Keyboard = MAX_DEVICE;
               Pad = MAX_DEVICE;
               Mouse = MAX_DEVICE;
               Stroke = new InterceptionStroke();
               if (Context == null) return;
               Interception_set_filter(Context, Interception_is_keyboard, (ushort)(KeyStateFilter.KeyDown | KeyStateFilter.KeyUp | KeyStateFilter.KeyE0 | KeyStateFilter.KeyE1));
               Interception_set_filter(Context, Interception_is_mouse, (ushort)(MouseStateFilter.MouseWheel | MouseStateFilter.MiddleButtonDown | MouseStateFilter.MouseWheel));
               
               // キーボードをハードウェアIDで検索する
               for (int i = 0; i < MAX_KEYBOARD; i++)
               {
                   InterceptionDevice d = GetKeyBoardNum(i);
                   StringBuilder sb = new StringBuilder(500);
                   
                   // ハードウェアIDが取得した場合
                   if (Interception_get_hardware_id(Context, d, sb, (uint)sb.Capacity) > 0)
                   {
                       switch (sb.ToString())
                       {
                           case KEYBD_HID:
                               Keyboard = d;
                               break;
                           case KEYBD_PAD:
                               Pad = d;
                               break;
                       }
                   }
               }
               
               // マウスをハードウェアIDで検索する
               for (int i = 0; i < MAX_MOUSE; i++)
               {
                   InterceptionDevice d = GetMouseNum(i);
                   StringBuilder sb = new StringBuilder(500);
                  
                   // ハードウェアIDを取得した、かつ対象のIDである場合
                   if (Interception_get_hardware_id(Context, d, sb, (uint)sb.Capacity) > 0 && (sb.ToString() == MOUSE_HID))
                   {
                       Mouse = d;
                       break;
                   }
               }
               
               // キーボード, パッド, マウスのいずれかが取得できなかった場合、終了
               if (Keyboard == MAX_DEVICE || Pad == MAX_DEVICE || Mouse == MAX_DEVICE)
               {
                   return;
               }
           }
       }
       
       /// <summary>
       ///  <summary>コンストラクタ</summary>
       /// </summary>
       public Main()
       {
           InitializeComponent();
           
           _ScanCode = new ScanCode();
           _KeyConfig = new KeyConfig();
           
           MainTimer = new ThreadTimer(1, MainThread);
           MainTimer.StartTimer();
       }
       
       // ===== イベント =====
       /// <summary>
       ///  <summary>設定ボタンクリックイベント</summary>
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
       private void ToolStripMenuItemConfig_Click(object sender, EventArgs e)
       {
           try
           {
               if (frmConfig.Visible) frmConfig.Activate();
               else frmConfig.Show();
           }
           catch
           {
           }
       }
       
       /// <summary>
       ///  <summary>終了ボタンクリックイベント</summary>
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
       private void ToolStripMenuItemExit_Click(object sender, EventArgs e)
       {
           try
           {
               MainTimer.StopTimer();
               frmConfig.Dispose();
           }
           catch
           {
           }
           finally
           {
               Application.Exit();
           }
       }
       
       // ===== メンバー関数 =====
       ///  <summary>キーボードデバイスの入力送信</summary>
       /// <param name="context">デバイス環境</param>
       /// <param name="device">デバイス</param>
       /// <param name="stroke">入力情報</param>
       /// <param name="keyStroke">キーボード入力情報</param>
       /// <param name="nstroke">入力種別</param>
       private static void SendKey(InterceptionContext context, InterceptionDevice device, InterceptionStroke stroke, InterceptionKeyStroke keyStroke, InterceptionFilter nstroke)
       {
           stroke.Key = keyStroke;
           Interception_send(context, device, ref stroke, nstroke);
       }
       
       ///  <summary>マウスデバイスの入力送信</summary>
       /// <param name="context">デバイス環境</param>
       /// <param name="device">デバイス</param>
       /// <param name="stroke">入力情報</param>
       /// <param name="mouseStroke">マウス入力情報</param>
       /// <param name="nstroke">入力種別</param>
       private static void SendKey(InterceptionContext context, InterceptionDevice device, InterceptionStroke stroke, InterceptionMouseStroke mouseStroke, InterceptionFilter nstroke)
       {
           stroke.Mouse = mouseStroke;
           Interception_send(context, device, ref stroke, nstroke);
       }
       
       ///  <summary>スルーするデバイスの入力送信</summary>
       /// <param name="context">デバイス環境</param>
       /// <param name="device">デバイス</param>
       /// <param name="stroke">入力情報</param>
       /// <param name="nstroke">入力種別</param>
       private static void SendKey(InterceptionContext context, InterceptionDevice device, InterceptionStroke stroke, InterceptionFilter nstroke)
       {
           Interception_send(context, device, ref stroke, nstroke);
       }
       
       ///  <summary>(ThreadTimer)メインスレッド</summary>
       /// <param name="o"></param>
       private void MainThread(object o)
       {
           int iErrorCode = -1;
           
           try
           {
               lock (LockObject)
               {
                   // キー入力を検出した場合
                   InterceptionDevice InputDevise;
                   
                   if (Interception_receive(Devices.Context, InputDevise = Interception_wait(Devices.Context), ref Devices.Stroke, 1) > 0)
                   {
                       // RazerTartarusProのキーボード入力を検知した場合
                       if (Devices.Keyboard == InputDevise)
                       {
                           InterceptionKeyStroke s = Devices.Stroke.Key;
                           bool isKeySend = true;
                           
                           // ScanCodeから入力キーを判別する
                           ScanCode.KEY key_code = _ScanCode.Key(s.Code);
                           string key_name = Config._KeyConfig.Value((KeyConfig.KEY)(int)key_code);
                           s.Code = (ushort)XT_CODE[key_name];
                           
                           if (isKeySend)
                               SendKey(Devices.Context, InputDevise, Devices.Stroke, s, 1);
                       }
                       // RazerTartarusProのパッド入力を検知した場合
                       else if (Devices.Pad == InputDevise)
                       {
                           InterceptionKeyStroke s = Devices.Stroke.Key;
                           bool isKeySend = true;
                           
                           // ScanCodeから入力キーを判別する
                           ScanCode.KEY key_code = _ScanCode.Key(s.Code);
                           string key_name = Config._KeyConfig.Value((KeyConfig.KEY)(int)key_code);
                           s.Code = (ushort)XT_CODE[key_name];
                           
                           if (isKeySend)
                               SendKey(Devices.Context, InputDevise, Devices.Stroke, s, 1);
                       }
                       // RazerTartarusProのマウス入力を検知した場合
                       else if (Devices.Mouse == InputDevise)
                       {
                           InterceptionMouseStroke s = Devices.Stroke.Mouse;
                           bool isKeySend = true;
                           
                           // ScanCodeから入力キーを判別する
                           ScanCode.KEY key_code = _ScanCode.Key(s.Flags);
                           string key_name = Config._KeyConfig.Value((KeyConfig.KEY)(int)key_code);
                           s.Flags = (ushort)XT_CODE[key_name];
                           
                           if (isKeySend)
                               SendKey(Devices.Context, InputDevise, Devices.Stroke, s, 1);
                       }
                       // 他デバイスはスルーする
                       else
                           SendKey(Devices.Context, InputDevise, Devices.Stroke, 1);
                   }
                   else
                   {
                   
                   }
               }
               Thread.Sleep(0);
           }
           catch
           {
               Interception_destroy_context(Devices.Context);
               MainTimer.StopTimer();
           }
       }
   }
}

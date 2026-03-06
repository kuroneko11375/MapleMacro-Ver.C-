using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using WinMessageBox = System.Windows.Forms.MessageBox;
using WinMBBtn = System.Windows.Forms.MessageBoxButtons;
using WinMBIcon = System.Windows.Forms.MessageBoxIcon;
using WinDR = System.Windows.Forms.DialogResult;
using WinKeys = System.Windows.Forms.Keys;
using WinOFD = System.Windows.Forms.OpenFileDialog;

namespace MapleStoryMacro
{
    using WCtl = System.Windows.Controls;
    using WMedia = System.Windows.Media;

    public partial class MainWindow
    {
        // ═══════════════════════════════════════════════
        //  WPF Dialog Factory Helpers (Cyber Theme)
        // ═══════════════════════════════════════════════

        private static readonly WMedia.FontFamily _cyFont = new("Microsoft YaHei UI");
        private static readonly WMedia.FontFamily _monoFont = new("Consolas");

        private static WMedia.SolidColorBrush _B(byte r, byte g, byte b)
            => new(WMedia.Color.FromRgb(r, g, b));

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private static System.Windows.Media.Imaging.BitmapSource _BmpToWpf(Bitmap bmp)
        {
            var h = bmp.GetHbitmap();
            try
            {
                return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    h, IntPtr.Zero, Int32Rect.Empty,
                    System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
            }
            finally { DeleteObject(h); }
        }

        private Window _CyWin(string title, double w, double h, bool resizable = false)
        {
            var win = new Window
            {
                Title = title, Width = w, Height = h,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = resizable ? ResizeMode.CanResize : ResizeMode.NoResize,
                Background = _B(45, 45, 48),
                WindowStyle = WindowStyle.ToolWindow
            };
            try { win.Owner = System.Windows.Application.Current.MainWindow; } catch { }
            return win;
        }

        private static WCtl.TextBlock _L(string text, double size = 12, WMedia.Brush? fg = null)
            => new() { Text = text, Foreground = fg ?? WMedia.Brushes.White, FontFamily = _cyFont, FontSize = size };

        private static WCtl.Button _Btn(string text, byte r, byte g, byte b, double w = 80, double h = 32)
        {
            var btn = new WCtl.Button
            {
                Content = text, Width = w, Height = h,
                Background = _B(r, g, b)
            };
            if (System.Windows.Application.Current.TryFindResource("CyberButton") is System.Windows.Style s)
                btn.Style = s;
            return btn;
        }

        private static WCtl.TextBox _Txt(double w = 200, double h = 28, string text = "", bool ro = false)
        {
            var t = new WCtl.TextBox
            {
                Width = w, Height = h, Text = text, IsReadOnly = ro
            };
            if (System.Windows.Application.Current.TryFindResource("CyberTextBox") is System.Windows.Style s)
                t.Style = s;
            return t;
        }

        private static WCtl.CheckBox _Chk(string text, bool chk = false, WMedia.Brush? fg = null)
            => new() { Content = text, IsChecked = chk, Foreground = fg ?? WMedia.Brushes.White, FontFamily = _cyFont, FontSize = 12 };

        private static WCtl.TextBox _Num(double w, double val, double min = 0, double max = 9999, int dec = 0)
        {
            var t = _Txt(w, 25, dec > 0 ? val.ToString($"F{dec}") : ((int)val).ToString());
            t.TextAlignment = System.Windows.TextAlignment.Center;
            var def = val;
            t.PreviewTextInput += (s, e) =>
            {
                string newText = t.Text.Remove(t.SelectionStart, t.SelectionLength).Insert(t.CaretIndex, e.Text);
                e.Handled = !double.TryParse(newText, out _);
            };
            t.LostFocus += (s, e) =>
            {
                if (double.TryParse(t.Text, out double v))
                {
                    v = Math.Clamp(v, min, max);
                    t.Text = dec > 0 ? v.ToString($"F{dec}") : ((int)v).ToString();
                }
                else t.Text = dec > 0 ? def.ToString($"F{dec}") : ((int)def).ToString();
            };
            return t;
        }

        private static double _NV(WCtl.TextBox tb) => double.TryParse(tb.Text, out double v) ? v : 0;

        private static WCtl.Border _Sep()
            => new() { Height = 2, Margin = new Thickness(0, 8, 0, 8), Background = _B(60, 60, 65) };

        private static WCtl.GroupBox _Grp(string header)
        {
            var g = new WCtl.GroupBox
            {
                Header = header,
                Margin = new Thickness(0, 4, 0, 0)
            };
            if (System.Windows.Application.Current.TryFindResource("CyberGroupBox") is System.Windows.Style s)
                g.Style = s;
            return g;
        }

        private static WCtl.StackPanel _Row(params UIElement[] items)
        {
            var sp = new WCtl.StackPanel { Orientation = WCtl.Orientation.Horizontal };
            foreach (var item in items) sp.Children.Add(item);
            return sp;
        }

        // ═══════════════════════════════════════════════
        //  ⚙ 熱鍵設定 (WPF)
        // ═══════════════════════════════════════════════

        private void OpenHotkeySettings()
        {
            var win = _CyWin("⚙ 熱鍵設定", 450, 380);
            var root = new WCtl.StackPanel { Margin = new Thickness(20) };
            var boxes = new Dictionary<string, WCtl.TextBox>();

            foreach (var (label, key) in new[] {
                ("播放熱鍵：", playHotkey), ("停止熱鍵：", stopHotkey),
                ("暫停熱鍵：", pauseHotkey), ("錄製熱鍵：", recordHotkey) })
            {
                var lbl = _L(label, 13); lbl.Width = 100; lbl.VerticalAlignment = VerticalAlignment.Center;
                var txt = _Txt(200, 28, GetKeyDisplayName(key), true); txt.Tag = key;
                txt.PreviewKeyDown += (s, e) =>
                {
                    e.Handled = true;
                    var k = e.Key == Key.System ? e.SystemKey : e.Key;
                    var fk = (Keys)KeyInterop.VirtualKeyFromKey(k);
                    txt.Text = GetKeyDisplayName(fk); txt.Tag = fk;
                };
                var row = _Row(lbl, txt); row.Margin = new Thickness(0, 0, 0, 10);
                root.Children.Add(row);
                boxes[label] = txt;
            }

            var chk = _Chk("啟用全局熱鍵", hotkeyEnabled);
            chk.Margin = new Thickness(0, 5, 0, 0);
            root.Children.Add(chk);

            var hint = _L("提示：點擊文字框後按下想要的按鍵", 11, _B(128, 128, 128));
            hint.Margin = new Thickness(0, 12, 0, 0);
            root.Children.Add(hint);

            var save = _Btn("儲存", 0, 122, 204);
            var cancel = _Btn("取消", 80, 80, 85);
            cancel.Margin = new Thickness(10, 0, 0, 0); cancel.BorderBrush = _B(100, 100, 105);

            save.Click += (s, e) =>
            {
                playHotkey = (Keys)boxes["播放熱鍵："].Tag;
                stopHotkey = (Keys)boxes["停止熱鍵："].Tag;
                pauseHotkey = (Keys)boxes["暫停熱鍵："].Tag;
                recordHotkey = (Keys)boxes["錄製熱鍵："].Tag;
                hotkeyEnabled = chk.IsChecked == true;
                AddLog($"設定已儲存：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}, 暫停={GetKeyDisplayName(pauseHotkey)}, 錄製={GetKeyDisplayName(recordHotkey)}");
                SaveAppSettings();
                win.Close();
            };
            cancel.Click += (s, e) => win.Close();

            var bp = _Row(save, cancel);
            bp.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            bp.Margin = new Thickness(0, 18, 0, 0);
            root.Children.Add(bp);

            win.Content = root;
            win.ShowDialog();
        }

        // ═══════════════════════════════════════════════
        //  ⚡ 自定義按鍵設定 (WPF)
        // ═══════════════════════════════════════════════

        private void OpenCustomKeySettings()
        {
            var win = _CyWin("⚡ 自定義按鍵設定 (15 格)", 750, 620, true);
            var root = new WCtl.StackPanel { Margin = new Thickness(10) };
            root.Children.Add(_L("設定最多 15 個自定義按鍵，在腳本播放時按間隔自動施放 | 按鍵欄位點擊後按下按鍵來設定", 10, _B(192, 192, 192)));

            // Header row
            var headerGrid = new WCtl.Grid { Margin = new Thickness(0, 8, 0, 2) };
            var colW = new[] { 35.0, 40, 120, 70, 70, 40, 70, 70 };
            var colH = new[] { "#", "啟用", "按鍵 (選中後按鍵)", "間隔(秒)", "開始(秒)", "暫停", "暫停(秒)", "延遲(秒)" };
            for (int c = 0; c < 8; c++)
            {
                headerGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(colW[c]) });
                var hl = _L(colH[c], 10, _B(200, 200, 200));
                hl.FontWeight = FontWeights.Bold; hl.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                WCtl.Grid.SetColumn(hl, c); headerGrid.Children.Add(hl);
            }
            root.Children.Add(headerGrid);

            // Data rows
            var sv = new WCtl.ScrollViewer { Height = 420, VerticalScrollBarVisibility = WCtl.ScrollBarVisibility.Auto };
            var rowPanel = new WCtl.StackPanel();
            var rowData = new List<(WCtl.CheckBox en, WCtl.TextBox key, WCtl.TextBox intv, WCtl.TextBox start, WCtl.CheckBox pe, WCtl.TextBox ps, WCtl.TextBox dl)>();

            for (int i = 0; i < 15; i++)
            {
                var slot = customKeySlots[i];
                var rowG = new WCtl.Grid { Margin = new Thickness(0, 1, 0, 1) };
                rowG.Background = i % 2 == 0 ? _B(40, 40, 45) : _B(35, 35, 40);
                for (int c = 0; c < 8; c++)
                    rowG.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(colW[c]) });

                var numLbl = _L($"#{i + 1}", 10, _B(0, 255, 255));
                numLbl.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                numLbl.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(numLbl, 0); rowG.Children.Add(numLbl);

                var chkEn = _Chk("", slot.Enabled);
                chkEn.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                chkEn.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(chkEn, 1); rowG.Children.Add(chkEn);

                string keyDisp = slot.KeyCode == Keys.None ? "(點擊設定)" : GetKeyDisplayName(slot.KeyCode);
                var txtKey = _Txt(110, 24, keyDisp, true); txtKey.Tag = slot.KeyCode;
                txtKey.Foreground = _B(144, 238, 144); txtKey.Background = _B(50, 60, 70);
                txtKey.TextAlignment = System.Windows.TextAlignment.Center;
                txtKey.PreviewKeyDown += (s, e) =>
                {
                    e.Handled = true;
                    var k = e.Key == Key.System ? e.SystemKey : e.Key;
                    var fk = (Keys)KeyInterop.VirtualKeyFromKey(k);
                    txtKey.Text = GetKeyDisplayName(fk); txtKey.Tag = fk;
                };
                txtKey.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(txtKey, 2); rowG.Children.Add(txtKey);

                var tI = _Num(60, slot.IntervalSeconds, 1, 9999); tI.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(tI, 3); rowG.Children.Add(tI);
                var tS = _Num(60, slot.StartAtSecond, 0, 9999); tS.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(tS, 4); rowG.Children.Add(tS);

                var chkP = _Chk("", slot.PauseScriptEnabled);
                chkP.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                chkP.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(chkP, 5); rowG.Children.Add(chkP);

                var tPS = _Num(60, slot.PauseScriptSeconds, 0, 999, 1);
                tPS.Foreground = _B(255, 255, 0); tPS.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(tPS, 6); rowG.Children.Add(tPS);
                var tDL = _Num(60, slot.PreDelaySeconds, 0, 999, 1);
                tDL.Foreground = _B(0, 255, 255); tDL.VerticalAlignment = VerticalAlignment.Center;
                WCtl.Grid.SetColumn(tDL, 7); rowG.Children.Add(tDL);

                rowPanel.Children.Add(rowG);
                rowData.Add((chkEn, txtKey, tI, tS, chkP, tPS, tDL));
            }
            sv.Content = rowPanel;
            root.Children.Add(sv);

            var hintLbl = _L("【說明】間隔: 每隔幾秒觸發 | 開始: 腳本播放幾秒後開始 | 暫停: 觸發前暫停腳本 | 延遲: 按鍵後等待\n【執行順序】腳本暫停 → 按下按鍵 → 延遲等待 → 繼續腳本", 9.5, _B(192, 192, 192));
            hintLbl.Margin = new Thickness(0, 6, 0, 0); hintLbl.TextWrapping = TextWrapping.Wrap;
            root.Children.Add(hintLbl);

            var bClear = _Btn("全部清除", 150, 60, 60, 100, 35);
            var bSave = _Btn("儲存", 0, 122, 204, 100, 35); bSave.Margin = new Thickness(10, 0, 0, 0);
            var bCancel = _Btn("取消", 80, 80, 85, 100, 35); bCancel.Margin = new Thickness(10, 0, 0, 0);
            bCancel.BorderBrush = _B(100, 100, 105);

            bClear.Click += (s, e) =>
            {
                for (int i = 0; i < 15; i++)
                {
                    var (en, key, intv, start, pe, ps, dl) = rowData[i];
                    en.IsChecked = false; key.Text = "(點擊設定)"; key.Tag = Keys.None;
                    intv.Text = "30"; start.Text = "0"; pe.IsChecked = false; ps.Text = "0.0"; dl.Text = "0.0";
                }
            };
            bSave.Click += (s, e) =>
            {
                try
                {
                    for (int i = 0; i < 15; i++)
                    {
                        var (en, key, intv, start, pe, ps, dl) = rowData[i];
                        customKeySlots[i].Enabled = en.IsChecked == true;
                        customKeySlots[i].KeyCode = key.Tag is Keys k ? k : Keys.None;
                        customKeySlots[i].Modifiers = Keys.None;
                        customKeySlots[i].IntervalSeconds = _NV(intv);
                        customKeySlots[i].StartAtSecond = _NV(start);
                        customKeySlots[i].PauseScriptEnabled = pe.IsChecked == true;
                        customKeySlots[i].PauseScriptSeconds = _NV(ps);
                        customKeySlots[i].PreDelaySeconds = _NV(dl);
                    }
                    int cnt = customKeySlots.Count(sl => sl.Enabled && sl.KeyCode != Keys.None);
                    AddLog($"✅ 自定義按鍵設定已更新：{cnt} 個已啟用");
                    AddLog("💡 提示：自定義按鍵會隨腳本一起保存");
                    win.Close();
                }
                catch (Exception ex) { WinMessageBox.Show($"儲存失敗: {ex.Message}", "錯誤", WinMBBtn.OK, WinMBIcon.Error); }
            };
            bCancel.Click += (s, e) => win.Close();

            var bp = _Row(bClear, bSave, bCancel);
            bp.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            bp.Margin = new Thickness(0, 10, 0, 0);
            root.Children.Add(bp);

            win.Content = root;
            win.ShowDialog();
        }

        // ═══════════════════════════════════════════════
        //  ⏰ 排程管理 (WPF)
        // ═══════════════════════════════════════════════

        private void OpenSchedulerSettings()
        {
            var win = _CyWin("⏰ 排程管理", 660, 830);
            var rootSv = new WCtl.ScrollViewer { VerticalScrollBarVisibility = WCtl.ScrollBarVisibility.Auto };
            var main = new WCtl.StackPanel { Margin = new Thickness(20, 15, 20, 15) };

            main.Children.Add(_L("設定排程任務，可指定腳本、開始/結束時間", 11, _B(192, 192, 192)));

            var h1 = _L("📋 新增排程", 11, _B(0, 255, 255));
            h1.FontWeight = FontWeights.Bold; h1.Margin = new Thickness(0, 8, 0, 4);
            main.Children.Add(h1);

            // Script
            var lblScript = _L("腳本："); lblScript.Width = 50; lblScript.VerticalAlignment = VerticalAlignment.Center;
            var txtScript = _Txt(360, 25, currentScriptPath ?? "(使用當前已載入的腳本)", true);
            var btnBrowse = _Btn("...", 80, 80, 85, 35, 25); btnBrowse.FontSize = 10; btnBrowse.Margin = new Thickness(4, 0, 0, 0);
            var btnUseCurrent = _Btn("當前", 0, 100, 160, 50, 25); btnUseCurrent.FontSize = 10; btnUseCurrent.Margin = new Thickness(4, 0, 0, 0);
            main.Children.Add(_Row(lblScript, txtScript, btnBrowse, btnUseCurrent));

            btnBrowse.Click += (s, e) =>
            {
                var ofd = new WinOFD { Filter = "Maple 腳本|*.mscript|舊版 JSON 腳本|*.json|所有檔案|*.*", Title = "選擇排程腳本" };
                if (ofd.ShowDialog() == WinDR.OK) { txtScript.Text = ofd.FileName; txtScript.Tag = ofd.FileName; }
            };
            btnUseCurrent.Click += (s, e) =>
            {
                if (recordedEvents.Count > 0) { txtScript.Text = currentScriptPath ?? "(使用當前已載入的腳本)"; txtScript.Tag = currentScriptPath; }
                else WinMessageBox.Show("當前沒有已載入的腳本！", "提示", WinMBBtn.OK, WinMBIcon.Warning);
            };

            // Start time
            var lblStart = _L("開始："); lblStart.Width = 50; lblStart.VerticalAlignment = VerticalAlignment.Center;
            var dpStart = new WCtl.DatePicker { SelectedDate = DateTime.Now.AddMinutes(1), Width = 130, Height = 25 };
            var txtStartTime = _Txt(80, 25, DateTime.Now.AddMinutes(1).ToString("HH:mm:ss"));
            txtStartTime.TextAlignment = System.Windows.TextAlignment.Center; txtStartTime.Margin = new Thickness(4, 0, 0, 0);
            var startRow = _Row(lblStart, dpStart, txtStartTime); startRow.Margin = new Thickness(0, 4, 0, 0);
            main.Children.Add(startRow);

            // End time
            var lblEnd = _L("結束："); lblEnd.Width = 50; lblEnd.VerticalAlignment = VerticalAlignment.Center;
            var chkEndTime = _Chk("", true); chkEndTime.VerticalAlignment = VerticalAlignment.Center; chkEndTime.Margin = new Thickness(0, 0, 4, 0);
            var dpEnd = new WCtl.DatePicker { SelectedDate = DateTime.Now.AddHours(1), Width = 130, Height = 25 };
            var txtEndTime = _Txt(80, 25, DateTime.Now.AddHours(1).ToString("HH:mm:ss"));
            txtEndTime.TextAlignment = System.Windows.TextAlignment.Center; txtEndTime.Margin = new Thickness(4, 0, 0, 0);
            chkEndTime.Checked += (s, e) => { dpEnd.IsEnabled = true; txtEndTime.IsEnabled = true; };
            chkEndTime.Unchecked += (s, e) => { dpEnd.IsEnabled = false; txtEndTime.IsEnabled = false; };
            var endRow = _Row(lblEnd, chkEndTime, dpEnd, txtEndTime); endRow.Margin = new Thickness(0, 2, 0, 0);
            main.Children.Add(endRow);

            // Loop
            var lblLoop = _L("循環："); lblLoop.Width = 50; lblLoop.VerticalAlignment = VerticalAlignment.Center;
            var numLoop = _Num(80, LoopCountValue, 1, 9999);
            var loopHint = _L("勾選「結束」可設定自動停止時間", 9, _B(128, 128, 128));
            loopHint.VerticalAlignment = VerticalAlignment.Center; loopHint.Margin = new Thickness(8, 0, 0, 0);
            var loopRow = _Row(lblLoop, numLoop, loopHint); loopRow.Margin = new Thickness(0, 4, 0, 0);
            main.Children.Add(loopRow);

            // Return to town
            var chkReturn = _Chk("🏠 回程（結束時自動執行回城序列）", true, _B(100, 220, 160));
            chkReturn.FontWeight = FontWeights.Bold; chkReturn.Margin = new Thickness(0, 8, 0, 0);
            main.Children.Add(chkReturn);

            var numCX = _Num(65, 652, 0, 3840); numCX.Foreground = _B(144, 238, 144);
            var numCY = _Num(65, 882, 0, 2160); numCY.Foreground = _B(144, 238, 144);
            var lCX = _L("點擊 X:"); lCX.VerticalAlignment = VerticalAlignment.Center;
            var lCY = _L("Y:"); lCY.VerticalAlignment = VerticalAlignment.Center; lCY.Margin = new Thickness(8, 0, 0, 0);
            numCX.Margin = new Thickness(4, 0, 0, 0); numCY.Margin = new Thickness(4, 0, 0, 0);
            var coordRow = _Row(lCX, numCX, lCY, numCY); coordRow.Margin = new Thickness(0, 4, 0, 0);
            main.Children.Add(coordRow);

            // Key capture helper
            Func<string, Keys, WMedia.SolidColorBrush, (WCtl.TextBox txt, WCtl.StackPanel row)> makeKeyCapture = (label, defKey, color) =>
            {
                var lbl = _L(label); lbl.Foreground = color; lbl.VerticalAlignment = VerticalAlignment.Center;
                var txt = _Txt(60, 25, defKey == Keys.None ? "(無)" : GetKeyDisplayName(defKey), true);
                txt.TextAlignment = System.Windows.TextAlignment.Center; txt.Foreground = color; txt.Tag = defKey;
                txt.PreviewKeyDown += (s, e) =>
                {
                    e.Handled = true;
                    var k = e.Key == Key.System ? e.SystemKey : e.Key;
                    var fk = (Keys)KeyInterop.VirtualKeyFromKey(k);
                    txt.Text = GetKeyDisplayName(fk); txt.Tag = fk;
                };
                txt.MouseRightButtonDown += (s, e) => { txt.Tag = Keys.None; txt.Text = "(無)"; };
                txt.Margin = new Thickness(4, 0, 0, 0);
                return (txt, _Row(lbl, txt));
            };

            var (txtPreKey, preKeyRow) = makeKeyCapture("前置鍵:", Keys.None, _B(255, 200, 100));
            var (txtConfirmKey, confirmKeyRow) = makeKeyCapture("確認鍵:", Keys.Enter, _B(130, 200, 255));
            var (txtSitKey, sitKeyRow) = makeKeyCapture("坐下鍵:", Keys.None, _B(0, 255, 255));
            var numSitDelay = _Num(55, 3, 0, 30, 1);
            var lSD = _L("延遲:"); lSD.VerticalAlignment = VerticalAlignment.Center;
            var lSec = _L("秒"); lSec.VerticalAlignment = VerticalAlignment.Center; lSec.Margin = new Thickness(2, 0, 0, 0);
            numSitDelay.Margin = new Thickness(4, 0, 0, 0);

            preKeyRow.Margin = new Thickness(0, 0, 12, 0);
            confirmKeyRow.Margin = new Thickness(0, 0, 12, 0);
            sitKeyRow.Margin = new Thickness(0, 0, 12, 0);
            var keyWrap = new WCtl.WrapPanel { Margin = new Thickness(0, 4, 0, 0) };
            keyWrap.Children.Add(preKeyRow); keyWrap.Children.Add(confirmKeyRow);
            keyWrap.Children.Add(sitKeyRow); keyWrap.Children.Add(_Row(lSD, numSitDelay, lSec));
            main.Children.Add(keyWrap);

            var keyHint = _L("右鍵清除 | 序列：停止 → 冷卻2s → [前置鍵] → 點擊(X,Y) → 偵測 → [確認鍵] → 延遲 → [坐下鍵]", 8.5, _B(128, 128, 128));
            keyHint.Margin = new Thickness(0, 2, 0, 0); keyHint.TextWrapping = TextWrapping.Wrap;
            main.Children.Add(keyHint);

            chkReturn.Checked += (s, e) => { numCX.IsEnabled = numCY.IsEnabled = txtPreKey.IsEnabled = txtConfirmKey.IsEnabled = txtSitKey.IsEnabled = numSitDelay.IsEnabled = true; };
            chkReturn.Unchecked += (s, e) => { numCX.IsEnabled = numCY.IsEnabled = txtPreKey.IsEnabled = txtConfirmKey.IsEnabled = txtSitKey.IsEnabled = numSitDelay.IsEnabled = false; };

            // Dialog detection
            var chkDialog = _Chk("🔍 對話框偵測（點擊後截圖比對，偵測到才按確認鍵）", false, _B(180, 180, 255));
            chkDialog.FontWeight = FontWeights.Bold; chkDialog.Margin = new Thickness(0, 8, 0, 0);
            main.Children.Add(chkDialog);

            var imgTemplate = new WCtl.Image { Width = 80, Height = 50, Stretch = WMedia.Stretch.Uniform };
            var imgBorder = new WCtl.Border { Child = imgTemplate, Width = 80, Height = 50, BorderBrush = _B(60, 60, 65), BorderThickness = new Thickness(1), Background = _B(30, 30, 35) };
            var lblTmplStatus = _L("(未設定模板)", 9, _B(128, 128, 128)); lblTmplStatus.VerticalAlignment = VerticalAlignment.Center; lblTmplStatus.Margin = new Thickness(4, 0, 0, 0);
            var btnCaptureTmpl = _Btn("📷 擷取模板", 80, 80, 140, 100, 26); btnCaptureTmpl.FontSize = 10;
            var btnTestDetect = _Btn("🔍 測試偵測", 140, 100, 0, 100, 26); btnTestDetect.FontSize = 10; btnTestDetect.Margin = new Thickness(4, 0, 0, 0);
            var numRetries = _Num(45, 10, 1, 30); var numThreshold = _Num(50, 85, 50, 99);
            var lRt = _L("重試:", 9, _B(192, 192, 192)); lRt.VerticalAlignment = VerticalAlignment.Center; lRt.Margin = new Thickness(8, 0, 0, 0);
            var lRtU = _L("次", 9, _B(192, 192, 192)); lRtU.VerticalAlignment = VerticalAlignment.Center;
            numRetries.Margin = new Thickness(2, 0, 2, 0);
            var lTh = _L("閾值:", 9, _B(192, 192, 192)); lTh.VerticalAlignment = VerticalAlignment.Center; lTh.Margin = new Thickness(8, 0, 0, 0);
            var lThU = _L("%", 9, _B(192, 192, 192)); lThU.VerticalAlignment = VerticalAlignment.Center;
            numThreshold.Margin = new Thickness(2, 0, 2, 0);

            var detectWrap = new WCtl.WrapPanel { Margin = new Thickness(0, 4, 0, 0) };
            detectWrap.Children.Add(btnCaptureTmpl); detectWrap.Children.Add(btnTestDetect);
            detectWrap.Children.Add(lRt); detectWrap.Children.Add(numRetries); detectWrap.Children.Add(lRtU);
            detectWrap.Children.Add(lTh); detectWrap.Children.Add(numThreshold); detectWrap.Children.Add(lThU);
            detectWrap.Children.Add(imgBorder); detectWrap.Children.Add(lblTmplStatus);
            main.Children.Add(detectWrap);

            string? dialogTemplatePath = null;
            btnCaptureTmpl.Click += (s, e) =>
            {
                if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle))
                { WinMessageBox.Show("請先鎖定目標視窗！\n\n步驟：先開啟遊戲 NPC 對話框，再點擊「擷取模板」", "未鎖定", WinMBBtn.OK, WinMBIcon.Warning); return; }
                using var screenshot = CaptureGameWindowForDetection();
                if (screenshot == null) { WinMessageBox.Show("無法截取遊戲視窗！", "截取失敗", WinMBBtn.OK, WinMBIcon.Warning); return; }
                win.Hide(); var region = RegionSelector.SelectRegion(screenshot); win.Show();
                if (region.HasValue && region.Value.Width > 5 && region.Value.Height > 5)
                {
                    using var cropped = new Bitmap(region.Value.Width, region.Value.Height, PixelFormat.Format24bppRgb);
                    using (var g = Graphics.FromImage(cropped)) g.DrawImage(screenshot, new Rectangle(0, 0, cropped.Width, cropped.Height), region.Value, GraphicsUnit.Pixel);
                    if (!Directory.Exists(DialogTemplateDir)) Directory.CreateDirectory(DialogTemplateDir);
                    string fn = $"dialog_template_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                    dialogTemplatePath = Path.Combine(DialogTemplateDir, fn);
                    cropped.Save(dialogTemplatePath, ImageFormat.Png);
                    imgTemplate.Source = _BmpToWpf(cropped);
                    lblTmplStatus.Text = $"✅ {region.Value.Width}x{region.Value.Height}"; lblTmplStatus.Foreground = _B(0, 255, 0);
                    AddLog($"📷 對話框模板已儲存: {fn} ({region.Value.Width}x{region.Value.Height})");
                }
            };
            btnTestDetect.Click += (s, e) =>
            {
                if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle)) { WinMessageBox.Show("請先鎖定目標視窗！", "未鎖定", WinMBBtn.OK, WinMBIcon.Warning); return; }
                if (string.IsNullOrEmpty(dialogTemplatePath) || !File.Exists(dialogTemplatePath)) { WinMessageBox.Show("請先擷取對話框模板！", "未設定", WinMBBtn.OK, WinMBIcon.Warning); return; }
                AddLog("🔍 測試偵測中..."); double mt = _NV(numThreshold) / 100.0; string dtp = dialogTemplatePath;
                new System.Threading.Thread(() =>
                {
                    try
                    {
                        using var tmpl = new Bitmap(dtp); using var ss = CaptureGameWindowForDetection();
                        if (ss == null) { Dispatcher.BeginInvoke(new Action(() => AddLog("⚠️ 截圖失敗"))); return; }
                        var (found, score, mx, my) = FindTemplateInImage(ss, tmpl, mt); double sr = Math.Round(score * 100, 1);
                        Dispatcher.BeginInvoke(new Action(() => { if (found) AddLog($"✅ 偵測成功！匹配度={sr}% 位置=({mx},{my})"); else AddLog($"❌ 未偵測到 (最佳匹配={sr}%，閾值={_NV(numThreshold)}%)"); }));
                    }
                    catch (Exception ex) { Dispatcher.BeginInvoke(new Action(() => AddLog($"❌ 測試偵測失敗: {ex.Message}"))); }
                }) { IsBackground = true }.Start();
            };

            // Test click + Add
            var btnTestClick = _Btn("🧪 測試點擊", 180, 120, 0, 110, 30);
            var btnAddTask = _Btn("➕ 新增排程", 0, 150, 80, 120, 30); btnAddTask.Margin = new Thickness(8, 0, 0, 0);
            var actionRow = _Row(btnTestClick, btnAddTask); actionRow.Margin = new Thickness(0, 8, 0, 0);
            main.Children.Add(actionRow);

            btnTestClick.Click += (s, e) =>
            {
                if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle)) { WinMessageBox.Show("請先鎖定目標視窗！", "未鎖定", WinMBBtn.OK, WinMBIcon.Warning); return; }
                int tx = (int)_NV(numCX), ty = (int)_NV(numCY); AddLog($"🧪 測試點擊: ({tx},{ty})");
                new System.Threading.Thread(() => { SendMouseClickToWindow(targetWindowHandle, tx, ty); Dispatcher.BeginInvoke(new Action(() => AddLog("🧪 測試點擊完成"))); }) { IsBackground = true }.Start();
            };

            // Task list
            var h2 = _L("📅 排程清單", 11, _B(255, 255, 0)); h2.FontWeight = FontWeights.Bold; h2.Margin = new Thickness(0, 10, 0, 4);
            main.Children.Add(h2);

            var lstTasks = new WCtl.ListBox { Height = 180, Background = _B(30, 30, 35), Foreground = WMedia.Brushes.White, BorderBrush = _B(60, 60, 65), BorderThickness = new Thickness(1), FontFamily = _monoFont, FontSize = 10 };
            main.Children.Add(lstTasks);

            Action refreshList = () =>
            {
                lstTasks.Items.Clear();
                foreach (var task in scheduleTasks)
                {
                    string sn = string.IsNullOrEmpty(task.ScriptPath) ? "(當前腳本)" : Path.GetFileName(task.ScriptPath);
                    string es = task.EndTime.HasValue ? task.EndTime.Value.ToString("HH:mm:ss") : "不限";
                    string st = task.HasStarted ? "已觸發" : (task.Enabled ? "等待中" : "已完成");
                    string rt = task.ReturnToTownEnabled ? "✔" : "";
                    lstTasks.Items.Add($"{sn,-20} 開始={task.StartTime:HH:mm:ss} 結束={es} 循環={task.LoopCount} 回程={rt} {st}");
                }
            };
            refreshList();

            Func<WCtl.DatePicker, WCtl.TextBox, DateTime> getDT = (dp, txt) =>
            {
                var d = dp.SelectedDate ?? DateTime.Today;
                return TimeSpan.TryParse(txt.Text, out var t) ? d.Date + t : d;
            };

            btnAddTask.Click += (s, e) =>
            {
                var startDT = getDT(dpStart, txtStartTime);
                if (startDT <= DateTime.Now) { WinMessageBox.Show("開始時間必須為未來！", "錯誤", WinMBBtn.OK, WinMBIcon.Warning); return; }
                DateTime? endDT = null;
                if (chkEndTime.IsChecked == true) { endDT = getDT(dpEnd, txtEndTime); if (endDT <= startDT) { WinMessageBox.Show("結束時間必須晚於開始時間！", "錯誤", WinMBBtn.OK, WinMBIcon.Warning); return; } }
                string? sp = txtScript.Tag as string;
                if (string.IsNullOrEmpty(sp) && recordedEvents.Count == 0) { WinMessageBox.Show("請選擇腳本或先載入腳本！", "錯誤", WinMBBtn.OK, WinMBIcon.Warning); return; }

                var nt = new ScheduleTask { ScriptPath = sp ?? string.Empty, StartTime = startDT, EndTime = endDT, LoopCount = (int)_NV(numLoop), Enabled = true, HasStarted = false, ReturnToTownEnabled = chkReturn.IsChecked == true, ReturnClickX = (int)_NV(numCX), ReturnClickY = (int)_NV(numCY), SitDownDelaySeconds = _NV(numSitDelay) };
                if (txtPreKey.Tag is Keys pk) nt.ReturnPreKeyCode = (int)pk;
                if (txtConfirmKey.Tag is Keys ck) nt.ReturnConfirmKeyCode = (int)ck;
                if (txtSitKey.Tag is Keys sk) nt.SitDownKeyCode = (int)sk;
                nt.DialogDetectionEnabled = chkDialog.IsChecked == true;
                nt.DialogTemplatePath = dialogTemplatePath ?? string.Empty;
                nt.DialogDetectionMaxRetries = (int)_NV(numRetries);
                nt.DialogDetectionThreshold = _NV(numThreshold) / 100.0;

                scheduleTasks.Add(nt); schedulerTimer.Start(); refreshList();
                string ei = endDT.HasValue ? $", 結束={endDT.Value:HH:mm:ss}" : "";
                string ri = "";
                if (chkReturn.IsChecked == true) { ri = $", 回程=點擊({_NV(numCX)},{_NV(numCY)})"; if (txtPreKey.Tag is Keys pk2 && pk2 != Keys.None) ri += $" 前置={GetKeyDisplayName(pk2)}"; if (txtConfirmKey.Tag is Keys ck2 && ck2 != Keys.None) ri += $" 確認={GetKeyDisplayName(ck2)}"; if (txtSitKey.Tag is Keys sk2 && sk2 != Keys.None) ri += $" 坐下={GetKeyDisplayName(sk2)}"; }
                AddLog($"新增排程：{(string.IsNullOrEmpty(sp) ? "當前腳本" : Path.GetFileName(sp))}, 開始={startDT:HH:mm:ss}{ei}, 循環={_NV(numLoop)}{ri}");
            };

            // Bottom buttons
            var btnRemove = _Btn("🗑️ 刪除選中", 150, 60, 60, 110, 30);
            var btnClearAll = _Btn("清空全部", 120, 80, 40, 90, 30); btnClearAll.Margin = new Thickness(8, 0, 0, 0);
            var btnCloseSched = _Btn("關閉", 80, 80, 85, 80, 30); btnCloseSched.Margin = new Thickness(8, 0, 0, 0); btnCloseSched.BorderBrush = _B(100, 100, 105);
            var lblCountdown = _L("", 10, _B(255, 255, 0)); lblCountdown.VerticalAlignment = VerticalAlignment.Center; lblCountdown.Margin = new Thickness(12, 0, 0, 0);

            btnRemove.Click += (s, e) => { if (lstTasks.SelectedIndex >= 0 && lstTasks.SelectedIndex < scheduleTasks.Count) { scheduleTasks.RemoveAt(lstTasks.SelectedIndex); refreshList(); AddLog("已刪除排程"); } };
            btnClearAll.Click += (s, e) => { scheduleTasks.Clear(); schedulerTimer.Stop(); refreshList(); AddLog("已清空所有排程"); };
            btnCloseSched.Click += (s, e) => win.Close();

            var bottomRow = _Row(btnRemove, btnClearAll, btnCloseSched, lblCountdown); bottomRow.Margin = new Thickness(0, 8, 0, 0);
            main.Children.Add(bottomRow);

            var cdTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            cdTimer.Tick += (s, e) =>
            {
                var next = scheduleTasks.Where(t => t.Enabled && !t.HasStarted).OrderBy(t => t.StartTime).FirstOrDefault();
                if (next != null) { var rem = next.StartTime - DateTime.Now; lblCountdown.Text = rem.TotalSeconds > 0 ? $"下個排程：{rem:hh\\:mm\\:ss} 後開始" : "正在觸發..."; }
                else
                {
                    var active = scheduleTasks.Where(t => t.HasStarted && t.Enabled && t.EndTime.HasValue).FirstOrDefault();
                    if (active != null) { var rem = active.EndTime!.Value - DateTime.Now; lblCountdown.Text = rem.TotalSeconds > 0 ? $"自動停止：{rem:hh\\:mm\\:ss} 後" : "正在停止..."; }
                    else lblCountdown.Text = scheduleTasks.Count > 0 ? "所有排程已完成" : "";
                }
                refreshList();
            };
            cdTimer.Start();
            win.Closing += (s, e) => cdTimer.Stop();

            rootSv.Content = main;
            win.Content = rootSv;
            win.ShowDialog();
        }

        // ═══════════════════════════════════════════════
        //  📊 即時執行統計 (WPF)
        // ═══════════════════════════════════════════════

        private void ShowStatistics()
        {
            var win = _CyWin("📊 即時執行統計", 450, 520);
            var root = new WCtl.StackPanel { Margin = new Thickness(20) };

            var lblInd = _L(statistics.CurrentSessionStart.HasValue ? "● 播放中" : "○ 已停止", 14,
                statistics.CurrentSessionStart.HasValue ? _B(0, 255, 0) : _B(128, 128, 128));
            lblInd.FontWeight = FontWeights.Bold; root.Children.Add(lblInd);

            var sh1 = _L("📌 當前會話", 12, _B(0, 255, 255)); sh1.FontWeight = FontWeights.Bold; sh1.Margin = new Thickness(0, 12, 0, 4);
            root.Children.Add(sh1);

            var lblST = _L("會話時長: --:--:--", 11); lblST.FontFamily = _monoFont;
            var lblCL = _L($"當前循環: {statistics.CurrentLoopCount}", 11); lblCL.FontFamily = _monoFont;
            var lblSI = _L($"腳本事件: {recordedEvents.Count} 個", 10, _B(192, 192, 192));
            root.Children.Add(lblST); root.Children.Add(lblCL); root.Children.Add(lblSI);
            root.Children.Add(_Sep());

            var sh2 = _L("📈 累計統計", 12, _B(255, 255, 0)); sh2.FontWeight = FontWeights.Bold; sh2.Margin = new Thickness(0, 0, 0, 4);
            root.Children.Add(sh2);

            var lblTP = _L($"播放次數: {statistics.TotalPlayCount}", 11); lblTP.FontFamily = _monoFont;
            var lblTT = _L("總播放時長: 00:00:00", 11); lblTT.FontFamily = _monoFont;
            var lblLP = _L($"最後播放: {(statistics.LastPlayTime.HasValue ? statistics.LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "從未播放")}", 10, _B(192, 192, 192));
            root.Children.Add(lblTP); root.Children.Add(lblTT); root.Children.Add(lblLP);
            root.Children.Add(_Sep());

            var sh3 = _L("⚡ 自定義按鍵觸發", 12, _B(200, 150, 255)); sh3.FontWeight = FontWeights.Bold; sh3.Margin = new Thickness(0, 0, 0, 4);
            root.Children.Add(sh3);

            var lst = new WCtl.ListBox { Height = 100, Background = _B(40, 40, 45), Foreground = WMedia.Brushes.White, BorderBrush = _B(60, 60, 65), BorderThickness = new Thickness(1), FontFamily = _monoFont, FontSize = 10 };
            root.Children.Add(lst);

            Action updateCK = () =>
            {
                lst.Items.Clear(); bool any = false;
                for (int i = 0; i < 15; i++)
                    if (customKeySlots[i].Enabled && customKeySlots[i].KeyCode != Keys.None)
                    { any = true; lst.Items.Add($"  #{i + 1} {GetKeyDisplayName(customKeySlots[i].KeyCode),-15} 觸發: {statistics.CustomKeyTriggerCounts[i]} 次"); }
                if (!any) lst.Items.Add("  (無啟用的自定義按鍵)");
            };
            updateCK();

            var bReset = _Btn("🔄 重置統計", 150, 80, 60, 100, 35);
            var bClose = _Btn("關閉", 80, 80, 85, 80, 35); bClose.Margin = new Thickness(10, 0, 0, 0); bClose.BorderBrush = _B(100, 100, 105);
            bReset.Click += (s, e) => { if (WinMessageBox.Show("確定重置所有統計資料？", "重置統計", WinMBBtn.YesNo, WinMBIcon.Warning) == WinDR.Yes) { statistics.Reset(); AddLog("統計資料已重置"); } };
            bClose.Click += (s, e) => win.Close();

            var bp = _Row(bReset, bClose); bp.HorizontalAlignment = System.Windows.HorizontalAlignment.Center; bp.Margin = new Thickness(0, 15, 0, 0);
            root.Children.Add(bp);

            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            timer.Tick += (s, e) =>
            {
                bool act = statistics.CurrentSessionStart.HasValue;
                lblInd.Text = act ? "● 播放中" : "○ 已停止";
                lblInd.Foreground = act ? _B(0, 255, 0) : _B(128, 128, 128);
                if (act) { var st = DateTime.Now - statistics.CurrentSessionStart!.Value; lblST.Text = $"會話時長: {(int)st.TotalHours:D2}:{st.Minutes:D2}:{st.Seconds:D2}"; }
                else lblST.Text = "會話時長: --:--:--";
                lblCL.Text = $"當前循環: {statistics.CurrentLoopCount}"; lblSI.Text = $"腳本事件: {recordedEvents.Count} 個";
                double live = statistics.TotalPlayTimeSeconds;
                if (act) live += (DateTime.Now - statistics.CurrentSessionStart!.Value).TotalSeconds;
                var tt = TimeSpan.FromSeconds(live);
                lblTP.Text = $"播放次數: {statistics.TotalPlayCount + (act ? 1 : 0)}";
                lblTT.Text = $"總播放時長: {(int)tt.TotalHours:D2}:{tt.Minutes:D2}:{tt.Seconds:D2}";
                lblLP.Text = $"最後播放: {(statistics.LastPlayTime.HasValue ? statistics.LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : (act ? "進行中..." : "從未播放"))}";
                updateCK();
            };
            timer.Start();
            win.Closing += (s, e) => timer.Stop();

            win.Content = root;
            win.ShowDialog();
        }

        // ═══════════════════════════════════════════════
        //  🗺️ 小地圖校準精靈 (WPF)
        // ═══════════════════════════════════════════════

        private void OpenMinimapCalibration()
        {
            if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle))
            {
                WinMessageBox.Show("請先鎖定目標視窗！\n\n步驟：\n1. 開啟遊戲\n2. 點擊「手動鎖定」選擇遊戲視窗",
                    "未鎖定視窗", WinMBBtn.OK, WinMBIcon.Warning);
                return;
            }

            if (minimapTracker == null) minimapTracker = new MinimapTracker();
            minimapTracker.AttachToWindow(targetWindowHandle);
            bool hasCalibration = minimapTracker.LoadCalibration(CalibrationFilePath);

            var win = _CyWin("🗺️ 小地圖校準精靈", 750, 600);
            var rootSv = new WCtl.ScrollViewer { VerticalScrollBarVisibility = WCtl.ScrollBarVisibility.Auto };
            var main = new WCtl.StackPanel { Margin = new Thickness(10) };

            // Two-column layout
            var topGrid = new WCtl.Grid();
            topGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(365) });
            topGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left: Minimap preview
            var grpPrev = _Grp("小地圖預覽 (點擊角色圖標學習顏色)");
            var prevPanel = new WCtl.StackPanel { Margin = new Thickness(4) };
            var imgMinimap = new WCtl.Image { Width = 330, Height = 220, Stretch = WMedia.Stretch.Uniform };
            var imgMinimapBorder = new WCtl.Border { Child = imgMinimap, Width = 340, Height = 225, BorderBrush = _B(60, 60, 65), BorderThickness = new Thickness(1), Background = _B(0, 0, 0), Cursor = System.Windows.Input.Cursors.Cross };
            prevPanel.Children.Add(imgMinimapBorder);
            var lblMmInfo = _L("區域: 未設定", 9, _B(192, 192, 192)); lblMmInfo.Margin = new Thickness(0, 4, 0, 0);
            var lblColorInfo = _L("💡 點擊預覽圖上的角色圖標來學習顏色", 9, _B(255, 255, 0));
            prevPanel.Children.Add(lblMmInfo); prevPanel.Children.Add(lblColorInfo);
            grpPrev.Content = prevPanel;
            WCtl.Grid.SetColumn(grpPrev, 0); topGrid.Children.Add(grpPrev);

            // Right panels
            var rightPanel = new WCtl.StackPanel { Margin = new Thickness(8, 0, 0, 0) };

            // Region
            var grpReg = _Grp("📍 小地圖區域"); var regPanel = new WCtl.StackPanel { Margin = new Thickness(4) };
            var nX = _Num(55, minimapTracker.MinimapRegion.X, 0, 2000); var nY = _Num(55, minimapTracker.MinimapRegion.Y, 0, 2000);
            var nW = _Num(50, Math.Max(10, minimapTracker.MinimapRegion.Width), 10, 500); var nH = _Num(50, Math.Max(10, minimapTracker.MinimapRegion.Height), 10, 500);
            var rX = _L("X:", 10, _B(192, 192, 192)); rX.VerticalAlignment = VerticalAlignment.Center;
            var rY = _L("Y:", 10, _B(192, 192, 192)); rY.VerticalAlignment = VerticalAlignment.Center; rY.Margin = new Thickness(6, 0, 0, 0);
            var rW = _L("寬:", 10, _B(192, 192, 192)); rW.VerticalAlignment = VerticalAlignment.Center; rW.Margin = new Thickness(6, 0, 0, 0);
            var rH = _L("高:", 10, _B(192, 192, 192)); rH.VerticalAlignment = VerticalAlignment.Center; rH.Margin = new Thickness(6, 0, 0, 0);
            nX.Margin = nY.Margin = nW.Margin = nH.Margin = new Thickness(2, 0, 0, 0);
            regPanel.Children.Add(_Row(rX, nX, rY, nY, rW, nW, rH, nH));
            var btnSelReg = _Btn("🖱️ 框選區域", 0, 140, 80, 95, 28); btnSelReg.FontSize = 10;
            var btnCapPrev = _Btn("📷 截取預覽", 80, 80, 90, 95, 28); btnCapPrev.FontSize = 10; btnCapPrev.Margin = new Thickness(4, 0, 0, 0);
            var regBtnRow = _Row(btnSelReg, btnCapPrev); regBtnRow.Margin = new Thickness(0, 4, 0, 0);
            regPanel.Children.Add(regBtnRow);
            grpReg.Content = regPanel; rightPanel.Children.Add(grpReg);

            // Bounds
            var grpBnd = _Grp("🎯 活動範圍邊界"); var bndPanel = new WCtl.StackPanel { Margin = new Thickness(4) };
            var lBL = _L("左: ---", 10, _B(192, 192, 192)); lBL.Width = 65;
            var lBR = _L("右: ---", 10, _B(192, 192, 192)); lBR.Width = 65;
            var lBT = _L("上: ---", 10, _B(192, 192, 192)); lBT.Width = 65;
            var lBB = _L("下: ---", 10, _B(192, 192, 192)); lBB.Width = 65;
            bndPanel.Children.Add(_Row(lBL, lBR, lBT, lBB));
            var bSL = _Btn("⬅左", 60, 60, 70, 52, 26); bSL.FontSize = 10;
            var bSR = _Btn("右➡", 60, 60, 70, 52, 26); bSR.FontSize = 10; bSR.Margin = new Thickness(3, 0, 0, 0);
            var bST = _Btn("⬆上", 60, 60, 70, 52, 26); bST.FontSize = 10; bST.Margin = new Thickness(3, 0, 0, 0);
            var bSB = _Btn("下⬇", 60, 60, 70, 52, 26); bSB.FontSize = 10; bSB.Margin = new Thickness(3, 0, 0, 0);
            var bRB = _Btn("🔄重設", 100, 60, 60, 56, 26); bRB.FontSize = 10; bRB.Margin = new Thickness(3, 0, 0, 0);
            var bndBtnRow = _Row(bSL, bSR, bST, bSB, bRB); bndBtnRow.Margin = new Thickness(0, 4, 0, 0);
            bndPanel.Children.Add(bndBtnRow);
            var lblBTip = _L("F7: 依序設定左→右→下→上", 9, _B(0, 255, 255));
            bndPanel.Children.Add(lblBTip);
            grpBnd.Content = bndPanel; rightPanel.Children.Add(grpBnd);

            // Color
            var grpClr = _Grp("🎨 角色圖標顏色"); var clrPanel = new WCtl.StackPanel { Margin = new Thickness(4) };
            var nHMin = _Num(50, (double)minimapTracker.HueRange.Min, 0, 360); var nHMax = _Num(50, (double)minimapTracker.HueRange.Max, 0, 360);
            var nSatV = _Num(50, (double)(minimapTracker.MinSaturation * 100), 0, 100);
            var clrPreview = new WCtl.Border { Width = 28, Height = 22, BorderBrush = _B(80, 80, 85), BorderThickness = new Thickness(1), Margin = new Thickness(8, 0, 0, 0) };
            var cH = _L("色相:", 10, _B(192, 192, 192)); cH.VerticalAlignment = VerticalAlignment.Center;
            var cT = _L("~", 10, _B(192, 192, 192)); cT.VerticalAlignment = VerticalAlignment.Center; cT.Margin = new Thickness(2, 0, 2, 0);
            var cS = _L("飽和:", 10, _B(192, 192, 192)); cS.VerticalAlignment = VerticalAlignment.Center; cS.Margin = new Thickness(8, 0, 0, 0);
            var cP = _L("%", 10, _B(192, 192, 192)); cP.VerticalAlignment = VerticalAlignment.Center;
            nHMin.Margin = nHMax.Margin = nSatV.Margin = new Thickness(2, 0, 0, 0);
            clrPanel.Children.Add(_Row(cH, nHMin, cT, nHMax, cS, nSatV, cP, clrPreview));
            grpClr.Content = clrPanel; rightPanel.Children.Add(grpClr);

            WCtl.Grid.SetColumn(rightPanel, 1); topGrid.Children.Add(rightPanel);
            main.Children.Add(topGrid);

            // Detection result
            var grpRes = _Grp("📊 偵測結果"); grpRes.Margin = new Thickness(0, 8, 0, 0);
            var resGrid = new WCtl.Grid();
            resGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(200) });
            resGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(250) });
            resGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var lblDet = _L("尚未偵測", 12, _B(128, 128, 128)); lblDet.FontFamily = _monoFont; lblDet.FontWeight = FontWeights.Bold; lblDet.TextWrapping = TextWrapping.Wrap;
            WCtl.Grid.SetColumn(lblDet, 0); resGrid.Children.Add(lblDet);
            var imgDbg = new WCtl.Image { Width = 240, Height = 70, Stretch = WMedia.Stretch.Uniform };
            var imgDbgBrd = new WCtl.Border { Child = imgDbg, BorderBrush = _B(60, 60, 65), BorderThickness = new Thickness(1), Background = _B(0, 0, 0) };
            WCtl.Grid.SetColumn(imgDbgBrd, 1); resGrid.Children.Add(imgDbgBrd);

            var detBtnP = new WCtl.StackPanel();
            var btnTst = _Btn("🎯 測試偵測", 180, 120, 0, 100, 30); btnTst.FontSize = 10;
            var btnAuto = _Btn("🔄 連續測試", 80, 80, 90, 85, 30); btnAuto.FontSize = 10; btnAuto.Margin = new Thickness(0, 4, 0, 0);
            detBtnP.Children.Add(btnTst); detBtnP.Children.Add(btnAuto);
            WCtl.Grid.SetColumn(detBtnP, 2); resGrid.Children.Add(detBtnP);
            grpRes.Content = resGrid; main.Children.Add(grpRes);

            // Bottom
            var lblStat = _L(hasCalibration ? "✅ 已載入先前的校準設定" : "⚠️ 尚未校準", 10, hasCalibration ? _B(0, 255, 0) : _B(128, 128, 128));
            lblStat.VerticalAlignment = VerticalAlignment.Center;
            var btnCorr = _Btn("🎯 位置修正", 180, 120, 0, 110, 35); btnCorr.FontSize = 11;
            var btnSaveCal = _Btn("💾 儲存校準", 0, 140, 80, 110, 35); btnSaveCal.FontSize = 11; btnSaveCal.Margin = new Thickness(6, 0, 0, 0);
            var btnCloseCal = _Btn("關閉", 80, 80, 85, 80, 35); btnCloseCal.Margin = new Thickness(6, 0, 0, 0); btnCloseCal.BorderBrush = _B(100, 100, 105);
            var btmGrid = new WCtl.Grid { Margin = new Thickness(0, 8, 0, 0) };
            btmGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            btmGrid.ColumnDefinitions.Add(new WCtl.ColumnDefinition { Width = GridLength.Auto });
            WCtl.Grid.SetColumn(lblStat, 0); btmGrid.Children.Add(lblStat);
            var btmBtns = _Row(btnCorr, btnSaveCal, btnCloseCal);
            WCtl.Grid.SetColumn(btmBtns, 1); btmGrid.Children.Add(btmBtns);
            main.Children.Add(btmGrid);

            // ═══ Logic & State ═══
            int boundLeft = 0, boundRight = 100, boundTop = 0, boundBottom = 100;
            bool boundsExplicitlySet = false;
            DispatcherTimer? autoTestTimer = null;
            Bitmap? currentMinimapBitmap = null;

            if (hasCalibration && minimapTracker.MapBounds.Width > 0)
            {
                boundLeft = minimapTracker.MapBounds.Left; boundRight = minimapTracker.MapBounds.Right;
                boundTop = minimapTracker.MapBounds.Top; boundBottom = minimapTracker.MapBounds.Bottom;
                boundsExplicitlySet = true;
                lBL.Text = $"左: {boundLeft}"; lBR.Text = $"右: {boundRight}";
                lBT.Text = $"上: {boundTop}"; lBB.Text = $"下: {boundBottom}";
            }

            // Boundary sync timer
            var syncTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            syncTimer.Tick += (s, e) =>
            {
                if (minimapTracker != null && minimapTracker.MapBounds.Width > 0)
                {
                    var mb = minimapTracker.MapBounds; int mbR = mb.Left + mb.Width, mbB = mb.Top + mb.Height;
                    if (mb.Left != boundLeft || mbR != boundRight || mb.Top != boundTop || mbB != boundBottom)
                    {
                        boundLeft = mb.Left; boundRight = mbR; boundTop = mb.Top; boundBottom = mbB; boundsExplicitlySet = true;
                        lBL.Text = $"左: {boundLeft}"; lBR.Text = $"右: {boundRight}";
                        lBT.Text = $"上: {boundTop}"; lBB.Text = $"下: {boundBottom}";
                        lBL.Foreground = lBR.Foreground = lBT.Foreground = lBB.Foreground = _B(0, 255, 0);
                    }
                }
                if (_boundarySetState >= 0 && _boundarySetState < 4)
                { string[] nn = { "左邊界", "右邊界", "下邊界", "上邊界" }; lblBTip.Text = $"F7: 等待設定【{nn[_boundarySetState]}】..."; lblBTip.Foreground = _B(255, 255, 0); }
                else { lblBTip.Text = "F7: 依序設定左→右→上→下"; lblBTip.Foreground = _B(0, 255, 255); }
            };
            syncTimer.Start();

            Action updateClrPreview = () =>
            {
                float hue = ((float)_NV(nHMin) + (float)_NV(nHMax)) / 2f;
                var c = ColorFromHSV(hue, 0.8, 0.9); clrPreview.Background = ToBrush(c);
            };
            updateClrPreview();

            Action updateRegion = () => { minimapTracker.MinimapRegion = new Rectangle((int)_NV(nX), (int)_NV(nY), (int)_NV(nW), (int)_NV(nH)); lblMmInfo.Text = $"區域: X={_NV(nX)}, Y={_NV(nY)}, {_NV(nW)}x{_NV(nH)}"; };
            Action updateBounds = () => minimapTracker.MapBounds = new Rectangle(boundLeft, boundTop, boundRight - boundLeft, boundBottom - boundTop);
            Action updateClr = () => { minimapTracker.HueRange = ((float)_NV(nHMin), (float)_NV(nHMax)); minimapTracker.MinSaturation = (float)_NV(nSatV) / 100f; updateClrPreview(); };
            Func<(int x, int y, bool ok)> detect = () => { updateRegion(); updateClr(); return minimapTracker.ReadPosition(); };

            nX.LostFocus += (s, e) => updateRegion(); nY.LostFocus += (s, e) => updateRegion();
            nW.LostFocus += (s, e) => updateRegion(); nH.LostFocus += (s, e) => updateRegion();
            nHMin.LostFocus += (s, e) => updateClr(); nHMax.LostFocus += (s, e) => updateClr();
            nSatV.LostFocus += (s, e) => updateClr();

            // Click to learn color
            imgMinimapBorder.MouseLeftButtonDown += (s, e) =>
            {
                if (currentMinimapBitmap == null || imgMinimap.Source == null) return;
                var pos = e.GetPosition(imgMinimap);
                double srcW = currentMinimapBitmap.Width, srcH = currentMinimapBitmap.Height;
                double sX = srcW / imgMinimap.ActualWidth, sY = srcH / imgMinimap.ActualHeight;
                double sc = Math.Max(sX, sY);
                double rW = srcW / sc, rH = srcH / sc;
                double oX = (imgMinimap.ActualWidth - rW) / 2, oY = (imgMinimap.ActualHeight - rH) / 2;
                int iX = (int)((pos.X - oX) * sc), iY = (int)((pos.Y - oY) * sc);
                if (iX < 0 || iX >= currentMinimapBitmap.Width || iY < 0 || iY >= currentMinimapBitmap.Height) return;
                var cc = currentMinimapBitmap.GetPixel(iX, iY);
                ColorToHSV(cc, out float ch, out float cs, out float cv);
                nHMin.Text = ((int)Math.Max(0, ch - 15)).ToString(); nHMax.Text = ((int)Math.Min(360, ch + 15)).ToString();
                nSatV.Text = ((int)Math.Max(20, cs * 100 - 20)).ToString(); updateClr();
                lblColorInfo.Text = $"✅ 已學習顏色: RGB({cc.R},{cc.G},{cc.B}) H={ch:F0}°"; lblColorInfo.Foreground = _B(0, 255, 0);
                lblStat.Text = $"✅ 已學習顏色 (色相: {ch:F0}°)"; lblStat.Foreground = _B(0, 255, 0);
            };

            btnSelReg.Click += (s, e) =>
            {
                using var ss = minimapTracker.CaptureFullWindow();
                if (ss == null) { WinMessageBox.Show("無法截取遊戲視窗！", "截取失敗", WinMBBtn.OK, WinMBIcon.Warning); return; }
                win.Hide(); var region = RegionSelector.SelectRegion(ss); win.Show();
                if (region.HasValue)
                {
                    nX.Text = region.Value.X.ToString(); nY.Text = region.Value.Y.ToString();
                    nW.Text = region.Value.Width.ToString(); nH.Text = region.Value.Height.ToString(); updateRegion();
                    using var bmp = minimapTracker.CaptureMinimap();
                    if (bmp != null) { currentMinimapBitmap?.Dispose(); currentMinimapBitmap = new Bitmap(bmp); imgMinimap.Source = _BmpToWpf(bmp); }
                    lblStat.Text = "✅ 區域已框選！點擊預覽圖上的角色學習顏色"; lblStat.Foreground = _B(0, 255, 0);
                }
            };
            btnCapPrev.Click += (s, e) =>
            {
                updateRegion();
                using var bmp = minimapTracker.CaptureMinimap();
                if (bmp != null) { currentMinimapBitmap?.Dispose(); currentMinimapBitmap = new Bitmap(bmp); imgMinimap.Source = _BmpToWpf(bmp); }
                else WinMessageBox.Show("截取失敗！", "錯誤", WinMBBtn.OK, WinMBIcon.Warning);
            };

            // Boundary buttons
            Action<string, Action<int>, WCtl.TextBlock> setBnd = (name, setter, label) =>
            {
                var (x, y, ok) = detect();
                if (ok) { int val = name.Contains("左") || name.Contains("右") ? x : y; setter(val); label.Text = $"{name[0]}: {val}"; label.Foreground = _B(0, 255, 0); boundsExplicitlySet = true; updateBounds(); lblStat.Text = $"✅ {name}邊界已設定: {val}"; lblStat.Foreground = _B(0, 255, 0); }
                else WinMessageBox.Show("偵測失敗！請先點擊預覽圖學習角色圖標顏色", "偵測失敗", WinMBBtn.OK, WinMBIcon.Warning);
            };
            bSL.Click += (s, e) => setBnd("左", v => boundLeft = v, lBL);
            bSR.Click += (s, e) => setBnd("右", v => boundRight = v, lBR);
            bST.Click += (s, e) => setBnd("上", v => boundTop = v, lBT);
            bSB.Click += (s, e) => setBnd("下", v => boundBottom = v, lBB);
            bRB.Click += (s, e) =>
            {
                boundLeft = 0; boundRight = (int)_NV(nW); boundTop = 0; boundBottom = (int)_NV(nH); boundsExplicitlySet = false;
                lBL.Text = "左: ---"; lBR.Text = "右: ---"; lBT.Text = "上: ---"; lBB.Text = "下: ---";
                lBL.Foreground = lBR.Foreground = lBT.Foreground = lBB.Foreground = _B(192, 192, 192); updateBounds();
            };

            // Test detection
            btnTst.Click += (s, e) =>
            {
                updateRegion(); updateBounds(); updateClr();
                var (px, py, ok) = detect();
                if (ok)
                {
                    string ci = $"聚類: {minimapTracker.LastTotalClusterCount}→{minimapTracker.LastValidClusterCount}";
                    if (boundsExplicitlySet)
                    {
                        int tol = 5; bool inB = px >= (boundLeft - tol) && px <= (boundRight + tol) && py >= (boundTop - tol) && py <= (boundBottom + tol);
                        int dL = px - boundLeft, dR = boundRight - px, dT = py - boundTop, dB = boundBottom - py;
                        int minD = Math.Min(Math.Min(dL, dR), Math.Min(dT, dB)); string di = inB ? "" : $" (偏移{Math.Abs(minD)}px)";
                        lblDet.Text = $"位置: ({px}, {py})\n{(inB ? "✅ 在範圍內" : $"⚠️ 超出範圍{di}")}\n信心度: {minimapTracker.DetectionConfidence}% | {ci}";
                        lblDet.Foreground = inB ? _B(0, 255, 0) : _B(255, 165, 0);
                    }
                    else { lblDet.Text = $"位置: ({px}, {py})\n📍 (未設定活動範圍)\n信心度: {minimapTracker.DetectionConfidence}% | {ci}"; lblDet.Foreground = _B(0, 255, 0); }
                    using var dbg = minimapTracker.CreateDebugImage();
                    if (dbg != null) imgDbg.Source = _BmpToWpf(dbg);
                }
                else { string fc = $"聚類: {minimapTracker.LastTotalClusterCount}→{minimapTracker.LastValidClusterCount}"; lblDet.Text = $"❌ 偵測失敗\n請先學習顏色\n{fc}"; lblDet.Foreground = _B(255, 0, 0); }
            };

            btnAuto.Click += (s, e) =>
            {
                if (autoTestTimer == null)
                {
                    autoTestTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
                    autoTestTimer.Tick += (ts, te) => btnTst.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
                    autoTestTimer.Start(); btnAuto.Content = "⏹ 停止"; btnAuto.Background = _B(150, 80, 80);
                }
                else { autoTestTimer.Stop(); autoTestTimer = null; btnAuto.Content = "🔄 連續測試"; btnAuto.Background = _B(80, 80, 90); }
            };

            btnCorr.Click += (s, e) => OpenPositionCorrectionSettings();
            btnSaveCal.Click += (s, e) =>
            {
                updateRegion(); updateBounds(); updateClr();
                if (minimapTracker.MinimapRegion.Width <= 0) { WinMessageBox.Show("請先設定小地圖區域！", "錯誤", WinMBBtn.OK, WinMBIcon.Warning); return; }
                minimapTracker.SaveCalibration(CalibrationFilePath);
                lblStat.Text = "✅ 校準設定已儲存！"; lblStat.Foreground = _B(0, 255, 0); AddLog("💾 小地圖校準設定已儲存");
            };
            btnCloseCal.Click += (s, e) => win.Close();

            win.Closing += (s, e) => { autoTestTimer?.Stop(); syncTimer.Stop(); currentMinimapBitmap?.Dispose(); _boundarySetState = -1; };
            updateRegion(); updateBounds();

            rootSv.Content = main;
            win.Content = rootSv;
            win.ShowDialog();
        }

        // ═══════════════════════════════════════════════
        //  🎯 位置修正設定 (WPF)
        // ═══════════════════════════════════════════════

        private void OpenPositionCorrectionSettings()
        {
            // ── 顏色常數 ──
            var clrDim  = _B(160, 160, 165);
            var clrAccH = _B(255, 210, 100);
            var clrAccV = _B(100, 220, 255);
            var clrJump = _B(180, 255, 180);
            var clrRoi  = _B(255, 180, 80);

            var win = _CyWin("🎯 位置修正設定", 620, 870, true);
            var rootSv = new WCtl.ScrollViewer { VerticalScrollBarVisibility = WCtl.ScrollBarVisibility.Auto };
            var main = new WCtl.StackPanel { Margin = new Thickness(15, 12, 15, 12) };

            // ── 啟用 ──
            var chkEnabled = _Chk("啟用位置修正", positionCorrectionSettings.Enabled);
            chkEnabled.FontSize = 13; chkEnabled.FontWeight = FontWeights.Bold;
            main.Children.Add(chkEnabled);
            var lblModeDesc = _L("📌 比對腳本錄製座標，每隔設定秒數自動修正偏差", 10, _B(120, 200, 255));
            lblModeDesc.Margin = new Thickness(0, 2, 0, 6);
            main.Children.Add(lblModeDesc);

            // ── 按鍵設定 GroupBox ──
            var grpK = _Grp("⌨️ 移動按鍵 (點擊框後按下按鍵)");
            var grpKPanel = new WCtl.StackPanel();

            WinKeys[] cL = positionCorrectionSettings.GetEffectiveLeftKeys(), cR = positionCorrectionSettings.GetEffectiveRightKeys();
            WinKeys[] cU = positionCorrectionSettings.GetEffectiveUpKeys(),   cD = positionCorrectionSettings.GetEffectiveDownKeys();
            WinKeys[] cJ = positionCorrectionSettings.GetEffectiveClimbJumpKeys();

            var chkH = _Chk("水平", positionCorrectionSettings.EnableHorizontalCorrection, clrDim);
            var chkV = _Chk("垂直", positionCorrectionSettings.EnableVerticalCorrection, clrDim);

            // Key capture text boxes
            WCtl.TextBox _KeyTxt(WinKeys[] keys)
            {
                var t = _Txt(100, 25, PositionCorrector.KeysToDisplayString(keys));
                t.IsReadOnly = true; t.Tag = keys;
                t.Cursor = System.Windows.Input.Cursors.Arrow;
                return t;
            }
            var tL = _KeyTxt(cL); var tR = _KeyTxt(cR);
            var tU = _KeyTxt(cU); var tD = _KeyTxt(cD);
            var tJ = _KeyTxt(cJ);

            var lL = _L("左:", 11, clrDim); lL.VerticalAlignment = VerticalAlignment.Center; lL.Margin = new Thickness(4, 0, 2, 0);
            var lR = _L("右:", 11, clrDim); lR.VerticalAlignment = VerticalAlignment.Center; lR.Margin = new Thickness(8, 0, 2, 0);
            var lU = _L("上:", 11, clrDim); lU.VerticalAlignment = VerticalAlignment.Center; lU.Margin = new Thickness(4, 0, 2, 0);
            var lD = _L("下:", 11, clrDim); lD.VerticalAlignment = VerticalAlignment.Center; lD.Margin = new Thickness(8, 0, 2, 0);
            var lJ = _L("跳離鍵:", 11, clrRoi); lJ.VerticalAlignment = VerticalAlignment.Center;
            var lJHint = _L("爬繩時用此鍵+方向鍵跳離繩子", 9, _B(120, 120, 130));
            lJHint.VerticalAlignment = VerticalAlignment.Center; lJHint.Margin = new Thickness(6, 0, 0, 0);

            var rowH = _Row(chkH, lL, tL, lR, tR); rowH.Margin = new Thickness(0, 2, 0, 2);
            var rowV = _Row(chkV, lU, tU, lD, tD); rowV.Margin = new Thickness(0, 2, 0, 2);
            var rowJ = _Row(lJ, tJ, lJHint); rowJ.Margin = new Thickness(0, 2, 0, 2);
            grpKPanel.Children.Add(rowH);
            grpKPanel.Children.Add(rowV);
            grpKPanel.Children.Add(rowJ);
            grpK.Content = grpKPanel;
            main.Children.Add(grpK);

            // ── 按鍵捕獲邏輯 ──
            WCtl.TextBox? actKI = null;
            var capK = new HashSet<WinKeys>();
            DispatcherTimer? cTmr = null;

            Action resetCaptureTimer = () =>
            {
                cTmr?.Stop();
                var cur = actKI;
                cTmr = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                cTmr.Tick += (ts, te) =>
                {
                    cTmr.Stop();
                    if (actKI == cur && cur != null && capK.Count > 0)
                    {
                        cur.Tag = capK.ToArray();
                        cur.Text = PositionCorrector.KeysToDisplayString(capK.ToArray());
                        cur.Foreground = WMedia.Brushes.White;
                        actKI = null;
                    }
                };
                cTmr.Start();
            };

            Action<WCtl.TextBox> setupCK = (tb) =>
            {
                tb.PreviewMouseLeftButtonDown += (s, e) =>
                {
                    actKI = tb; capK.Clear();
                    tb.Text = "按下..."; tb.Foreground = _B(255, 255, 0);
                    e.Handled = true;
                };
                tb.LostFocus += (s, e) =>
                {
                    if (actKI != tb) return;
                    actKI = null;
                    if (capK.Count > 0) { tb.Tag = capK.ToArray(); tb.Text = PositionCorrector.KeysToDisplayString(capK.ToArray()); }
                    else tb.Text = tb.Tag is WinKeys[] ks ? PositionCorrector.KeysToDisplayString(ks) : "未設定";
                    tb.Foreground = WMedia.Brushes.White;
                };
            };
            setupCK(tL); setupCK(tR); setupCK(tU); setupCK(tD); setupCK(tJ);

            win.PreviewKeyDown += (s, e) =>
            {
                if (actKI == null) return;
                e.Handled = true;
                var wk = (WinKeys)KeyInterop.VirtualKeyFromKey(e.Key == Key.System ? e.SystemKey : e.Key);
                capK.Add(wk);
                if (System.Windows.Input.Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && wk != WinKeys.ControlKey) capK.Add(WinKeys.ControlKey);
                if (System.Windows.Input.Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && wk != WinKeys.ShiftKey) capK.Add(WinKeys.ShiftKey);
                if (System.Windows.Input.Keyboard.Modifiers.HasFlag(ModifierKeys.Alt) && wk != WinKeys.Menu) capK.Add(WinKeys.Menu);
                actKI.Text = PositionCorrector.KeysToDisplayString(capK.ToArray());
                resetCaptureTimer();
            };

            // ── 參數列 1：檢查間隔 | 容差 | 超時 ──
            var lChkInt = _L("檢查間隔:", 11, clrDim); lChkInt.VerticalAlignment = VerticalAlignment.Center;
            var nChkInt = _Num(40, positionCorrectionSettings.CorrectionCheckIntervalSec, 1, 60);
            var lChkSec = _L("秒", 11, clrDim); lChkSec.VerticalAlignment = VerticalAlignment.Center; lChkSec.Margin = new Thickness(2, 0, 8, 0);
            var lHTol = _L("容差 水平:", 11, clrDim); lHTol.VerticalAlignment = VerticalAlignment.Center;
            var nHTol = _Num(40, positionCorrectionSettings.HorizontalTolerance, 1, 50);
            var lVTol = _L("垂直:", 11, clrDim); lVTol.VerticalAlignment = VerticalAlignment.Center; lVTol.Margin = new Thickness(4, 0, 2, 0);
            var nVTol = _Num(40, positionCorrectionSettings.VerticalTolerance, 1, 50);
            var lTolPx = _L("px", 11, clrDim); lTolPx.VerticalAlignment = VerticalAlignment.Center; lTolPx.Margin = new Thickness(2, 0, 8, 0);
            var lTO = _L("超時:", 11, clrDim); lTO.VerticalAlignment = VerticalAlignment.Center;
            var nTO = _Num(68, positionCorrectionSettings.MaxCorrectionTimeMs, 1000, 30000);
            var lTOms = _L("ms", 11, clrDim); lTOms.VerticalAlignment = VerticalAlignment.Center;
            var row1 = _Row(lChkInt, nChkInt, lChkSec, lHTol, nHTol, lVTol, nVTol, lTolPx, lTO, nTO, lTOms);
            row1.Margin = new Thickness(0, 8, 0, 2);
            main.Children.Add(row1);

            // ── 參數列 2：按鍵間隔 ──
            var lHInt = _L("按鍵間隔 水平:", 11, clrAccH); lHInt.VerticalAlignment = VerticalAlignment.Center;
            var nHInt = _Num(58, (positionCorrectionSettings.HorizontalKeyIntervalMinMs + positionCorrectionSettings.HorizontalKeyIntervalMaxMs) / 2, 100, 5000);
            var lHIntMs = _L("ms", 11, clrAccH); lHIntMs.VerticalAlignment = VerticalAlignment.Center; lHIntMs.Margin = new Thickness(2, 0, 8, 0);
            var lVInt = _L("垂直:", 11, clrAccV); lVInt.VerticalAlignment = VerticalAlignment.Center;
            var nVInt = _Num(58, (positionCorrectionSettings.VerticalKeyIntervalMinMs + positionCorrectionSettings.VerticalKeyIntervalMaxMs) / 2, 50, 3000);
            var lVIntMs = _L("ms   (±150ms 自動浮動)", 11, clrDim); lVIntMs.VerticalAlignment = VerticalAlignment.Center;
            var row2 = _Row(lHInt, nHInt, lHIntMs, lVInt, nVInt, lVIntMs);
            row2.Margin = new Thickness(0, 4, 0, 2);
            main.Children.Add(row2);

            // ── 參數列 3：連跳 ──
            var lJump = _L("連跳 次數:", 11, clrJump); lJump.VerticalAlignment = VerticalAlignment.Center;
            var nJumpN = _Num(40, positionCorrectionSettings.ConsecutiveJumpCount, 1, 5);
            var lJumpN = _L("次    間隔:", 11, clrJump); lJumpN.VerticalAlignment = VerticalAlignment.Center; lJumpN.Margin = new Thickness(4, 0, 2, 0);
            var nJumpI = _Num(58, positionCorrectionSettings.ConsecutiveJumpIntervalMs, 50, 1000);
            var lJumpMs = _L("ms", 11, clrJump); lJumpMs.VerticalAlignment = VerticalAlignment.Center;
            var row3 = _Row(lJump, nJumpN, lJumpN, nJumpI, lJumpMs);
            row3.Margin = new Thickness(0, 4, 0, 4);
            main.Children.Add(row3);

            // ── ToolTips ──
            nHTol.ToolTip = "水平偏差在此範圍內停止修正";
            nVTol.ToolTip = "垂直偏差在此範圍內停止修正";
            nHInt.ToolTip = "水平修正按鍵間等待時間（程式自動套用 ±150ms 隨機浮動以避免被偵測）";
            nVInt.ToolTip = "垂直修正按鍵間等待時間（垂直通常設短一點）";
            nJumpN.ToolTip = "向上修正連跳次數（1=單跳）";
            nJumpI.ToolTip = "連跳之間的間隔（ms）";

            // ── 🔬 方向診斷 GroupBox ──
            var grpDiag = _Grp("🔬 方向診斷");
            grpDiag.Foreground = _B(255, 255, 0);
            var grpDiagPanel = new WCtl.StackPanel();

            var bDL = _Btn("←左", 55, 55, 65, 62, 28); bDL.FontSize = 11;
            var bDR = _Btn("→右", 55, 55, 65, 62, 28); bDR.FontSize = 11; bDR.Margin = new Thickness(4, 0, 0, 0);
            var bDU = _Btn("↑上", 55, 55, 65, 62, 28); bDU.FontSize = 11; bDU.Margin = new Thickness(4, 0, 0, 0);
            var bDD = _Btn("↓下", 55, 55, 65, 62, 28); bDD.FontSize = 11; bDD.Margin = new Thickness(4, 0, 0, 0);
            var lblDiag = _L("點按鈕 → 角色移一步 → 顯示座標變化", 9, clrDim);
            lblDiag.FontFamily = _monoFont; lblDiag.VerticalAlignment = VerticalAlignment.Center; lblDiag.Margin = new Thickness(8, 0, 0, 0);
            grpDiagPanel.Children.Add(_Row(bDL, bDR, bDU, bDD, lblDiag));
            grpDiag.Content = grpDiagPanel;
            main.Children.Add(grpDiag);

            // ── 📋 修正日誌 GroupBox ──
            var grpLog = _Grp("📋 修正日誌");
            var grpLogPanel = new WCtl.StackPanel();

            var txtLog = new WCtl.TextBox
            {
                Height = 100, IsReadOnly = true, TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = WCtl.ScrollBarVisibility.Visible,
                Background = _B(22, 22, 28), Foreground = _B(144, 238, 144),
                BorderBrush = _B(60, 60, 65), BorderThickness = new Thickness(1),
                FontFamily = _monoFont, FontSize = 9.5, CaretBrush = WMedia.Brushes.White,
                AcceptsReturn = true
            };
            grpLogPanel.Children.Add(txtLog);

            var btnTestCorr = _Btn("🧪 測試偵測", 160, 100, 0, 95, 26); btnTestCorr.FontSize = 11;
            var btnClearLog = _Btn("清除", 55, 55, 65, 50, 26); btnClearLog.FontSize = 11; btnClearLog.Margin = new Thickness(4, 0, 0, 0);
            var lblTestRes = _L("", 9.5, clrDim); lblTestRes.VerticalAlignment = VerticalAlignment.Center; lblTestRes.Margin = new Thickness(6, 0, 0, 0);
            var logBtnRow = _Row(btnTestCorr, btnClearLog, lblTestRes);
            logBtnRow.Margin = new Thickness(0, 4, 0, 0);
            grpLogPanel.Children.Add(logBtnRow);
            grpLog.Content = grpLogPanel;
            main.Children.Add(grpLog);

            // ── 🪢 爬繩偵測 GroupBox ──
            var grpRoi = _Grp("🪢 爬繩偵測 (框選角色爬繩姿態作為模板，在畫面中心搜尋比對)");
            grpRoi.Foreground = clrRoi;
            var grpRoiPanel = new WCtl.StackPanel();

            // 參數列
            var lRoiThr = _L("匹配閾值:", 11, clrDim); lRoiThr.VerticalAlignment = VerticalAlignment.Center;
            var nRoiThr = _Num(55, (int)(_climbingMatchThreshold * 100), 30, 99);
            var lRoiThrPct = _L("%", 11, clrDim); lRoiThrPct.VerticalAlignment = VerticalAlignment.Center; lRoiThrPct.Margin = new Thickness(2, 0, 4, 0);
            var btnApplyRoi = _Btn("套用", 55, 75, 100, 58, 26); btnApplyRoi.FontSize = 11;
            nRoiThr.ToolTip = "邊緣輪廓匹配分數需 ≥ 此百分比才判定為爬繩\n使用 NCC 演算法，專注角色輪廓不受背景影響\n建議 65~80%";
            var roiParamRow = _Row(lRoiThr, nRoiThr, lRoiThrPct, btnApplyRoi);
            roiParamRow.Margin = new Thickness(0, 2, 0, 4);
            grpRoiPanel.Children.Add(roiParamRow);

            // 操作按鈕列
            var btnCaptureTemplate = _Btn("🖱️ 框選爬繩模板", 130, 55, 0, 130, 26); btnCaptureTemplate.FontSize = 10;
            var btnSnapClimb  = _Btn("📸 爬繩快照", 100, 70, 0, 100, 26); btnSnapClimb.FontSize = 10; btnSnapClimb.Margin = new Thickness(4, 0, 0, 0);
            var btnSnapNormal = _Btn("📸 正常快照", 0, 70, 130, 100, 26); btnSnapNormal.FontSize = 10; btnSnapNormal.Margin = new Thickness(4, 0, 0, 0);
            var btnDetectNow  = _Btn("▶ 連續偵測", 70, 70, 0, 100, 26); btnDetectNow.FontSize = 10; btnDetectNow.Margin = new Thickness(4, 0, 0, 0);
            var btnOpenFolder = _Btn("📂", 45, 45, 55, 32, 26); btnOpenFolder.FontSize = 10; btnOpenFolder.Margin = new Thickness(4, 0, 0, 0);
            btnCaptureTemplate.ToolTip = "截取遊戲畫面 → 框選爬繩中的角色姿態 → 儲存為模板";
            btnOpenFolder.ToolTip = "開啟快照資料夾";
            var roiActionRow = _Row(btnCaptureTemplate, btnSnapClimb, btnSnapNormal, btnDetectNow, btnOpenFolder);
            roiActionRow.Margin = new Thickness(0, 0, 0, 4);
            grpRoiPanel.Children.Add(roiActionRow);

            // 模板預覽 + 匹配預覽 + 說明
            var imgTemplate = new WCtl.Image { Width = 70, Height = 85, Stretch = WMedia.Stretch.Uniform };
            var bdrTemplate = new WCtl.Border
            {
                Width = 70, Height = 85, Background = _B(25, 25, 30),
                BorderBrush = _B(80, 80, 85), BorderThickness = new Thickness(1),
                Child = imgTemplate
            };
            var lblTemplateTag = _L("模板", 8, _B(180, 180, 180));
            lblTemplateTag.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;

            var imgRoi = new WCtl.Image { Width = 70, Height = 85, Stretch = WMedia.Stretch.Uniform };
            var bdrRoi = new WCtl.Border
            {
                Width = 70, Height = 85, Background = _B(0, 0, 0),
                BorderBrush = _B(80, 80, 85), BorderThickness = new Thickness(1),
                Child = imgRoi, Margin = new Thickness(6, 0, 0, 0)
            };
            var lblRoiTag = _L("匹配區", 8, _B(180, 180, 180));
            lblRoiTag.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            lblRoiTag.Margin = new Thickness(6, 0, 0, 0);

            if (_climbingTemplate != null)
                imgTemplate.Source = _BmpToWpf(_climbingTemplate);

            var lblRoiStatus = _L(
                _climbingTemplate != null
                    ? $"✅ 模板已載入 ({_climbingTemplate.Width}x{_climbingTemplate.Height})\n點「▶ 連續偵測」即時預覽匹配結果\n⚡ 框選越貼近角色輪廓，辨識越準確"
                    : "⚠️ 尚未設定爬繩模板\n請讓角色爬繩後點「🖱️ 框選爬繩模板」\n⚡ 盡量緊貼角色輪廓框選，少含背景",
                9, clrDim);
            lblRoiStatus.TextWrapping = TextWrapping.Wrap; lblRoiStatus.VerticalAlignment = VerticalAlignment.Top;
            lblRoiStatus.Margin = new Thickness(8, 0, 0, 0);

            var templateCol = new WCtl.StackPanel(); templateCol.Children.Add(bdrTemplate); templateCol.Children.Add(lblTemplateTag);
            var roiCol = new WCtl.StackPanel(); roiCol.Children.Add(bdrRoi); roiCol.Children.Add(lblRoiTag);
            var roiPreviewRow = _Row(templateCol, roiCol, lblRoiStatus);
            grpRoiPanel.Children.Add(roiPreviewRow);
            grpRoi.Content = grpRoiPanel;
            main.Children.Add(grpRoi);

            // ── 底部按鈕 ──
            var btnSave  = _Btn("💾 儲存", 0, 130, 75, 90, 30); btnSave.FontSize = 12;
            var btnClose = _Btn("關閉", 70, 70, 75, 82, 30); btnClose.FontSize = 12; btnClose.Margin = new Thickness(8, 0, 0, 0); btnClose.BorderBrush = _B(100, 100, 105);
            var bottomRow = _Row(btnSave, btnClose);
            bottomRow.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
            bottomRow.Margin = new Thickness(0, 10, 0, 0);
            main.Children.Add(bottomRow);

            // ── Helper: 建立臨時修正器 ──
            Func<PositionCorrector> makeCorrector = () => new PositionCorrector
            {
                TargetWindow = targetWindowHandle,
                HorizontalTolerance = (int)_NV(nHTol), VerticalTolerance = (int)_NV(nVTol),
                Tolerance = (int)_NV(nHTol), SoftToleranceMin = (int)_NV(nHTol), SoftToleranceMax = (int)_NV(nVTol),
                MaxCorrectionTimeMs = (int)_NV(nTO),
                MoveLeftKeys  = tL.Tag as WinKeys[] ?? new[] { WinKeys.Left },
                MoveRightKeys = tR.Tag as WinKeys[] ?? new[] { WinKeys.Right },
                MoveUpKeys    = tU.Tag as WinKeys[] ?? new[] { WinKeys.Up },
                MoveDownKeys  = tD.Tag as WinKeys[] ?? new[] { WinKeys.Down },
                EnableHorizontalCorrection = chkH.IsChecked == true, EnableVerticalCorrection = chkV.IsChecked == true,
                InvertY = false,
                ExternalKeySender = SendKeyForCorrection,
                IsAnimationLockedCheck = () => IsAnimationLocked,
                HorizontalKeyIntervalMinMs = Math.Max(100, (int)_NV(nHInt) - 150),
                HorizontalKeyIntervalMaxMs = (int)_NV(nHInt) + 150,
                VerticalKeyIntervalMinMs   = Math.Max(50, (int)_NV(nVInt) - 150),
                VerticalKeyIntervalMaxMs   = (int)_NV(nVInt) + 150,
                ConsecutiveJumpCount = (int)_NV(nJumpN), ConsecutiveJumpIntervalMs = (int)_NV(nJumpI)
            };

            Action<string> appendLog = (msg) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    txtLog.AppendText(msg + Environment.NewLine);
                    txtLog.ScrollToEnd();
                }));
            };

            // ── 方向診斷事件 ──
            Action<string, Func<WinKeys[]>> doDiag = (dir, getKeys) =>
            {
                if (minimapTracker == null || !minimapTracker.IsCalibrated) { lblDiag.Text = "❌ 請先校準小地圖"; return; }
                lblDiag.Text = $"測試 {dir} 中..."; lblDiag.Foreground = _B(255, 255, 0);
                var c = makeCorrector();
                System.Threading.Tasks.Task.Run(() =>
                {
                    var r = c.DiagnoseDirection(minimapTracker, dir, getKeys());
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        lblDiag.Text = r.ToString();
                        lblDiag.Foreground = r.Error != null ? _B(255, 0, 0) : _B(0, 255, 0);
                        appendLog(r.ToString());
                    }));
                });
            };
            bDL.Click += (s, e) => doDiag("←左", () => tL.Tag as WinKeys[] ?? new[] { WinKeys.Left });
            bDR.Click += (s, e) => doDiag("→右", () => tR.Tag as WinKeys[] ?? new[] { WinKeys.Right });
            bDU.Click += (s, e) => doDiag("↑上", () => tU.Tag as WinKeys[] ?? new[] { WinKeys.Up });
            bDD.Click += (s, e) => doDiag("↓下", () => tD.Tag as WinKeys[] ?? new[] { WinKeys.Down });
            btnClearLog.Click += (s, e) => txtLog.Clear();

            // ── 測試偵測事件 ──
            btnTestCorr.Click += (s, e) =>
            {
                if (minimapTracker == null || !minimapTracker.IsCalibrated)
                { lblTestRes.Text = "❌ 小地圖未校準"; lblTestRes.Foreground = _B(255, 0, 0); return; }
                btnTestCorr.IsEnabled = false; lblTestRes.Text = "偵測中..."; lblTestRes.Foreground = _B(255, 255, 0);
                System.Threading.Tasks.Task.Run(() =>
                {
                    var (cx, cy, ok) = minimapTracker.ReadPosition();
                    string msg = ok ? $"✅ 當前位置 ({cx},{cy}) — 播放時將比對腳本座標" : "⚠️ 偵測失敗";
                    if (ok) appendLog(msg);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        btnTestCorr.IsEnabled = true;
                        lblTestRes.Text = msg;
                        lblTestRes.Foreground = ok ? _B(0, 255, 0) : _B(255, 165, 0);
                    }));
                });
            };

            // ── ROI 爬繩偵測事件 ──
            Action syncRoiParams = () => { _climbingMatchThreshold = (int)_NV(nRoiThr) / 100.0; };

            btnApplyRoi.Click += (s, e) =>
            {
                syncRoiParams();
                lblRoiStatus.Text = $"✅ 參數已套用 (閾值={_climbingMatchThreshold:P0})";
                lblRoiStatus.Foreground = _B(0, 255, 0);
            };

            btnCaptureTemplate.Click += (s, e) =>
            {
                if (minimapTracker == null)
                { lblRoiStatus.Text = "❌ 小地圖追蹤器未初始化"; lblRoiStatus.Foreground = _B(255, 0, 0); return; }
                using (var screenshot = minimapTracker.CaptureFullWindow())
                {
                    if (screenshot == null)
                    { lblRoiStatus.Text = "❌ 無法截取遊戲視窗"; lblRoiStatus.Foreground = _B(255, 0, 0); return; }
                    win.Hide(); var region = RegionSelector.SelectRegion(screenshot); win.Show();
                    if (region.HasValue && region.Value.Width > 5 && region.Value.Height > 5)
                    {
                        try
                        {
                            using var cropped = new Bitmap(region.Value.Width, region.Value.Height, PixelFormat.Format24bppRgb);
                            using (var g = Graphics.FromImage(cropped))
                                g.DrawImage(screenshot, new Rectangle(0, 0, cropped.Width, cropped.Height), region.Value, GraphicsUnit.Pixel);
                            string dir = Path.GetDirectoryName(ClimbingTemplatePath)!;
                            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                            cropped.Save(ClimbingTemplatePath, ImageFormat.Png);
                            _climbingTemplate?.Dispose();
                            _climbingTemplate = new Bitmap(cropped);
                            _templateEdgeCache = null;
                            _templateHueHistogram = null;
                            imgTemplate.Source = _BmpToWpf(cropped);
                            lblRoiStatus.Text = $"✅ 爬繩模板已儲存 ({region.Value.Width}x{region.Value.Height})\n路徑: {ClimbingTemplatePath}\n請用「爬繩快照」和「正常快照」驗證效果";
                            lblRoiStatus.Foreground = _B(0, 255, 0);
                        }
                        catch (Exception ex)
                        { lblRoiStatus.Text = $"❌ 儲存模板失敗: {ex.Message}"; lblRoiStatus.Foreground = _B(255, 0, 0); }
                    }
                }
            };

            Action<string> doSnapshot = (label) =>
            {
                syncRoiParams();
                if (_climbingTemplate == null)
                { lblRoiStatus.Text = "❌ 尚未設定爬繩模板，請先框選"; lblRoiStatus.Foreground = _B(255, 0, 0); return; }
                if (minimapTracker == null || !minimapTracker.IsCalibrated)
                { lblRoiStatus.Text = "❌ 小地圖未校準"; lblRoiStatus.Foreground = _B(255, 0, 0); return; }
                lblRoiStatus.Text = "截圖中..."; lblRoiStatus.Foreground = _B(255, 255, 0);
                System.Threading.Tasks.Task.Run(() =>
                {
                    var (result, path) = SaveRoiDebugSnapshot(label);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (result.CaptureBitmap != null) imgRoi.Source = _BmpToWpf(result.CaptureBitmap);
                        double scorePct = result.MatchScore * 100, thrPct = result.Threshold * 100;
                        bool expectHit = label == "CLIMBING";
                        string verdict = result.IsClimbing == expectHit
                            ? $"✅ {(expectHit ? "爬繩" : "正常")}判定正確"
                            : (expectHit ? $"⚠️ 爬繩未命中 → 降低匹配閾值或重新框選模板" : $"⚠️ 正常狀態誤判 → 提高匹配閾值或重新框選模板");
                        double rawPctSnap = result.RawNccScore * 100, colorPctSnap = result.ColorSimilarity * 100;
                        string posInfo = result.IsClimbing ? $"找到位置: ({result.FoundX},{result.FoundY})" : "未找到匹配";
                        string distInfoSnap = result.CenterDist > 0 ? $" | 中心距{result.CenterDist:F0}px" : "";
                        string colorInfoSnap = $" | 色彩{colorPctSnap:F0}%";
                        string diagInfo = result.GameW > 0 ? $"\n視窗: {result.GameW}x{result.GameH} | ROI: ({result.RoiLeft},{result.RoiTop}) {result.RoiW}x{result.RoiH}" : "";
                        string errInfo = result.Error != null ? $"\n⚠️ {result.Error}" : "";
                        lblRoiStatus.Text = $"{verdict}\n加權: {scorePct:F1}% 原始: {rawPctSnap:F1}% 色彩: {colorPctSnap:F0}% (閾值 {thrPct:F0}%)\n{posInfo}{distInfoSnap}{colorInfoSnap}{diagInfo}{errInfo}";
                        lblRoiStatus.Foreground = result.IsClimbing == expectHit ? _B(0, 255, 0) : _B(255, 165, 0);
                    }));
                });
            };
            btnSnapClimb.Click  += (s, e) => doSnapshot("CLIMBING");
            btnSnapNormal.Click += (s, e) => doSnapshot("NORMAL");

            // ★ 連續偵測計時器
            DispatcherTimer? climbDetectTimer = null;
            bool climbDetectRunning = false;

            btnDetectNow.Click += (s, e) =>
            {
                if (!climbDetectRunning)
                {
                    syncRoiParams();
                    if (_climbingTemplate == null) { lblRoiStatus.Text = "❌ 尚未設定爬繩模板"; lblRoiStatus.Foreground = _B(255, 0, 0); return; }
                    if (minimapTracker == null || !minimapTracker.IsCalibrated) { lblRoiStatus.Text = "❌ 小地圖未校準"; lblRoiStatus.Foreground = _B(255, 0, 0); return; }
                    climbDetectRunning = true;
                    btnDetectNow.Content = "⏹ 停止偵測"; btnDetectNow.Background = _B(160, 50, 50);
                    climbDetectTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                    climbDetectTimer.Tick += (ts, te) =>
                    {
                        if (!climbDetectRunning) return;
                        System.Threading.Tasks.Task.Run(() =>
                        {
                            var r = ScanClimbingRoi();
                            try
                            {
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    if (!climbDetectRunning) return;
                                    if (r.CaptureBitmap != null) imgRoi.Source = _BmpToWpf(r.CaptureBitmap);
                                    double scorePct = r.MatchScore * 100, rawPct = r.RawNccScore * 100;
                                    string pos = r.IsClimbing ? $" @({r.FoundX},{r.FoundY})" : "";
                                    string errInfo = r.Error != null ? $"\n⚠️ {r.Error}" : "";
                                    string sizeInfo = r.GameW > 0 ? $" | 視窗{r.GameW}x{r.GameH} ROI({r.RoiLeft},{r.RoiTop} {r.RoiW}x{r.RoiH})" : "";
                                    string distInfo = r.CenterDist > 0 ? $" 中心距{r.CenterDist:F0}px" : "";
                                    string colorInfo = $" 色彩{r.ColorSimilarity * 100:F0}%";
                                    lblRoiStatus.Text = $"🔴 加權{scorePct:F1}% 原始{rawPct:F1}% (閾值{r.Threshold * 100:F0}%){pos}{distInfo}{colorInfo} → {(r.IsClimbing ? "🪢 爬繩中" : "🏃 正常")}{sizeInfo}{errInfo}";
                                    lblRoiStatus.Foreground = r.IsClimbing ? clrRoi : _B(0, 255, 0);
                                }));
                            }
                            catch { /* window closed */ }
                        });
                    };
                    climbDetectTimer.Start();
                }
                else
                {
                    climbDetectRunning = false;
                    climbDetectTimer?.Stop(); climbDetectTimer = null;
                    btnDetectNow.Content = "▶ 連續偵測"; btnDetectNow.Background = _B(70, 70, 0);
                    lblRoiStatus.Text = "⏹ 連續偵測已停止"; lblRoiStatus.Foreground = clrDim;
                }
            };

            win.Closing += (s, e) =>
            {
                climbDetectRunning = false;
                climbDetectTimer?.Stop();
            };

            btnOpenFolder.Click += (s, e) =>
            {
                Directory.CreateDirectory(RoiDebugDir);
                System.Diagnostics.Process.Start("explorer.exe", RoiDebugDir);
            };

            // ── 儲存事件 ──
            btnSave.Click += (s, e) =>
            {
                int hIntVal = (int)_NV(nHInt);
                int vIntVal = (int)_NV(nVInt);

                positionCorrectionSettings.Enabled = chkEnabled.IsChecked == true;
                positionCorrectionSettings.UseDeviationMode = true;
                positionCorrectionSettings.HorizontalTolerance = (int)_NV(nHTol);
                positionCorrectionSettings.VerticalTolerance   = (int)_NV(nVTol);
                positionCorrectionSettings.SoftToleranceMin = (int)_NV(nHTol);
                positionCorrectionSettings.SoftToleranceMax = (int)_NV(nVTol);
                positionCorrectionSettings.Tolerance = (int)_NV(nHTol);
                positionCorrectionSettings.MaxCorrectionTimeMs = (int)_NV(nTO);
                positionCorrectionSettings.InvertY = false;
                positionCorrectionSettings.EnableHorizontalCorrection = chkH.IsChecked == true;
                positionCorrectionSettings.EnableVerticalCorrection   = chkV.IsChecked == true;
                positionCorrectionSettings.MoveLeftKeys  = PositionCorrector.KeysToIntArray(tL.Tag as WinKeys[] ?? new[] { WinKeys.Left });
                positionCorrectionSettings.MoveRightKeys = PositionCorrector.KeysToIntArray(tR.Tag as WinKeys[] ?? new[] { WinKeys.Right });
                positionCorrectionSettings.MoveUpKeys    = PositionCorrector.KeysToIntArray(tU.Tag as WinKeys[] ?? new[] { WinKeys.Up });
                positionCorrectionSettings.MoveDownKeys  = PositionCorrector.KeysToIntArray(tD.Tag as WinKeys[] ?? new[] { WinKeys.Down });
                positionCorrectionSettings.ClimbEscapeJumpKeys = PositionCorrector.KeysToIntArray(tJ.Tag as WinKeys[] ?? new[] { WinKeys.Alt });
                var lk = tL.Tag as WinKeys[]; var rk = tR.Tag as WinKeys[]; var uk = tU.Tag as WinKeys[]; var dk = tD.Tag as WinKeys[];
                positionCorrectionSettings.MoveLeftKey  = lk?.Length > 0 ? (int)lk[0] : (int)WinKeys.Left;
                positionCorrectionSettings.MoveRightKey = rk?.Length > 0 ? (int)rk[0] : (int)WinKeys.Right;
                positionCorrectionSettings.MoveUpKey    = uk?.Length > 0 ? (int)uk[0] : (int)WinKeys.Up;
                positionCorrectionSettings.MoveDownKey  = dk?.Length > 0 ? (int)dk[0] : (int)WinKeys.Down;
                positionCorrectionSettings.CorrectionCheckIntervalSec = (int)_NV(nChkInt);
                positionCorrectionSettings.HorizontalKeyIntervalMinMs = Math.Max(100, hIntVal - 150);
                positionCorrectionSettings.HorizontalKeyIntervalMaxMs = hIntVal + 150;
                positionCorrectionSettings.KeyIntervalMinMs = positionCorrectionSettings.HorizontalKeyIntervalMinMs;
                positionCorrectionSettings.KeyIntervalMaxMs = positionCorrectionSettings.HorizontalKeyIntervalMaxMs;
                positionCorrectionSettings.VerticalKeyIntervalMinMs   = Math.Max(50,  vIntVal - 150);
                positionCorrectionSettings.VerticalKeyIntervalMaxMs   = vIntVal + 150;
                positionCorrectionSettings.ConsecutiveJumpCount       = (int)_NV(nJumpN);
                positionCorrectionSettings.ConsecutiveJumpIntervalMs  = (int)_NV(nJumpI);
                positionCorrectionSettings.ConsecutiveJumpThreshold   = 0;
                positionCorrectionSettings.MaxStepsPerCorrection      = 0;
                positionCorrectionSettings.MaxCorrectionsPerLoop      = 0;
                syncRoiParams();
                _correctionSettingsChanged = true;
                SaveAppSettings();
                AddLog("💾 [位置修正設定已儲存，下個循環生效]");
                win.Close();
            };

            btnClose.Click += (s, e) => win.Close();

            rootSv.Content = main;
            win.Content = rootSv;
            win.ShowDialog();
        }

        // ═══════════════════════════════════════════════
        //  📝 腳本編輯器 (WPF)
        // ═══════════════════════════════════════════════
        private void OpenFoldedEventEditor_Wpf()
        {
            AddLog("正在開啟事件編輯器...");
            var foldedActions = ConvertToFoldedActions();

            var win = _CyWin($"腳本編輯器 ({recordedEvents.Count} 個事件 → {foldedActions.Count} 個動作)", 580, 660);
            win.ResizeMode = ResizeMode.CanResize;
            win.MinWidth = 500; win.MinHeight = 500;

            var root = new WCtl.DockPanel { Margin = new Thickness(10) };

            // ── 頂部提示 ──
            var lblHint = _L("💡 雙擊修改按鍵/時間/持續時間 | Delete 刪除 | ➕插入按鍵", 9, _B(0, 255, 255));
            WCtl.DockPanel.SetDock(lblHint, WCtl.Dock.Top);
            lblHint.Margin = new Thickness(0, 0, 0, 6);
            root.Children.Add(lblHint);

            // ── 底部按鈕列 ──
            var btnPanel = new WCtl.WrapPanel { Margin = new Thickness(0, 8, 0, 0) };
            WCtl.DockPanel.SetDock(btnPanel, WCtl.Dock.Bottom);

            var btnClear  = _Btn("🗑️ 清空全部", 150, 60, 60, 100, 32);
            var btnInsert = _Btn("➕ 插入按鍵", 80, 120, 180, 100, 32); btnInsert.Margin = new Thickness(6, 0, 0, 0);
            var btnDelete = _Btn("🗑️ 刪除選取", 180, 80, 60, 100, 32); btnDelete.Margin = new Thickness(6, 0, 0, 0);
            var btnSave   = _Btn("💾 儲存並關閉", 50, 150, 80, 120, 32); btnSave.Margin = new Thickness(6, 0, 0, 0);
            var btnCancel = _Btn("取消", 100, 100, 100, 65, 32); btnCancel.Margin = new Thickness(6, 0, 0, 0);
            var lblCount  = _L($"動作數: {foldedActions.Count}", 11, _B(0, 255, 255));
            lblCount.Margin = new Thickness(12, 0, 0, 0);
            lblCount.VerticalAlignment = VerticalAlignment.Center;
            btnPanel.Children.Add(btnClear);
            btnPanel.Children.Add(btnInsert);
            btnPanel.Children.Add(btnDelete);
            btnPanel.Children.Add(btnSave);
            btnPanel.Children.Add(btnCancel);
            btnPanel.Children.Add(lblCount);
            root.Children.Add(btnPanel);

            // ── 列表 (WPF ListView + GridView) ──
            var lv = new WCtl.ListView
            {
                Background = _B(25, 25, 30),
                Foreground = WMedia.Brushes.White,
                FontFamily = _monoFont,
                FontSize = 12,
                BorderBrush = _B(60, 60, 65),
                BorderThickness = new Thickness(1)
            };
            var gv = new WCtl.GridView();
            gv.Columns.Add(new WCtl.GridViewColumn { Header = "按鍵",    Width = 80,  DisplayMemberBinding = new System.Windows.Data.Binding("KeyName") });
            gv.Columns.Add(new WCtl.GridViewColumn { Header = "重複",    Width = 50,  DisplayMemberBinding = new System.Windows.Data.Binding("RepeatText") });
            gv.Columns.Add(new WCtl.GridViewColumn { Header = "持續時間", Width = 100, DisplayMemberBinding = new System.Windows.Data.Binding("DurationText") });
            gv.Columns.Add(new WCtl.GridViewColumn { Header = "狀態",    Width = 60,  DisplayMemberBinding = new System.Windows.Data.Binding("StatusText") });
            gv.Columns.Add(new WCtl.GridViewColumn { Header = "時間點",  Width = 90,  DisplayMemberBinding = new System.Windows.Data.Binding("TimeText") });
            lv.View = gv;
            root.Children.Add(lv);

            // ── 列表資料輔助類 ──
            Action refreshList = () =>
            {
                lv.Items.Clear();
                foreach (var a in foldedActions)
                {
                    lv.Items.Add(new
                    {
                        a.KeyName,
                        RepeatText = a.RepeatCount > 1 ? $"x{a.RepeatCount}" : "-",
                        DurationText = a.Duration >= 1.0 ? $"{a.Duration:F2}秒" : $"{a.Duration * 1000:F0}ms",
                        StatusText = a.IsReleased ? "完成" : "按住中",
                        TimeText = $"{a.PressTime:F3}s",
                        Action = a
                    });
                }
                win.Title = $"腳本編輯器 ({foldedActions.Count} 個動作)";
                lblCount.Text = $"動作數: {foldedActions.Count}";
            };
            refreshList();

            // ── 雙擊編輯 ──
            lv.MouseDoubleClick += (s, args) =>
            {
                if (lv.SelectedItem == null) return;
                dynamic sel = lv.SelectedItem;
                FoldedKeyAction action = sel.Action;

                var ew = _CyWin($"編輯 {action.KeyName}", 370, 310);
                var ep = new WCtl.StackPanel { Margin = new Thickness(14) };

                // 按鍵
                ep.Children.Add(_L("按鍵:"));
                var txtKey = _Txt(200, 28, action.KeyName, true);
                txtKey.Tag = (WinKeys?)action.KeyCode;
                txtKey.Cursor = System.Windows.Input.Cursors.Arrow;
                var keyCapturing = false;
                txtKey.MouseDown += (ts, te) => { txtKey.Text = "按下新按鍵..."; txtKey.Foreground = WMedia.Brushes.Yellow; keyCapturing = true; };
                ew.PreviewKeyDown += (ts, te) =>
                {
                    if (!keyCapturing) return;
                    te.Handled = true;
                    var wfKey = (WinKeys)KeyInterop.VirtualKeyFromKey(te.Key == Key.System ? te.SystemKey : te.Key);
                    txtKey.Tag = (WinKeys?)wfKey;
                    txtKey.Text = GetKeyDisplayName(wfKey);
                    txtKey.Foreground = _B(144, 238, 144);
                    keyCapturing = false;
                };
                ep.Children.Add(txtKey);

                // 時間點
                ep.Children.Add(_L("時間點 (秒):", 12)); ep.Children.Add(_Num(130, Math.Max(0, action.PressTime), 0, 99999, 3));
                // 持續時間
                ep.Children.Add(_L("持續時間 (秒):", 12)); ep.Children.Add(_Num(130, Math.Max(0.01, action.Duration), 0.01, 9999, 3));
                // 重複次數
                ep.Children.Add(_L("重複次數:", 12)); ep.Children.Add(_Num(80, action.RepeatCount, 1, 9999));

                ep.Children.Add(_L($"原始: {action.KeyName}, {action.PressTime:F3}s, {action.Duration:F3}秒, x{action.RepeatCount}", 10, _B(128, 128, 128)));

                var ebtn = _Row(
                    _Btn("確定", 60, 140, 80), _Btn("取消", 100, 100, 100));
                ebtn.Margin = new Thickness(0, 10, 0, 0);
                ((WCtl.Button)ebtn.Children[1]).Margin = new Thickness(8, 0, 0, 0);
                ep.Children.Add(ebtn);

                bool saved = false;
                ((WCtl.Button)ebtn.Children[0]).Click += (ts, te) => { saved = true; ew.Close(); };
                ((WCtl.Button)ebtn.Children[1]).Click += (ts, te) => ew.Close();

                ew.Content = ep;
                ew.ShowDialog();

                if (saved)
                {
                    if (txtKey.Tag is WinKeys nk)
                    {
                        action.KeyCode = nk;
                        action.KeyName = GetKeyDisplayName(nk);
                    }
                    // ep children: Label, TextBox(key), Label, Num(time), Label, Num(dur), Label, Num(repeat), Label(info), Row(buttons)
                    double newTime = _NV((WCtl.TextBox)ep.Children[3]);
                    double newDuration = _NV((WCtl.TextBox)ep.Children[5]);
                    int newRepeat = (int)_NV((WCtl.TextBox)ep.Children[7]);
                    action.PressTime = newTime;
                    action.ReleaseTime = newTime + newDuration;
                    action.RepeatCount = Math.Max(1, newRepeat);
                    foldedActions.Sort((a, b) => a.PressTime.CompareTo(b.PressTime));
                    refreshList();
                    AddLog($"✅ 已修改 {action.KeyName}: @{newTime:F3}s {newDuration:F3}秒 x{newRepeat}");
                }
            };

            // ── Delete 鍵刪除 ──
            lv.KeyDown += (s, args) =>
            {
                if (args.Key == Key.Delete && lv.SelectedItem != null)
                {
                    dynamic sel = lv.SelectedItem;
                    FoldedKeyAction action = sel.Action;
                    if (WinMessageBox.Show($"確定刪除 {action.KeyName}？", "確認", WinMBBtn.YesNo, WinMBIcon.Warning) == WinDR.Yes)
                    {
                        foldedActions.Remove(action);
                        refreshList();
                        AddLog($"✅ 已刪除 {action.KeyName}");
                    }
                }
            };

            // ── 清空全部 ──
            btnClear.Click += (s, args) =>
            {
                if (WinMessageBox.Show($"確定清空全部 {foldedActions.Count} 個動作？", "確認", WinMBBtn.YesNo, WinMBIcon.Warning) == WinDR.Yes)
                {
                    foldedActions.Clear();
                    refreshList();
                    AddLog("✅ 已清空所有動作");
                }
            };

            // ── 刪除選取 ──
            btnDelete.Click += (s, args) =>
            {
                if (lv.SelectedItem == null) { WinMessageBox.Show("請先選取要刪除的動作"); return; }
                dynamic sel = lv.SelectedItem;
                FoldedKeyAction action = sel.Action;
                if (WinMessageBox.Show($"確定刪除 {action.KeyName}？", "確認", WinMBBtn.YesNo, WinMBIcon.Warning) == WinDR.Yes)
                {
                    foldedActions.Remove(action);
                    refreshList();
                    AddLog($"✅ 已刪除 {action.KeyName}");
                }
            };

            // ── 插入按鍵 ──
            btnInsert.Click += (s, args) =>
            {
                var iw = _CyWin("插入按鍵事件", 370, 310);
                var ip = new WCtl.StackPanel { Margin = new Thickness(14) };

                ip.Children.Add(_L("按鍵（點擊框後按鍵）:"));
                var txtIKey = _Txt(290, 28, "點擊後按下按鍵...");
                txtIKey.IsReadOnly = true; txtIKey.Cursor = System.Windows.Input.Cursors.Arrow;
                txtIKey.Foreground = WMedia.Brushes.Yellow;
                txtIKey.Tag = (WinKeys?)null;
                var iCapture = false;
                txtIKey.MouseDown += (ts, te) => { txtIKey.Text = "按下按鍵..."; txtIKey.Foreground = WMedia.Brushes.Yellow; iCapture = true; };
                iw.PreviewKeyDown += (ts, te) =>
                {
                    if (!iCapture) return;
                    te.Handled = true;
                    var wfKey = (WinKeys)KeyInterop.VirtualKeyFromKey(te.Key == Key.System ? te.SystemKey : te.Key);
                    txtIKey.Tag = (WinKeys?)wfKey;
                    txtIKey.Text = wfKey.ToString();
                    txtIKey.Foreground = WMedia.Brushes.White;
                    iCapture = false;
                };
                ip.Children.Add(txtIKey);

                ip.Children.Add(_L("持續時間 (秒):", 12)); ip.Children.Add(_Num(100, 0.1, 0.01, 9999, 3));
                ip.Children.Add(_L("重複次數:", 12)); ip.Children.Add(_Num(80, 1, 1, 9999));

                ip.Children.Add(_L("插入位置:", 12));
                var rdoAfter = new WCtl.RadioButton { Content = "選取項目之後", IsChecked = true, Foreground = _B(200, 200, 200), FontFamily = _cyFont, FontSize = 12, Margin = new Thickness(0, 2, 0, 0) };
                var rdoEnd = new WCtl.RadioButton { Content = "末尾", Foreground = _B(200, 200, 200), FontFamily = _cyFont, FontSize = 12 };
                ip.Children.Add(rdoAfter);
                ip.Children.Add(rdoEnd);

                var ibtn = _Row(_Btn("插入", 50, 150, 80), _Btn("取消", 100, 100, 100));
                ibtn.Margin = new Thickness(0, 10, 0, 0);
                ((WCtl.Button)ibtn.Children[1]).Margin = new Thickness(8, 0, 0, 0);
                ip.Children.Add(ibtn);

                ((WCtl.Button)ibtn.Children[0]).Click += (ts, te) =>
                {
                    if (txtIKey.Tag == null) { WinMessageBox.Show("請先按下要插入的按鍵"); return; }
                    WinKeys key = (WinKeys)txtIKey.Tag;
                    // ip children: Label, TextBox(key), Label, Num(dur), Label, Num(repeat), Label, RadioAfter, RadioEnd, Row
                    double duration = _NV((WCtl.TextBox)ip.Children[3]);
                    int repeat = (int)_NV((WCtl.TextBox)ip.Children[5]);

                    double insertTime; int insertIndex;
                    if (rdoAfter.IsChecked == true && lv.SelectedItem != null)
                    {
                        dynamic selA = lv.SelectedItem;
                        FoldedKeyAction selAction = selA.Action;
                        insertIndex = foldedActions.IndexOf(selAction) + 1;
                        insertTime = selAction.ReleaseTime + 0.01;
                    }
                    else
                    {
                        insertIndex = foldedActions.Count;
                        insertTime = foldedActions.Count > 0 ? foldedActions.Last().ReleaseTime + 0.01 : 0;
                    }

                    double totalInserted = duration + 0.01;
                    for (int i = insertIndex; i < foldedActions.Count; i++)
                    {
                        foldedActions[i].PressTime += totalInserted;
                        foldedActions[i].ReleaseTime += totalInserted;
                    }

                    foldedActions.Insert(insertIndex, new FoldedKeyAction
                    {
                        KeyCode = key,
                        KeyName = GetKeyDisplayName(key),
                        PressTime = insertTime,
                        ReleaseTime = insertTime + duration,
                        RepeatCount = repeat,
                        IsReleased = true
                    });
                    refreshList();
                    AddLog($"✅ 已插入 {GetKeyDisplayName(key)} @{insertTime:F3}s (持續{duration}s x{repeat})");
                    iw.Close();
                };
                ((WCtl.Button)ibtn.Children[1]).Click += (ts, te) => iw.Close();

                iw.Content = ip;
                iw.ShowDialog();
            };

            // ── 儲存並關閉 ──
            btnSave.Click += (s, args) =>
            {
                RebuildEventsFromActions(foldedActions);
                lblRecordingStatus.Text = $"已編輯 | 事件數: {recordedEvents.Count}";
                AddLog($"✅ 已儲存 {recordedEvents.Count} 個事件");
                win.Close();
            };

            // ── 取消 ──
            btnCancel.Click += (s, args) => win.Close();

            win.Content = root;
            win.ShowDialog();
            UpdateUI();
        }
    }
}

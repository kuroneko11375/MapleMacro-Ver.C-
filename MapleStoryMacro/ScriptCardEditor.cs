using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 卡片式腳本編輯器表單 (簡化版)
    /// </summary>
    public class ScriptCardEditor : Form
    {
        private List<ScriptCard> cards;
        private ListBox lstCards;
        private TextBox txtJson;
        private Label lblStatus;
        
        public List<ScriptCard> Cards => cards;
        public bool HasChanges { get; private set; } = false;

        public ScriptCardEditor(List<ScriptCard> existingCards)
        {
            // 深拷貝卡片
            cards = existingCards.Select(c => new ScriptCard
            {
                Id = c.Id,
                Type = c.Type,
                Key = c.Key,
                Value = c.Value,
                IntervalMs = c.IntervalMs,
                RandomJitterMs = c.RandomJitterMs,
                Note = c.Note,
                Enabled = c.Enabled
            }).ToList();

            InitializeForm();
            RefreshList();
        }

        private void InitializeForm()
        {
            this.Text = "🎴 卡片式腳本編輯器";
            this.Size = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft JhengHei", 9);
            this.BackColor = Color.FromArgb(30, 30, 35);
            this.ForeColor = Color.White;

            // === 左側面板 ===
            var pnlLeft = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(500, 590),
                BackColor = Color.FromArgb(35, 35, 40)
            };

            // 工具列按鈕
            var btnAddWait = CreateBtn("⏱️Wait", 5, 5, Color.FromArgb(70, 130, 180));
            btnAddWait.Click += (s, e) => AddCard(CardType.Wait);
            
            var btnAddClick = CreateBtn("🖱️Click", 85, 5, Color.FromArgb(60, 179, 113));
            btnAddClick.Click += (s, e) => AddCard(CardType.Click);
            
            var btnAddSpam = CreateBtn("🔥Spam", 165, 5, Color.FromArgb(255, 99, 71));
            btnAddSpam.Click += (s, e) => AddCard(CardType.Spam);
            
            var btnAddHold = CreateBtn("⏸️Hold", 245, 5, Color.FromArgb(255, 165, 0));
            btnAddHold.Click += (s, e) => AddCard(CardType.Hold);

            var btnAddPos = CreateBtn("📍位置修正", 325, 5, Color.FromArgb(150, 80, 200));
            btnAddPos.Click += (s, e) => AddCard(CardType.PositionCorrect);

            var btnEdit = CreateBtn("✏️編輯", 420, 5, Color.FromArgb(100, 100, 120));
            btnEdit.Click += (s, e) => EditSelected();

            var btnDelete = CreateBtn("🗑️刪除", 505, 5, Color.FromArgb(180, 60, 60));
            btnDelete.Click += (s, e) => DeleteSelected();

            // 卡片列表
            lstCards = new ListBox
            {
                Location = new Point(5, 40),
                Size = new Size(490, 420),
                BackColor = Color.FromArgb(25, 25, 30),
                ForeColor = Color.White,
                Font = new Font("Consolas", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            lstCards.DoubleClick += (s, e) => EditSelected();

            // 移動按鈕
            var btnUp = CreateBtn("▲上移", 5, 465, Color.FromArgb(80, 80, 100));
            btnUp.Click += (s, e) => MoveCard(-1);
            
            var btnDown = CreateBtn("▼下移", 85, 465, Color.FromArgb(80, 80, 100));
            btnDown.Click += (s, e) => MoveCard(1);

            var btnAddJitter = CreateBtn("加擾動", 170, 465, Color.FromArgb(100, 80, 150));
            btnAddJitter.Click += (s, e) => AddJitterToAll();

            var btnClear = CreateBtn("清空", 250, 465, Color.FromArgb(150, 60, 60));
            btnClear.Click += (s, e) => ClearAll();

            // 狀態
            lblStatus = new Label
            {
                Location = new Point(5, 500),
                Size = new Size(490, 40),
                ForeColor = Color.Cyan,
                Text = "就緒"
            };

            // 儲存/取消
            var btnSave = new Button
            {
                Text = "💾 儲存變更",
                Location = new Point(5, 545),
                Size = new Size(120, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 150, 80),
                ForeColor = Color.White
            };
            btnSave.Click += (s, e) => { HasChanges = true; DialogResult = DialogResult.OK; Close(); };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(130, 545),
                Size = new Size(80, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(100, 100, 100),
                ForeColor = Color.White
            };
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            pnlLeft.Controls.AddRange(new Control[] {
                btnAddWait, btnAddClick, btnAddSpam, btnAddHold, btnEdit, btnDelete,
                lstCards, btnUp, btnDown, btnAddJitter, btnClear,
                lblStatus, btnSave, btnCancel
            });

            // === 右側面板 (JSON) ===
            var pnlRight = new Panel
            {
                Location = new Point(520, 10),
                Size = new Size(360, 590),
                BackColor = Color.FromArgb(35, 35, 40)
            };

            var lblJson = new Label
            {
                Text = "📋 JSON (可複製給 AI)",
                Location = new Point(5, 5),
                AutoSize = true,
                ForeColor = Color.White
            };

            txtJson = new TextBox
            {
                Location = new Point(5, 30),
                Size = new Size(350, 480),
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                BackColor = Color.FromArgb(20, 20, 25),
                ForeColor = Color.LightGreen,
                Font = new Font("Consolas", 9),
                WordWrap = false
            };

            var btnCopy = new Button
            {
                Text = "📋 複製",
                Location = new Point(5, 520),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 120, 180),
                ForeColor = Color.White
            };
            btnCopy.Click += (s, e) => { 
                if (!string.IsNullOrEmpty(txtJson.Text))
                {
                    Clipboard.SetText(txtJson.Text);
                    lblStatus.Text = "✅ 已複製到剪貼簿";
                }
            };

            var btnImport = new Button
            {
                Text = "📥 匯入",
                Location = new Point(90, 520),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 150, 100),
                ForeColor = Color.White
            };
            btnImport.Click += (s, e) => ImportJson();

            var btnPaste = new Button
            {
                Text = "📋 貼上",
                Location = new Point(175, 520),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(100, 100, 120),
                ForeColor = Color.White
            };
            btnPaste.Click += (s, e) => {
                if (Clipboard.ContainsText())
                {
                    txtJson.Text = Clipboard.GetText();
                    lblStatus.Text = "已貼上剪貼簿內容";
                }
            };

            pnlRight.Controls.AddRange(new Control[] { lblJson, txtJson, btnCopy, btnImport, btnPaste });

            this.Controls.Add(pnlLeft);
            this.Controls.Add(pnlRight);
        }

        private Button CreateBtn(string text, int x, int y, Color bgColor)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(75, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei", 8)
            };
        }

        private void RefreshList()
        {
            lstCards.Items.Clear();
            foreach (var card in cards)
            {
                string status = card.Enabled ? "✓" : "✗";
                lstCards.Items.Add($"[{status}] #{card.Id} {card.GetDisplayText()}");
            }
            UpdateJson();
            UpdateStatus();
        }

        private void UpdateJson()
        {
            txtJson.Text = ScriptCardConverter.ToSimplifiedJson(cards);
        }

        private void UpdateStatus()
        {
            double totalDuration = cards.Where(c => c.Enabled).Sum(c => c.EstimateDuration());
            lblStatus.Text = $"卡片: {cards.Count} | 啟用: {cards.Count(c => c.Enabled)} | 預估時長: {totalDuration:F2}s";
        }

        private void AddCard(CardType type)
        {
            int newId = cards.Count > 0 ? cards.Max(c => c.Id) + 1 : 1;
            var card = new ScriptCard
            {
                Id = newId,
                Type = type,
                Key = Keys.A,
                Value = type switch
                {
                    CardType.Wait => 1.0,
                    CardType.Spam => 10,
                    CardType.Hold => 0.5,
                    CardType.PositionCorrect => 0,
                    _ => 0
                },
                IntervalMs = 50,
                Enabled = true
            };

            // 插入到選中位置之後
            int insertIndex = lstCards.SelectedIndex >= 0 ? lstCards.SelectedIndex + 1 : cards.Count;
            cards.Insert(insertIndex, card);
            
            RefreshList();
            lstCards.SelectedIndex = insertIndex;
            lblStatus.Text = $"✅ 新增 {type} 卡片";
            
            // 自動開啟編輯
            EditSelected();
        }

        private void EditSelected()
        {
            if (lstCards.SelectedIndex < 0 || lstCards.SelectedIndex >= cards.Count) return;
            
            var card = cards[lstCards.SelectedIndex];
            using var form = new CardEditDialog(card);
            if (form.ShowDialog() == DialogResult.OK)
            {
                RefreshList();
                lblStatus.Text = $"✅ 已更新卡片 #{card.Id}";
            }
        }

        private void DeleteSelected()
        {
            if (lstCards.SelectedIndex < 0) return;
            
            var result = MessageBox.Show("確定刪除此卡片？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                cards.RemoveAt(lstCards.SelectedIndex);
                RefreshList();
                lblStatus.Text = "✅ 已刪除";
            }
        }

        private void MoveCard(int direction)
        {
            int index = lstCards.SelectedIndex;
            if (index < 0) return;
            
            int newIndex = index + direction;
            if (newIndex < 0 || newIndex >= cards.Count) return;

            var temp = cards[index];
            cards[index] = cards[newIndex];
            cards[newIndex] = temp;

            RefreshList();
            lstCards.SelectedIndex = newIndex;
        }

        private void AddJitterToAll()
        {
            using var inputForm = new Form
            {
                Text = "加入隨機擾動",
                Size = new Size(280, 150),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(40, 40, 45)
            };

            var lbl = new Label
            {
                Text = "輸入擾動值 (±ms):",
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = Color.White
            };

            var num = new NumericUpDown
            {
                Location = new Point(20, 50),
                Size = new Size(100, 25),
                Minimum = 0,
                Maximum = 500,
                Value = 5
            };

            var btnOk = new Button
            {
                Text = "確定",
                Location = new Point(60, 85),
                Size = new Size(70, 28),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(140, 85),
                Size = new Size(70, 28),
                DialogResult = DialogResult.Cancel
            };

            inputForm.Controls.AddRange(new Control[] { lbl, num, btnOk, btnCancel });
            inputForm.AcceptButton = btnOk;
            inputForm.CancelButton = btnCancel;

            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                int jitter = (int)num.Value;
                int count = 0;
                foreach (var card in cards)
                {
                    if (card.Type == CardType.Wait || card.Type == CardType.Spam)
                    {
                        card.RandomJitterMs = jitter;
                        count++;
                    }
                }
                RefreshList();
                lblStatus.Text = $"✅ 已為 {count} 張卡片加入 ±{jitter}ms 擾動";
            }
        }

        private void ClearAll()
        {
            var result = MessageBox.Show($"確定清空全部 {cards.Count} 張卡片？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                cards.Clear();
                RefreshList();
                lblStatus.Text = "✅ 已清空";
            }
        }

        private void ImportJson()
        {
            var imported = ScriptCardConverter.FromSimplifiedJson(txtJson.Text);
            if (imported == null)
            {
                MessageBox.Show("JSON 格式錯誤！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show($"確定匯入 {imported.Count} 張卡片？\n這將取代目前所有卡片。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                cards = imported;
                RefreshList();
                lblStatus.Text = $"✅ 已匯入 {cards.Count} 張卡片";
            }
        }
    }

    /// <summary>
    /// 卡片編輯對話框
    /// </summary>
    public class CardEditDialog : Form
    {
        private ScriptCard card;
        private ComboBox cmbType;
        private ComboBox cmbKey;
        private NumericUpDown numValue;
        private NumericUpDown numInterval;
        private NumericUpDown numJitter;
        private CheckBox chkEnabled;

        public CardEditDialog(ScriptCard card)
        {
            this.card = card;
            InitForm();
        }

        private void InitForm()
        {
            this.Text = $"編輯卡片 #{card.Id}";
            this.Size = new Size(350, 320);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(40, 40, 45);
            this.ForeColor = Color.White;

            int y = 15;
            int labelX = 15;
            int inputX = 100;

            // 類型
            AddLabel("類型:", labelX, y);
            cmbType = new ComboBox
            {
                Location = new Point(inputX, y),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(50, 50, 55),
                ForeColor = Color.White
            };
            foreach (CardType ct in Enum.GetValues(typeof(CardType)))
                cmbType.Items.Add(ct);
            cmbType.SelectedItem = card.Type;
            cmbType.SelectedIndexChanged += (s, e) => UpdateFieldState();
            this.Controls.Add(cmbType);
            y += 35;

            // 按鍵
            AddLabel("按鍵:", labelX, y);
            cmbKey = new ComboBox
            {
                Location = new Point(inputX, y),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(50, 50, 55),
                ForeColor = Color.White
            };
            var keys = new Keys[] {
                Keys.A, Keys.B, Keys.C, Keys.D, Keys.E, Keys.F, Keys.G, Keys.H, Keys.I, Keys.J, Keys.K, Keys.L, Keys.M,
                Keys.N, Keys.O, Keys.P, Keys.Q, Keys.R, Keys.S, Keys.T, Keys.U, Keys.V, Keys.W, Keys.X, Keys.Y, Keys.Z,
                Keys.D0, Keys.D1, Keys.D2, Keys.D3, Keys.D4, Keys.D5, Keys.D6, Keys.D7, Keys.D8, Keys.D9,
                Keys.F1, Keys.F2, Keys.F3, Keys.F4, Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12,
                Keys.Space, Keys.Enter, Keys.Tab, Keys.Escape,
                Keys.Left, Keys.Right, Keys.Up, Keys.Down,
                Keys.LShiftKey, Keys.RShiftKey, Keys.LControlKey, Keys.RControlKey, Keys.LMenu, Keys.RMenu,
                Keys.Home, Keys.End, Keys.PageUp, Keys.PageDown, Keys.Insert, Keys.Delete
            };
            foreach (var k in keys) cmbKey.Items.Add(k);
            cmbKey.SelectedItem = card.Key;
            this.Controls.Add(cmbKey);
            y += 35;

            // 數值
            AddLabel("數值:", labelX, y);
            numValue = new NumericUpDown
            {
                Location = new Point(inputX, y),
                Size = new Size(120, 25),
                Minimum = 0,
                Maximum = 99999,
                DecimalPlaces = 3,
                Value = (decimal)Math.Max(0, Math.Min(99999, card.Value)),
                BackColor = Color.FromArgb(50, 50, 55),
                ForeColor = Color.White
            };
            this.Controls.Add(numValue);
            AddLabel("(秒/次)", 230, y);
            y += 35;

            // 間隔
            AddLabel("間隔ms:", labelX, y);
            numInterval = new NumericUpDown
            {
                Location = new Point(inputX, y),
                Size = new Size(120, 25),
                Minimum = 1,
                Maximum = 10000,
                Value = card.IntervalMs,
                BackColor = Color.FromArgb(50, 50, 55),
                ForeColor = Color.White
            };
            this.Controls.Add(numInterval);
            y += 35;

            // 擾動
            AddLabel("擾動±ms:", labelX, y);
            numJitter = new NumericUpDown
            {
                Location = new Point(inputX, y),
                Size = new Size(120, 25),
                Minimum = 0,
                Maximum = 500,
                Value = card.RandomJitterMs,
                BackColor = Color.FromArgb(50, 50, 55),
                ForeColor = Color.White
            };
            this.Controls.Add(numJitter);
            y += 35;

            // 啟用
            chkEnabled = new CheckBox
            {
                Text = "啟用此卡片",
                Location = new Point(inputX, y),
                Checked = card.Enabled,
                ForeColor = Color.White,
                AutoSize = true
            };
            this.Controls.Add(chkEnabled);
            y += 40;

            // 按鈕
            var btnOk = new Button
            {
                Text = "確定",
                Location = new Point(100, y),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 140, 80),
                ForeColor = Color.White,
                DialogResult = DialogResult.OK
            };
            btnOk.Click += (s, e) => SaveCard();
            this.Controls.Add(btnOk);

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(190, y),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(100, 100, 100),
                ForeColor = Color.White,
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            UpdateFieldState();
        }

        private void AddLabel(string text, int x, int y)
        {
            this.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                AutoSize = true,
                ForeColor = Color.White
            });
        }

        private void UpdateFieldState()
        {
            var type = (CardType)cmbType.SelectedItem;
            bool isPosCorrect = type == CardType.PositionCorrect;
            cmbKey.Enabled = type != CardType.Wait && !isPosCorrect;
            numValue.Enabled = !isPosCorrect;
            numInterval.Enabled = type == CardType.Spam;
            numJitter.Enabled = type == CardType.Wait || type == CardType.Spam;
        }

        private void SaveCard()
        {
            card.Type = (CardType)cmbType.SelectedItem;
            if (cmbKey.SelectedItem != null)
                card.Key = (Keys)cmbKey.SelectedItem;
            card.Value = (double)numValue.Value;
            card.IntervalMs = (int)numInterval.Value;
            card.RandomJitterMs = (int)numJitter.Value;
            card.Enabled = chkEnabled.Checked;
        }
    }
}

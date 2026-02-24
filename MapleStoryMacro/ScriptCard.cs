using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 腳本卡片類型
    /// </summary>
    public enum CardType
    {
        /// <summary>等待指定秒數</summary>
        Wait,
        /// <summary>按下並放開按鍵</summary>
        Click,
        /// <summary>快速連打按鍵</summary>
        Spam,
        /// <summary>長按按鍵</summary>
        Hold,
        /// <summary>按下按鍵（不放開）</summary>
        KeyDown,
        /// <summary>放開按鍵</summary>
        KeyUp,
        /// <summary>執行位置修正（回到目標座標）</summary>
        PositionCorrect
    }

    /// <summary>
    /// 腳本卡片 - 用於卡片式編輯器的數據結構
    /// </summary>
    public class ScriptCard
    {
        /// <summary>卡片唯一識別碼</summary>
        public int Id { get; set; }

        /// <summary>卡片類型</summary>
        public CardType Type { get; set; }

        /// <summary>按鍵（用於 Click, Spam, Hold, KeyDown, KeyUp）</summary>
        public Keys Key { get; set; }

        /// <summary>數值（秒數或次數，取決於類型）</summary>
        public double Value { get; set; }

        /// <summary>間隔毫秒（用於 Spam）</summary>
        public int IntervalMs { get; set; } = 50;

        /// <summary>隨機擾動範圍（用於 Spam 和 Wait）</summary>
        public int RandomJitterMs { get; set; } = 0;

        /// <summary>位置修正：目標 X（-1 表示使用全域設定）</summary>
        public int TargetX { get; set; } = -1;

        /// <summary>位置修正：目標 Y（-1 表示使用全域設定）</summary>
        public int TargetY { get; set; } = -1;

        /// <summary>備註說明</summary>
        public string? Note { get; set; }

        /// <summary>是否啟用此卡片</summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// 取得卡片的顯示文字
        /// </summary>
        public string GetDisplayText()
        {
            return Type switch
            {
                CardType.Wait => $"⏱️ Wait({Value:F2}s){(RandomJitterMs > 0 ? $" ±{RandomJitterMs}ms" : "")}",
                CardType.Click => $"🖱️ Click({Key})",
                CardType.Spam => $"🔥 Spam({Key}, {(int)Value}次, {IntervalMs}ms){(RandomJitterMs > 0 ? $" ±{RandomJitterMs}ms" : "")}",
                CardType.Hold => $"⏸️ Hold({Key}, {Value:F2}s)",
                CardType.KeyDown => $"⬇️ KeyDown({Key})",
                CardType.KeyUp => $"⬆️ KeyUp({Key})",
                CardType.PositionCorrect => TargetX >= 0 ? $"📍 PositionCorrect({TargetX},{TargetY})" : "📍 PositionCorrect（全域）)",
                _ => "❓ Unknown"
            };
        }

        /// <summary>
        /// 取得卡片的詳細說明
        /// </summary>
        public string GetDescription()
        {
            return Type switch
            {
                CardType.Wait => $"等待 {Value:F3} 秒{(RandomJitterMs > 0 ? $"（隨機 ±{RandomJitterMs}ms）" : "")}",
                CardType.Click => $"點擊 {Key} 鍵（按下後立即放開）",
                CardType.Spam => $"連打 {Key} 鍵 {(int)Value} 次，間隔 {IntervalMs}ms{(RandomJitterMs > 0 ? $"（隨機 ±{RandomJitterMs}ms）" : "")}",
                CardType.Hold => $"長按 {Key} 鍵 {Value:F3} 秒",
                CardType.KeyDown => $"按下 {Key} 鍵（不放開）",
                CardType.KeyUp => $"放開 {Key} 鍵",
                CardType.PositionCorrect => TargetX >= 0 ? $"位置修正到 ({TargetX}, {TargetY})" : "位置修正（使用全域目標座標）",
                _ => "未知動作"
            };
        }

        /// <summary>
        /// 估算此卡片執行所需的時間（秒）
        /// </summary>
        public double EstimateDuration()
        {
            return Type switch
            {
                CardType.Wait => Value,
                CardType.Click => 0.05, // 預估 50ms
                CardType.Spam => (Value * IntervalMs) / 1000.0,
                CardType.Hold => Value,
                CardType.KeyDown => 0.01,
                CardType.KeyUp => 0.01,
                CardType.PositionCorrect => 3.0, // 預估最多 3 秒
                _ => 0
            };
        }
    }

    /// <summary>
    /// 卡片腳本數據 - 用於 JSON 序列化/反序列化
    /// </summary>
    public class CardScriptData
    {
        public int Version { get; set; } = 2;
        public string? Name { get; set; }
        public DateTime ModifiedAt { get; set; } = DateTime.Now;
        public int LoopCount { get; set; } = 1;
        public List<ScriptCard> Cards { get; set; } = new List<ScriptCard>();

        /// <summary>
        /// 序列化為 JSON
        /// </summary>
        public string ToJson()
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Converters = { new JsonStringEnumConverter() }
            };
            return JsonSerializer.Serialize(this, options);
        }

        /// <summary>
        /// 從 JSON 反序列化
        /// </summary>
        public static CardScriptData? FromJson(string json)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                Converters = { new JsonStringEnumConverter() }
            };
            return JsonSerializer.Deserialize<CardScriptData>(json, options);
        }
    }

    /// <summary>
    /// 卡片與 MacroEvent 之間的轉換器
    /// </summary>
    public static class ScriptCardConverter
    {
        /// <summary>
        /// 將 MacroEvent 列表轉換為 ScriptCard 列表
        /// 會自動識別模式並合併成高級卡片
        /// </summary>
        public static List<ScriptCard> FromMacroEvents(List<Form1.MacroEvent> events)
        {
            var cards = new List<ScriptCard>();
            if (events == null || events.Count == 0) return cards;

            // 按時間排序
            var sortedEvents = events.OrderBy(e => e.Timestamp).ToList();

            int cardId = 1;
            int i = 0;

            while (i < sortedEvents.Count)
            {
                var current = sortedEvents[i];

                // 嘗試識別 Click 模式（down 緊接著 up）
                if (current.EventType == "down" && i + 1 < sortedEvents.Count)
                {
                    var next = sortedEvents[i + 1];
                    if (next.EventType == "up" && next.KeyCode == current.KeyCode)
                    {
                        double holdTime = next.Timestamp - current.Timestamp;

                        // 檢查是否為 Spam 模式（連續相同的 click）
                        int spamCount = 1;
                        double lastUpTime = next.Timestamp;
                        int j = i + 2;
                        var intervals = new List<double>();

                        while (j + 1 < sortedEvents.Count)
                        {
                            var downEvt = sortedEvents[j];
                            var upEvt = sortedEvents[j + 1];

                            if (downEvt.EventType == "down" && upEvt.EventType == "up" &&
                                downEvt.KeyCode == current.KeyCode && upEvt.KeyCode == current.KeyCode)
                            {
                                double interval = downEvt.Timestamp - lastUpTime;
                                if (interval < 0.3) // 300ms 內視為連打
                                {
                                    intervals.Add(interval * 1000); // 轉換為 ms
                                    spamCount++;
                                    lastUpTime = upEvt.Timestamp;
                                    j += 2;
                                }
                                else break;
                            }
                            else break;
                        }

                        if (spamCount >= 3) // 至少 3 次才算 Spam
                        {
                            int avgInterval = intervals.Count > 0 ? (int)intervals.Average() : 50;
                            cards.Add(new ScriptCard
                            {
                                Id = cardId++,
                                Type = CardType.Spam,
                                Key = current.KeyCode,
                                Value = spamCount,
                                IntervalMs = avgInterval
                            });
                            i = j;

                            // 添加後續等待時間
                            if (i < sortedEvents.Count)
                            {
                                double waitTime = sortedEvents[i].Timestamp - lastUpTime;
                                if (waitTime > 0.01)
                                {
                                    cards.Add(new ScriptCard
                                    {
                                        Id = cardId++,
                                        Type = CardType.Wait,
                                        Value = waitTime
                                    });
                                }
                            }
                            continue;
                        }

                        // 判斷是 Click 還是 Hold
                        if (holdTime < 0.2) // 200ms 以內算 Click
                        {
                            cards.Add(new ScriptCard
                            {
                                Id = cardId++,
                                Type = CardType.Click,
                                Key = current.KeyCode
                            });
                        }
                        else // 超過 200ms 算 Hold
                        {
                            cards.Add(new ScriptCard
                            {
                                Id = cardId++,
                                Type = CardType.Hold,
                                Key = current.KeyCode,
                                Value = holdTime
                            });
                        }

                        // 計算到下一個事件的等待時間
                        if (i + 2 < sortedEvents.Count)
                        {
                            double waitTime = sortedEvents[i + 2].Timestamp - next.Timestamp;
                            if (waitTime > 0.01) // 超過 10ms 才添加等待
                            {
                                cards.Add(new ScriptCard
                                {
                                    Id = cardId++,
                                    Type = CardType.Wait,
                                    Value = waitTime
                                });
                            }
                        }

                        i += 2;
                        continue;
                    }
                }

                // 單獨的 KeyDown 或 KeyUp
                cards.Add(new ScriptCard
                {
                    Id = cardId++,
                    Type = current.EventType == "down" ? CardType.KeyDown : CardType.KeyUp,
                    Key = current.KeyCode
                });

                // 計算等待時間
                if (i + 1 < sortedEvents.Count)
                {
                    double waitTime = sortedEvents[i + 1].Timestamp - current.Timestamp;
                    if (waitTime > 0.01)
                    {
                        cards.Add(new ScriptCard
                        {
                            Id = cardId++,
                            Type = CardType.Wait,
                            Value = waitTime
                        });
                    }
                }

                i++;
            }

            // 合併連續的等待卡片
            cards = MergeConsecutiveWaits(cards);

            return cards;
        }

        /// <summary>
        /// 合併連續的等待卡片
        /// </summary>
        private static List<ScriptCard> MergeConsecutiveWaits(List<ScriptCard> cards)
        {
            var result = new List<ScriptCard>();
            int id = 1;

            for (int i = 0; i < cards.Count; i++)
            {
                if (cards[i].Type == CardType.Wait)
                {
                    double totalWait = cards[i].Value;
                    while (i + 1 < cards.Count && cards[i + 1].Type == CardType.Wait)
                    {
                        totalWait += cards[i + 1].Value;
                        i++;
                    }
                    result.Add(new ScriptCard
                    {
                        Id = id++,
                        Type = CardType.Wait,
                        Value = totalWait
                    });
                }
                else
                {
                    var card = cards[i];
                    card.Id = id++;
                    result.Add(card);
                }
            }

            return result;
        }

        /// <summary>
        /// 將 ScriptCard 列表轉換為 MacroEvent 列表
        /// </summary>
        public static List<Form1.MacroEvent> ToMacroEvents(List<ScriptCard> cards)
        {
            var events = new List<Form1.MacroEvent>();
            double currentTime = 0;
            var random = new Random();

            foreach (var card in cards.Where(c => c.Enabled))
            {
                switch (card.Type)
                {
                    case CardType.Wait:
                        double waitTime = card.Value;
                        if (card.RandomJitterMs > 0)
                        {
                            waitTime += (random.NextDouble() * 2 - 1) * card.RandomJitterMs / 1000.0;
                            waitTime = Math.Max(0, waitTime);
                        }
                        currentTime += waitTime;
                        break;

                    case CardType.Click:
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = card.Key,
                            EventType = "down",
                            Timestamp = currentTime
                        });
                        currentTime += 0.03; // 30ms 延遲
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = card.Key,
                            EventType = "up",
                            Timestamp = currentTime
                        });
                        currentTime += 0.02; // 20ms 緩衝
                        break;

                    case CardType.Spam:
                        int count = (int)card.Value;
                        for (int i = 0; i < count; i++)
                        {
                            events.Add(new Form1.MacroEvent
                            {
                                KeyCode = card.Key,
                                EventType = "down",
                                Timestamp = currentTime
                            });
                            currentTime += 0.02; // 20ms 按下
                            events.Add(new Form1.MacroEvent
                            {
                                KeyCode = card.Key,
                                EventType = "up",
                                Timestamp = currentTime
                            });

                            if (i < count - 1) // 不是最後一次
                            {
                                double interval = card.IntervalMs / 1000.0;
                                if (card.RandomJitterMs > 0)
                                {
                                    interval += (random.NextDouble() * 2 - 1) * card.RandomJitterMs / 1000.0;
                                    interval = Math.Max(0.01, interval);
                                }
                                currentTime += interval;
                            }
                        }
                        break;

                    case CardType.Hold:
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = card.Key,
                            EventType = "down",
                            Timestamp = currentTime
                        });
                        currentTime += card.Value;
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = card.Key,
                            EventType = "up",
                            Timestamp = currentTime
                        });
                        break;

                    case CardType.KeyDown:
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = card.Key,
                            EventType = "down",
                            Timestamp = currentTime
                        });
                        currentTime += 0.01;
                        break;

                    case CardType.KeyUp:
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = card.Key,
                            EventType = "up",
                            Timestamp = currentTime
                        });
                        currentTime += 0.01;
                        break;

                    case CardType.PositionCorrect:
                        events.Add(new Form1.MacroEvent
                        {
                            KeyCode = Keys.None,
                            EventType = "position_correct",
                            Timestamp = currentTime,
                            CorrectTargetX = card.TargetX,
                            CorrectTargetY = card.TargetY
                        });
                        currentTime += 0.01; // 佔位，實際執行時間由修正器決定
                        break;
                }
            }

            return events;
        }

        /// <summary>
        /// 將卡片列表匯出為簡化 JSON（給 AI 用）
        /// </summary>
        public static string ToSimplifiedJson(List<ScriptCard> cards)
        {
            var simplified = cards.Select(c => new
            {
                id = c.Id,
                type = c.Type.ToString(),
                key = c.Type != CardType.Wait ? c.Key.ToString() : null,
                value = c.Value,
                intervalMs = c.Type == CardType.Spam ? c.IntervalMs : (int?)null,
                jitterMs = c.RandomJitterMs > 0 ? c.RandomJitterMs : (int?)null,
                note = c.Note,
                enabled = c.Enabled
            });

            return JsonSerializer.Serialize(simplified, new JsonSerializerOptions { WriteIndented = true });
        }

        /// <summary>
        /// 從簡化 JSON 匯入卡片列表
        /// </summary>
        public static List<ScriptCard>? FromSimplifiedJson(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                var cards = new List<ScriptCard>();

                foreach (var element in doc.RootElement.EnumerateArray())
                {
                    var card = new ScriptCard();

                    if (element.TryGetProperty("id", out var idProp))
                        card.Id = idProp.GetInt32();

                    if (element.TryGetProperty("type", out var typeProp))
                    {
                        if (Enum.TryParse<CardType>(typeProp.GetString(), true, out var type))
                            card.Type = type;
                    }

                    if (element.TryGetProperty("key", out var keyProp) && keyProp.ValueKind != JsonValueKind.Null)
                    {
                        if (Enum.TryParse<Keys>(keyProp.GetString(), true, out var key))
                            card.Key = key;
                    }

                    if (element.TryGetProperty("value", out var valueProp))
                        card.Value = valueProp.GetDouble();

                    if (element.TryGetProperty("intervalMs", out var intervalProp) && intervalProp.ValueKind != JsonValueKind.Null)
                        card.IntervalMs = intervalProp.GetInt32();

                    if (element.TryGetProperty("jitterMs", out var jitterProp) && jitterProp.ValueKind != JsonValueKind.Null)
                        card.RandomJitterMs = jitterProp.GetInt32();

                    if (element.TryGetProperty("note", out var noteProp) && noteProp.ValueKind != JsonValueKind.Null)
                        card.Note = noteProp.GetString();

                    if (element.TryGetProperty("enabled", out var enabledProp))
                        card.Enabled = enabledProp.GetBoolean();
                    else
                        card.Enabled = true;

                    cards.Add(card);
                }

                return cards;
            }
            catch
            {
                return null;
            }
        }
    }
}

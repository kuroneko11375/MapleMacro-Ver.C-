using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace MapleStoryMacro
{
    /// <summary>
    /// 記憶體掃描器擴展 - 新增 AOB 模式和指針鏈功能
    /// 這些方法應該添加到現有的 MemoryScanner 類別中
    /// </summary>
    public partial class MemoryScanner
    {
        #region AOB (Array of Bytes) 模式掃描

        /// <summary>
        /// AOB 模式掃描 - 搜尋特徵碼（可用通配符 ?? ）
        /// 例如：FindAOB("8B 0D ?? ?? ?? ?? 85 C9 74 ?? 8B 01")
        /// 用於尋找不會因遊戲更新而改變的程式碼特徵
        /// </summary>
        /// <param name="pattern">AOB 模式字串，用空格分隔，?? 表示通配符</param>
        /// <param name="maxResults">最大結果數（0 表示不限制）</param>
        /// <returns>找到的位址列表</returns>
        public List<IntPtr> FindAOB(string pattern, int maxResults = 100)
        {
            Logger.Info($"開始 AOB 掃描: {pattern}");
            using (Logger.PerfScope("AOB 掃描"))
            {
                List<IntPtr> results = new List<IntPtr>();

                if (processHandle == IntPtr.Zero)
                {
                    Logger.Warning("未附加到進程，AOB 掃描失敗");
                    return results;
                }

                // 解析 AOB 模式
                if (!ParseAOBPattern(pattern, out byte[] patternBytes, out bool[] mask))
                {
                    Logger.Error($"無效的 AOB 模式: {pattern}");
                    return results;
                }

                Logger.Debug($"AOB 模式長度: {patternBytes.Length} 字節");

                IntPtr address = IntPtr.Zero;
                long maxAddress = 0x7FFFFFFF; // 32 位元進程上限
                int scannedRegions = 0;

                while (address.ToInt64() < maxAddress)
                {
                    // 查詢記憶體區域
                    int result = VirtualQueryEx(processHandle, address, out MEMORY_BASIC_INFORMATION mbi,
                        Marshal.SizeOf<MEMORY_BASIC_INFORMATION>());
                    if (result == 0)
                        break;

                    long regionSize = mbi.RegionSize.ToInt64();
                    if (regionSize <= 0)
                        break;

                    // 掃描可讀取的記憶體區域
                    if (mbi.State == MEM_COMMIT && IsReadable(mbi.Protect))
                    {
                        scannedRegions++;
                        ScanRegionForAOB(mbi.BaseAddress, (int)Math.Min(regionSize, int.MaxValue),
                            patternBytes, mask, results, maxResults);

                        if (maxResults > 0 && results.Count >= maxResults)
                        {
                            Logger.Info($"已達到最大結果數 {maxResults}，停止掃描");
                            break;
                        }
                    }

                    // 移動到下一個區域
                    long nextAddr = mbi.BaseAddress.ToInt64() + regionSize;
                    if (nextAddr <= address.ToInt64())
                        break;
                    address = (IntPtr)nextAddr;
                }

                Logger.Info($"AOB 掃描完成: 掃描了 {scannedRegions} 個區域，找到 {results.Count} 個匹配");
                return results;
            }
        }

        /// <summary>
        /// 解析 AOB 模式字串
        /// </summary>
        private bool ParseAOBPattern(string pattern, out byte[] bytes, out bool[] mask)
        {
            bytes = Array.Empty<byte>();
            mask = Array.Empty<bool>();

            try
            {
                // 移除多餘空白
                pattern = pattern.Trim();
                string[] parts = pattern.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                bytes = new byte[parts.Length];
                mask = new bool[parts.Length];

                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i] == "??" || parts[i] == "?")
                    {
                        bytes[i] = 0;
                        mask[i] = false; // 通配符
                    }
                    else
                    {
                        bytes[i] = Convert.ToByte(parts[i], 16);
                        mask[i] = true;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"解析 AOB 模式失敗: {pattern}", ex);
                return false;
            }
        }

        /// <summary>
        /// 掃描單個記憶體區域尋找 AOB 模式
        /// </summary>
        private void ScanRegionForAOB(IntPtr baseAddress, int regionSize, byte[] pattern, bool[] mask,
            List<IntPtr> results, int maxResults)
        {
            try
            {
                byte[] buffer = new byte[regionSize];
                if (!ReadProcessMemory(processHandle, baseAddress, buffer, regionSize, out int bytesRead))
                    return;

                // 在緩衝區中搜尋模式
                for (int i = 0; i <= bytesRead - pattern.Length; i++)
                {
                    if (ComparePattern(buffer, i, pattern, mask))
                    {
                        IntPtr foundAddress = (IntPtr)(baseAddress.ToInt64() + i);
                        results.Add(foundAddress);

                        Logger.Debug($"找到 AOB 匹配: 0x{foundAddress.ToInt64():X}");

                        if (maxResults > 0 && results.Count >= maxResults)
                            return;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"掃描區域 0x{baseAddress.ToInt64():X} 時發生錯誤", ex);
            }
        }

        /// <summary>
        /// 比對 AOB 模式
        /// </summary>
        private bool ComparePattern(byte[] buffer, int offset, byte[] pattern, bool[] mask)
        {
            for (int i = 0; i < pattern.Length; i++)
            {
                if (mask[i] && buffer[offset + i] != pattern[i])
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 檢查記憶體區域是否可讀取
        /// </summary>
        private bool IsReadable(uint protect)
        {
            const uint PAGE_NOACCESS = 0x01;
            const uint PAGE_GUARD = 0x100;

            if ((protect & PAGE_NOACCESS) != 0 || (protect & PAGE_GUARD) != 0)
                return false;

            // 可讀取的保護類型
            const uint PAGE_READONLY = 0x02;
            const uint PAGE_READWRITE = 0x04;
            const uint PAGE_WRITECOPY = 0x08;
            const uint PAGE_EXECUTE_READ = 0x20;
            const uint PAGE_EXECUTE_READWRITE = 0x40;
            const uint PAGE_EXECUTE_WRITECOPY = 0x80;

            return (protect & (PAGE_READONLY | PAGE_READWRITE | PAGE_WRITECOPY |
                               PAGE_EXECUTE_READ | PAGE_EXECUTE_READWRITE | PAGE_EXECUTE_WRITECOPY)) != 0;
        }

        #endregion

        #region 指針鏈功能

        /// <summary>
        /// 解析指針鏈並讀取最終值
        /// </summary>
        public bool ReadPointerChain(PointerPath path, out int value)
        {
            value = 0;

            try
            {
                // 1. 取得模組基址
                IntPtr moduleBase = GetModuleBaseAddress(path.ModuleName);
                if (moduleBase == IntPtr.Zero)
                {
                    Logger.Warning($"找不到模組: {path.ModuleName}");
                    return false;
                }

                Logger.Debug($"模組 {path.ModuleName} 基址: 0x{moduleBase.ToInt64():X}");

                // 2. 計算第一層位址
                IntPtr currentAddress = (IntPtr)(moduleBase.ToInt64() + path.BaseOffset);
                Logger.Debug($"起始位址: [0x{moduleBase.ToInt64():X}+0x{path.BaseOffset:X}] = 0x{currentAddress.ToInt64():X}");

                // 3. 逐層解析指針
                for (int i = 0; i < path.Offsets.Count; i++)
                {
                    // 讀取當前位址的值（作為下一層的位址）
                    if (!ReadInt32(currentAddress, out int pointerValue))
                    {
                        Logger.Warning($"無法讀取指針 (層 {i}): 0x{currentAddress.ToInt64():X}");
                        return false;
                    }

                    currentAddress = (IntPtr)(pointerValue + path.Offsets[i]);
                    Logger.Debug($"  層 {i + 1}: [0x{pointerValue:X}+0x{path.Offsets[i]:X}] = 0x{currentAddress.ToInt64():X}");
                }

                // 4. 讀取最終值
                bool success = ReadInt32(currentAddress, out value);
                if (success)
                {
                    Logger.Debug($"最終值: {value} (0x{value:X})");
                }
                else
                {
                    Logger.Warning($"無法讀取最終值: 0x{currentAddress.ToInt64():X}");
                }

                return success;
            }
            catch (Exception ex)
            {
                Logger.Error($"讀取指針鏈失敗: {path}", ex);
                return false;
            }
        }

        /// <summary>
        /// 取得模組基址
        /// </summary>
        public IntPtr GetModuleBaseAddress(string moduleName)
        {
            try
            {
                Process process = Process.GetProcessById((int)processId);
                foreach (ProcessModule module in process.Modules)
                {
                    if (module.ModuleName.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return module.BaseAddress;
                    }
                }

                Logger.Warning($"找不到模組: {moduleName}");
                return IntPtr.Zero;
            }
            catch (Exception ex)
            {
                Logger.Error($"取得模組基址失敗: {moduleName}", ex);
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// 取得所有模組列表
        /// </summary>
        public List<ModuleInfo> GetModuleList()
        {
            var modules = new List<ModuleInfo>();

            try
            {
                Process process = Process.GetProcessById((int)processId);
                foreach (ProcessModule module in process.Modules)
                {
                    modules.Add(new ModuleInfo
                    {
                        Name = module.ModuleName,
                        BaseAddress = module.BaseAddress,
                        Size = module.ModuleMemorySize
                    });
                }

                Logger.Debug($"取得 {modules.Count} 個模組");
            }
            catch (Exception ex)
            {
                Logger.Error("取得模組列表失敗", ex);
            }

            return modules;
        }

        public class ModuleInfo
        {
            public string Name { get; set; } = string.Empty;
            public IntPtr BaseAddress { get; set; }
            public int Size { get; set; }
        }

        /// <summary>
        /// 自動搜尋指針鏈（深度優先搜索）
        /// 警告：這是一個非常耗時的操作！
        /// </summary>
        public List<PointerPath> FindPointerPaths(IntPtr targetAddress, int maxDepth = 5, int maxResults = 50)
        {
            Logger.Info($"開始搜尋指向 0x{targetAddress.ToInt64():X} 的指針鏈 (最大深度: {maxDepth})");
            using (Logger.PerfScope("指針鏈搜尋"))
            {
                List<PointerPath> results = new List<PointerPath>();
                HashSet<IntPtr> visited = new HashSet<IntPtr>();

                // 取得所有模組
                var modules = GetModuleList();
                Logger.Info($"將在 {modules.Count} 個模組中搜尋");

                foreach (var module in modules)
                {
                    Logger.Debug($"搜尋模組: {module.Name}");

                    // 在模組範圍內搜尋
                    SearchPointerRecursive(module.BaseAddress, module.Name, module.Size,
                        targetAddress, new List<int>(), 0, maxDepth, visited, results, maxResults);

                    if (results.Count >= maxResults)
                    {
                        Logger.Info($"已達到最大結果數 {maxResults}，停止搜尋");
                        break;
                    }
                }

                Logger.Info($"指針鏈搜尋完成: 找到 {results.Count} 條路徑");
                return results;
            }
        }

        /// <summary>
        /// 遞迴搜尋指針鏈
        /// </summary>
        private void SearchPointerRecursive(IntPtr currentAddress, string moduleName, int moduleSize,
            IntPtr targetAddress, List<int> offsets, int depth, int maxDepth,
            HashSet<IntPtr> visited, List<PointerPath> results, int maxResults)
        {
            if (depth > maxDepth || results.Count >= maxResults || visited.Contains(currentAddress))
                return;

            visited.Add(currentAddress);

            // 讀取當前位址的值
            if (!ReadInt32(currentAddress, out int value))
                return;

            IntPtr valueAsPointer = (IntPtr)value;

            // 在目標位址附近搜尋（±0x1000 範圍）
            const int searchRange = 0x1000;
            for (int offset = -searchRange; offset <= searchRange; offset += 4)
            {
                IntPtr testAddress = (IntPtr)(valueAsPointer.ToInt64() + offset);

                if (testAddress == targetAddress)
                {
                    // 找到目標！
                    var path = new PointerPath
                    {
                        ModuleName = moduleName,
                        BaseOffset = currentAddress.ToInt64(),
                        Offsets = new List<int>(offsets) { offset }
                    };
                    results.Add(path);

                    Logger.Info($"找到指針鏈: {path}");
                    return;
                }
            }

            // 繼續深度搜尋（如果值看起來像有效的指針）
            if (IsValidPointer(valueAsPointer))
            {
                var newOffsets = new List<int>(offsets) { 0 };
                SearchPointerRecursive(valueAsPointer, moduleName, moduleSize,
                    targetAddress, newOffsets, depth + 1, maxDepth, visited, results, maxResults);
            }
        }

        /// <summary>
        /// 檢查是否為有效的指針值
        /// </summary>
        private bool IsValidPointer(IntPtr pointer)
        {
            long value = pointer.ToInt64();

            // 32 位元進程的有效範圍
            if (value < 0x10000 || value > 0x7FFFFFFF)
                return false;

            // 檢查是否對齊（通常是 4 字節對齊）
            if (value % 4 != 0)
                return false;

            return true;
        }

        #endregion

        #region 座標配置持久化

        /// <summary>
        /// 驗證座標配置是否仍然有效
        /// </summary>
        public bool VerifyCoordinateConfig(CoordinateConfig config)
        {
            Logger.Info("驗證座標配置...");

            try
            {
                if (config.XCoordinatePath == null || config.YCoordinatePath == null)
                {
                    Logger.Warning("座標路徑為空");
                    return false;
                }

                // 嘗試讀取 X 和 Y 座標
                bool xValid = ReadPointerChain(config.XCoordinatePath, out int x);
                bool yValid = ReadPointerChain(config.YCoordinatePath, out int y);

                if (xValid && yValid)
                {
                    // 檢查座標是否在合理範圍內
                    bool inRange = (x >= -50000 && x <= 50000) && (y >= -50000 && y <= 50000);

                    if (inRange)
                    {
                        Logger.Info($"座標配置有效: X={x}, Y={y}");
                        config.LastVerified = DateTime.Now;
                        return true;
                    }
                    else
                    {
                        Logger.Warning($"座標超出合理範圍: X={x}, Y={y}");
                    }
                }
                else
                {
                    Logger.Warning("無法讀取座標值");
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.Error("驗證座標配置時發生錯誤", ex);
                return false;
            }
        }

        /// <summary>
        /// 自動尋找並建立座標配置
        /// </summary>
        public CoordinateConfig? AutoFindCoordinates(int currentX, int currentY)
        {
            Logger.Info($"自動尋找座標位址: X={currentX}, Y={currentY}");

            using (Logger.PerfScope("自動尋找座標"))
            {
                // 1. 首次掃描尋找 X 座標
                Logger.Info("搜尋 X 座標...");
                int xCount = FirstScan(currentX, tolerance: 5);
                Logger.Info($"找到 {xCount} 個可能的 X 座標位址");

                if (xCount == 0 || xCount > 1000)
                {
                    Logger.Warning($"X 座標候選數量異常: {xCount}");
                    return null;
                }

                var xCandidates = new List<IntPtr>(Results.Select(r => r.Address));

                // 2. 首次掃描尋找 Y 座標
                Logger.Info("搜尋 Y 座標...");
                int yCount = FirstScan(currentY, tolerance: 5);
                Logger.Info($"找到 {yCount} 個可能的 Y 座標位址");

                if (yCount == 0 || yCount > 1000)
                {
                    Logger.Warning($"Y 座標候選數量異常: {yCount}");
                    return null;
                }

                var yCandidates = new List<IntPtr>(Results.Select(r => r.Address));

                // 3. 嘗試建立指針鏈（簡化版本：使用直接位址）
                // 實際使用時應該等待座標變化，進行多次掃描縮小範圍，然後搜尋指針鏈
                var config = new CoordinateConfig
                {
                    XCoordinatePath = new PointerPath
                    {
                        ModuleName = "MapleStory.exe",
                        BaseOffset = xCandidates[0].ToInt64(),
                        Offsets = new List<int>()
                    },
                    YCoordinatePath = new PointerPath
                    {
                        ModuleName = "MapleStory.exe",
                        BaseOffset = yCandidates[0].ToInt64(),
                        Offsets = new List<int>()
                    },
                    GameVersion = GetGameVersion(),
                    LastVerified = DateTime.Now
                };

                Logger.Info($"建立座標配置: X={config.XCoordinatePath}, Y={config.YCoordinatePath}");
                return config;
            }
        }

        /// <summary>
        /// 取得遊戲版本（用於檢測更新）
        /// </summary>
        private string GetGameVersion()
        {
            try
            {
                Process process = Process.GetProcessById((int)processId);
                if (process.MainModule != null)
                {
                    var versionInfo = FileVersionInfo.GetVersionInfo(process.MainModule.FileName);
                    return versionInfo.FileVersion ?? "Unknown";
                }
            }
            catch (Exception ex)
            {
                Logger.Warning("無法取得遊戲版本", ex);
            }

            return "Unknown";
        }

        #endregion
    }
}
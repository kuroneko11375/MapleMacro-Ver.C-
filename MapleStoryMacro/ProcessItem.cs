using System;

namespace MapleStoryMacro
{
    /// <summary>
    /// 視窗選擇器中的進程項目
    /// </summary>
    public class ProcessItem
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; } = string.Empty;
        public bool Is32Bit { get; set; }
    }
}

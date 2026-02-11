using System;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// ����]�w���O - �Ω��x�s��}��������]�w (.json)
    /// ���]�t�}���S�w��ơ]�ƥ�B�۩w�q����B�`����ơ^
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// �������
        /// </summary>
        public Keys PlayHotkey { get; set; } = Keys.F9;

        /// <summary>
        /// �������
        /// </summary>
        public Keys StopHotkey { get; set; } = Keys.F10;

        /// <summary>
        /// ����O�_�ҥ�
        /// </summary>
        public bool HotkeyEnabled { get; set; } = true;

        /// <summary>
        /// �ؼе������D
        /// </summary>
        public string WindowTitle { get; set; } = "MapleStory";

        /// <summary>
        /// ��V��o�e�Ҧ�
        /// 0=SendToChild, 1=ThreadAttachWithBlocker, 2=SendInputWithBlock
        /// </summary>
        public int ArrowKeyMode { get; set; } = 0;

        /// <summary>
        /// �̫���J���}����|�]��K�U���۰ʸ��J�^
        /// </summary>
        public string? LastScriptPath { get; set; }
    }

    /// <summary>
    /// �۩w�q����Ѧ��� (�i�ǦC��)
    /// </summary>
    public class CustomKeySlotData
    {
        public int SlotNumber { get; set; }
        public int KeyCode { get; set; } = (int)Keys.None;
        public double IntervalSeconds { get; set; } = 30.0;
        public bool Enabled { get; set; } = false;
        public double StartAtSecond { get; set; } = 0;
        public double PreDelaySeconds { get; set; } = 0;
        public double PauseScriptSeconds { get; set; } = 0;
        public bool PauseScriptEnabled { get; set; } = false;
    }
}

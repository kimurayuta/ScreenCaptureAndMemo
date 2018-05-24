using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Runtime.InteropServices;

namespace ScreenCaptureAndMemo
{
    public partial class ImageForm : Form
    {
        public ImageForm()
        {
            InitializeComponent();
            // http://stackoverflow.com/questions/15413172/capture-a-keyboard-keypress-in-the-background
            // Modifier keys codes: Alt = 1, Ctrl = 2, Shift = 4, Win = 8
            // Compute the addition of each combination of the keys you want to be pressed
            // ALT+CTRL = 1 + 2 = 3 , CTRL+SHIFT = 2 + 4 = 6...
            RegisterHotKey(this.Handle, MYACTION_HOTKEY_ID, 0, (int)Keys.F2);
        }

        private void ImageForm3_Load(object sender, EventArgs e)
        {

        }

        public void WriteMemo()
        {
            if (pictureBox1.ImageLocation != null)
            {
                richTextBox1.SaveFile(pictureBox1.ImageLocation + ".txt", RichTextBoxStreamType.UnicodePlainText);
            }
        }

        public void doCapture()
        {
            WriteMemo();
            Bitmap bmp = Program.getActiveCapture();
            string rootPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\capt"
                    + DateTime.Now.ToString("yyyyMMddHHmmss");
            string imagePath = rootPath + ".png";
            bmp.Save(imagePath,
                System.Drawing.Imaging.ImageFormat.Png);
            bmp.Dispose();

            pictureBox1.ImageLocation = imagePath;
            if (Clipboard.GetText() != null)
            {
                richTextBox1.Text = Clipboard.GetText();
            }
            else
            {
                richTextBox1.Text = "";
            }

            this.Activate();
            this.richTextBox1.Focus();
        }

        // http://stackoverflow.com/questions/15413172/capture-a-keyboard-keypress-in-the-background
        // DLL libraries used to manage hotkeys
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public const int MYACTION_HOTKEY_ID = 1;

        public void Unregister()
        {
            WriteMemo();
            UnregisterHotKey(this.Handle, MYACTION_HOTKEY_ID);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == MYACTION_HOTKEY_ID)
            {
                doCapture();
            }
            base.WndProc(ref m);
        }
    }
}

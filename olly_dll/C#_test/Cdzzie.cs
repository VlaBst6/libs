using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;

namespace dz
{
    public static class UI
    {
        public static string pad(string x)
        {
            while (x.Length < 14) x += " ";
            return x;
        }

        public static string GetListViewContents(ListView lv)
        {

            string table = "";

            foreach (ColumnHeader ch in lv.Columns)
            {
                table += pad(ch.Text);
            }
            table += "\r\n" + new String('-', 75) + "\r\n";

            foreach (ListViewItem li in lv.Items)
            {
                try
                {
                    table += pad(li.Text);
                    for (int i = 1; i < li.SubItems.Count; i++)
                    {
                        table += pad(li.SubItems[i].Text);
                    }
                    table += "\r\n";
                }
                catch (Exception e) { };
            }

            return table;

        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

    }
    public static class HexDumper
    {

        private const int LineLen = 16;
        private static int bCount = 0;
        private static byte[] bytes = new byte[LineLen];
        private static StringBuilder buf;

        public static string HexDump(string str)
        {
            buf = new StringBuilder();
            char[] ch = str.ToCharArray();
            for (int i = 0; i < ch.Length; i++) AddByte((byte)ch[i], (i == ch.Length - 1));
            return buf.ToString();
        }

        public static string HexDump(byte[] b)
        {
            if (b == null) return "";
            buf = new StringBuilder();
            for (int i = 0; i < b.Length; i++) AddByte(b[i], (i == b.Length - 1));
            return buf.ToString();
        }

        public static string HexDump(byte[] b, bool showOffset)
        {
            if (b == null) return "";
            buf = new StringBuilder();
            for (int i = 0; i < b.Length; i++)
            {
                if (showOffset && (i == 0 || i % 16 == 0)) buf.Append(i.ToString("X05") + "   ");
                AddByte(b[i], (i == b.Length - 1));
            }
            return buf.ToString();
        }


        private static void AddByte(byte b, bool final)
        {

            bytes[bCount++] = b;
            if (!final) if (bCount != LineLen) return;
            if (bCount <= 0) return;

            //main dump section
            for (int i = 0; i < LineLen; i++)
            {
                buf.Append(i >= bCount ? "   " : bytes[i].ToString("X2") + " ");
            }

            buf.Append("  ");

            //char display pad
            for (int i = 0; i < LineLen; i++)
            {
                byte ch = bytes[i] >= 32 && bytes[i] <= 126 ? bytes[i] : (byte)0x2e; //dot
                buf.Append(i >= bCount ? " " : (char)ch + "");
            }

            buf.Append("\n");
            bCount = 0;
        }
    }

    public static class Hasher
    {
        public static string MD5String(string input)
        {
            MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
            byte[] hash = md5.ComputeHash(inputBytes);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                sb.Append(hash[i].ToString("X2"));
            }
            return sb.ToString();
        }

        public static string MD5Bytes(byte[] inputBytes)
        {
            MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] hash = md5.ComputeHash(inputBytes);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                sb.Append(hash[i].ToString("X2"));
            }
            return sb.ToString();
        }

        public static string MD5File(string path)
        {
            if (!File.Exists(path)) return "";
            byte[] b = File.ReadAllBytes(path);
            return MD5Bytes(b);
        }
    }

    public static class Conv
    {
        public static string B64Encode(byte[] buf){
            try
            {
                return System.Convert.ToBase64String(buf,0,buf.Length);
            }
            catch (System.ArgumentNullException)
            {
                return "";
            }
        }

        public static byte[] B64Decode(string buf)
        {
            try
            {
                return System.Convert.FromBase64String(buf);
            }
            catch (System.ArgumentNullException)
            {
                return new byte[0];
            }
        }

        public static uint GetHexNum(string buf){
            return UInt32.Parse(buf, System.Globalization.NumberStyles.HexNumber);
        }

        public static uint GetInt(string buf)
        {
            return UInt32.Parse(buf, System.Globalization.NumberStyles.Integer);
        }

        public static string BytesToString(byte[] buf)
        {
            if (buf == null) return "";
            System.Text.UTF8Encoding enc = new System.Text.UTF8Encoding();
            return enc.GetString(buf);
        }

        public static byte[] UnicodeToASCII(string buf)
        {
            return System.Text.Encoding.ASCII.GetBytes(buf);
        }

        public static byte[] StringToBytes(string buf)
        {
            return System.Text.Encoding.Unicode.GetBytes(buf);
        }


    }


}

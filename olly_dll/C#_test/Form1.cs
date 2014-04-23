using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CS_test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            byte[] b = new byte[0];
            Int32 offset = 0x401000;
            
            if( CAssembler.Asm("jmp 0x401234", offset, ref b) > 0 ){
                rtf.Text = "Asm( jmp 0x401234, 0x401000)  -> \n"
                            + "\tLength: " + b.Length + "\n"
                            + "\tBytes:  " + dz.HexDumper.HexDump(b)
                            + "\n\n";
            }

            CInstruction ci = CDisassembler.Dsm(b,0,offset);
            if(ci.instLen > 0){
                rtf.Text += "Disasm(buf) -> \n" 
                           + "\tLen:  " + ci.instLen + "\n"
                           + "\tDump: " + ci.dump + "\n" 
                           + "\tCmd:  " + ci.command; 
            }else{
                 rtf.Text += "\n Disasm(buf) -> Error";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> cmds = new List<string>();
            List<byte> buf = new List<byte>();

            cmds.Add("MOV EAX, [EAX]");
            cmds.Add("CALL 4012bb");
            cmds.Add("MOV AL, CL");
            cmds.Add("JNZ 401567");

            int sz = CAssembler.AsmBlock(cmds, 0x401000, ref buf);

            if (sz == 0)
            {
                MessageBox.Show("Assembly Error: " + CAssembler.ErrorMessage);
                return;
            }

            rtf.Text = "Assembled: \n\t" + string.Join("\n\t", cmds.ToArray()) + "\n\n" +
                       dz.HexDumper.HexDump(buf.ToArray()) + "\n\n";


            List<CInstruction> disasm = CDisassembler.DsmBlock(buf.ToArray(), 0x401000);

            rtf.Text += "Block Disasm:\n\n";

            foreach (CInstruction ins in disasm)
            {
                rtf.Text += "\t" + ins.offset.ToString("X") + "\t " + pad(ins.dump) + ins.command + "\n";
            }
        }

        private string pad(string s)
        {
            while(s.Length < 20) s +=" ";
            return s;
        }

    }
}

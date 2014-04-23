using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;


/* VB6 def
 * 
    Private Type t_asmmodel
        code(15) As Byte
        mask(15) As Byte
        Length As Long
        jmpsize As Long
        jmpoffset As Long
        jmppos As Long
    End Type

    Private Declare Function Assemble Lib "olly.dll" ( _
            ByVal CMD As String, ByVal ip As Long, model As t_asmmodel, _
            ByVal attempt As Long, ByVal constsize As Long, ByVal errtext As String) As Long

    Private Declare Sub VB_SetOptions Lib "olly.dll" ( _
            Optional ByVal isideal As Long = 0, Optional ByVal isLower As Long = 0, _
            Optional doTabs As Long = 0, Optional ByVal dispseg As Long = 1)

*/

namespace CS_test
{
    class CAssembler
    {
        public static string ErrorMessage = "";

        [StructLayout(LayoutKind.Sequential, Pack = 1),Serializable]
        private struct t_asmmodel
        {
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 15)]
            public byte[] code;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 15)]
            public byte[] mask;
            public Int32 length;
            public Int32 jmpsize;
            public Int32 jmpoffset;
            public Int32 jmppos;
        }

        [DllImport("olly.dll")]
        private static extern int Assemble(
            string cmd,
            Int32 ip,
            out t_asmmodel model,
            Int32 attempt,
            Int32 constSize,
            string errText
            );

        public static int Asm(string cmd, Int32 offset, ref byte[] buf)
        {
            List<byte> tmp = new List<byte>();
            int ret = Asm(cmd, offset, ref tmp);
            if (ret<1) return 0;
            buf = tmp.ToArray();
            return ret;
        }

        public static int Asm(string cmd, Int32 offset, ref List<byte> buf)
        {
            ErrorMessage = new String(' ', 256);
            t_asmmodel am = new t_asmmodel();
            int asmLen;

            asmLen = Assemble(cmd, offset, out am, 0, 0, ErrorMessage);
            if (asmLen < 1) return 0;

            ErrorMessage = "";
            for (int i = 0; i < asmLen; i++) buf.Add(am.code[i]);
            return asmLen;

        }

        public static int AsmBlock(List<string> cmds, Int32 offset, ref List<byte> buf)
        {
            int curOffset = offset;
            foreach (string s in cmds)
            {
                int sz = Asm(s, curOffset, ref buf);
                if (sz < 1) return 0;
                curOffset += sz;
            }

            return buf.Count;

        }



    }
}

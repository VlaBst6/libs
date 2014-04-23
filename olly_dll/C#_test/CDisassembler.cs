using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace CS_test
{
    class CInstruction
    {
        public string dump;
        public Int32 offset;
        public string command;
        public Int32 instLen;
    }

    class CDisassembler
    {
        [StructLayout(LayoutKind.Sequential, Pack = 1), Serializable]
        public struct t_Disasm { //                // Results of disassembling
              public Int32 ip;                     // Instrucion pointer
              [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 256)]
              public string dump;                   // Hexadecimal dump of the command
              [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 256)]
              public string result;
              [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 256)]
              public string comment;                 // Brief comment
              public Int32 cmdtype;                 // One of C_xxx
              public Int32 memtype;                 // Type of addressed variable in memory
              public Int32 nprefix;                 // Number of prefixes
              public Int32 indexed;                 // Address contains register(s)
              public Int32 jmpconst;                // Constant jump address
              public Int32 jmptable;                // Possible address of switch table
              public Int32 adrconst;                // Constant part of address
              public Int32 immconst;                // Immediate constant
              public Int32 zeroconst;               // Whether contains zero constant
              public Int32 fixupoffset;             // Possible offset of 32-bit fixups
              public Int32 fixupsize;               // Possible total size of fixups or 0
              public Int32 error;                   // Error while disassembling command
              public Int32 warnings;                 // Combination of DAW_xxx
        }

        public enum disasmMode{
            DISASM_SIZE = 0,                 // Determine command size only
            DISASM_DATA = 1,                 // Determine size and analysis data
            DISASM_FILE = 3,                 // Disassembly, no symbols
            DISASM_CODE = 4                  // Full disassembly
        }

        //Private Declare Function disasm Lib "olly.dll" Alias "Disasm" ( _
        //      ByRef src As Byte, ByVal srcsize As Long, ByVal ip As Long, _
        //      disasm As t_Disasm, dMode As disasmMode) As Long

        [System.Runtime.InteropServices.DllImportAttribute("olly.dll", EntryPoint = "Disasm")]
        private static extern int Disasm(
            IntPtr bytes,
            Int32 srcsize,
            Int32 ip,
            out t_Disasm dsm,
            disasmMode dm
            );

        //todo error handling and more prototypes...
        public static unsafe CInstruction Dsm(byte[] buf, Int32 bufOffset, Int32 va){

            t_Disasm dsm;
            CInstruction ci = new CInstruction();

            fixed (byte* pByte = &buf[bufOffset])
            {
                IntPtr ipBuf = new IntPtr((void*)pByte);
                Int32 x = Disasm(ipBuf, buf.Length, va, out dsm, disasmMode.DISASM_CODE);
                ci.instLen = x;
                ci.offset = va;
                ci.dump = dsm.dump;
                ci.command = dsm.result;
                return ci;
            }

        }

        public static List<CInstruction> DsmBlock(byte[] buf, Int32 va){

            int bufOffset = 0;
            int curVa = va;
            List<CInstruction> ret = new List<CInstruction>();

            while (bufOffset < buf.Length)
            {
                CInstruction ci = Dsm(buf, bufOffset, curVa);
                if (ci.instLen == 0) break;
                curVa += ci.instLen;
                bufOffset += ci.instLen;
                ret.Add(ci);
            }

            return ret;
        }


    }
}

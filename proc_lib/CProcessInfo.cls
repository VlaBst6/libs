VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcessLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

'Used in several projects do not change interface!
'
'8.25.12
'  changed interface, added cx64 class dependancy to support info on x64 processes
'  changed GetRunningProcesses to use WTSQuerySessionInformation for better support across privledge levels/os
'  fixed GetProcessCmdLine
'  Fixed GetUsername to work across more privledge levels/os
'  added is64Bit, domain, fullpath, sessionID to CProcess

' x64 safe functions..
'    DumpProcess, DumpMemory, GetProcessModules, EnumDrivers, GetRunningProcesses
'
' x64unsafe: GetMemoryMap,ReadMemory

Private Declare Function NtQuerySystemInformation Lib "ntdll.dll" ( _
    ByVal dwInfoType As Long, _
    ByRef lpStructure As Any, _
    ByVal dwSize As Long, _
    dwReserved As Long) As Long

Private Type SYSTEM_MODULE_INFORMATION
    Reserved(1) As Long                'ULONG reserved[2];
    Base As Long                       'PVOID Base;
    Size As Long                       'ULONG Size;
    Flags As Long                      'ULONG Flags;
    Index As Integer                   'USHORT Index;
    unkn As Integer                    'USHORT Unknown;
    lcount As Integer                  'USHORT LoadCount;
    modoffset As Integer               'USHORT ModuleNameOffset;
    ImageName As String * 256          'CHAR ImageName[256];
End Type

Private Const SYSMODINFO_SPECIFIER = 11

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Enum ProcessAccessTypes
  PROCESS_TERMINATE = (&H1)
  PROCESS_CREATE_THREAD = (&H2)
  PROCESS_SET_SESSIONID = (&H4)
  PROCESS_VM_OPERATION = (&H8)
  PROCESS_VM_READ = (&H10)
  PROCESS_VM_WRITE = (&H20)
  PROCESS_DUP_HANDLE = (&H40)
  PROCESS_CREATE_PROCESS = (&H80)
  PROCESS_SET_QUOTA = (&H100)
  PROCESS_SET_INFORMATION = (&H200)
  PROCESS_QUERY_INFORMATION = (&H400)
  'STANDARD_RIGHTS_REQUIRED = &HF0000
  SYNCHRONIZE = &H100000
  PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
End Enum

Enum MemoryTypes
    MEM_COMMIT = &H1000
    MEM_RESERVE = &H2000
    MEM_DECOMMIT = &H4000
    MEM_RELEASE = &H8000
    MEM_FREE = &H10000
    MEM_PRIVATE = &H20000
    MEM_MAPPED = &H40000
    MEM_RESET = &H80000
    MEM_TOP_DOWN = &H100000
    MEM_IMAGE = &H1000000
    MEM_PHYSICAL = &H400000
End Enum

Private Type MEMORY_BASIC_INFORMATION
   BaseAddress As Long
   AllocationBase As Long
   InitialProtect As Long
   RegionSize As Long
   State As Long
   Protect As Long
   Type As Long
End Type

Enum MemAccess
        PAGE_EXECUTE_READWRITE = &H40
        PAGE_EXECUTE_READ = &H20
        PAGE_EXECUTE_WRITECOPY = &H80
        PAGE_GUARD = &H100
        PAGE_NOACCESS = &H1
        PAGE_NOCACHE = &H200
        PAGE_READONLY = &H2
        PAGE_READWRITE = &H4
        PAGE_WRITECOMBINE = &H400
        PAGE_WRITECOPY = &H8
End Enum


Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPHEAPLIST = 1
Private Const TH32CS_SNAPPROCESS = 2
Private Const TH32CS_SNAPTHREAD = 4
Private Const TH32CS_SNAPMODULE = 8
Private Const TOKEN_QUERY = 8&
Private Const TOKEN_USER = 1
Private Const SPARE_LEN = 512
'Private Const MEM_RELEASE = &H8000
'Private Const MEM_COMMIT = &H1000
'Private Const MEM_RESERVE = &H2000
'Private Const PAGE_EXECUTE_READWRITE = &H40
'Private Const PAGE_EXECUTE_READ = &H20
'Private Const PAGE_EXECUTE_WRITECOPY = &H80
'Private Const PAGE_GUARD = &H100
'Private Const PAGE_NOACCESS = &H1
'Private Const PAGE_NOCACHE = &H200
'Private Const PAGE_READONLY = &H2
'Private Const PAGE_READWRITE = &H4
'Private Const PAGE_WRITECOMBINE = &H400
'Private Const PAGE_WRITECOPY = &H8
'Private Const MEM_MAPPED = &H40000
'Private Const MEM_IMAGE = &H1000000
'Private Const MEM_PRIVATE = &H20000
'Private Const MEM_PHYSICAL = &H400000
'Private Const PROCESS_TERMINATE = 1

Private Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function WriteProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessBytes Lib "Kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessLong Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, outVal As Long, ByVal nSize As Long, lpNumberOfBytesRead As Long) As Long

Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleInformation Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByRef lpmodinfo As LPMODULEINFO, ByVal cb As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ResumeThread Lib "Kernel32" (ByVal hThread As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
'Private Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Declare Function VirtualQueryEx Lib "Kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long

Private Declare Function CreateRemoteThread Lib "Kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function VirtualAllocEx Lib "Kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Private Declare Function CreateProcess Lib "Kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long

'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcessImageFileNameA Lib "PSAPI.DLL" _
        (ByVal hProcess As Long, _
         ByVal lpImageName As String, _
         ByVal nSize As Long) As Long

Private Declare Function NtQueryInformationProcess Lib "ntdll" ( _
    ByVal ProcessHandle As Long, ByVal InformationClass As Long, _
    ByRef ProcessInformation As Any, ByVal ProcessInformationLength As Long, _
    ByRef ReturnLength As Any) As Long

Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Declare Function WTSQuerySessionInformation _
    Lib "wtsapi32" Alias "WTSQuerySessionInformationW" ( _
    ByVal hServer As Long, _
    ByVal sessionID As Long, _
    ByVal WTSInfoClass As Long, _
    ByRef address As Long, _
    ByRef pBytesReturned As Long _
) As Long

'Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" _
'                                        (ByVal lpSystemName As String, _
'                                         ByVal sID As Long, _
'                                         ByVal name As String, _
'                                         cbName As Long, _
'                                         ByVal ReferencedDomainName As String, _
'                                         cbReferencedDomainName As Long, _
'                                         peUse As Integer) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Sub ProcessIdToSessionId Lib "kernel32.dll" (ByVal lngPID As Long, ByRef lngSID As Long)

Private Declare Function QueryDosDeviceW Lib "kernel32.dll" ( _
    ByVal lpDeviceName As Long, _
    ByVal lpTargetPath As Long, _
    ByVal ucchMax As Long _
    ) As Long
Private Const MAX_PATH = 260

Private Declare Function GetLogicalDriveStringsA Lib "Kernel32" ( _
    ByVal nBufferLength As Long, lpBuffer As Any) As Long


Private Type WTS_PROCESS_INFO
   sessionID As Long
   ProcessID As Long
   pProcessName As Long
   pUserSid As Long
End Type

Private Declare Function WTSEnumerateProcesses _
   Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" _
   (ByVal hServer As Long, ByVal Reserved As Long, _
   ByVal Version As Long, ByRef ppProcessInfo As Long, _
   ByRef pCount As Long _
   ) As Long

Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" _
   (ByVal pMemory As Long)

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'   (Destination As Any, Source As Any, ByVal length As Long)

'Private Declare Function LookupAccountName _
'   Lib "advapi32.dll" Alias "LookupAccountNameA" _
'   (ByVal lpSystemName As String, _
'    ByVal lpAccountName As String, _
'    sID As Any, _
'    cbSid As Long, _
'    ReferencedDomainName As Any, _
'    cbReferencedDomainName As Long, _
'    peUse As Integer) As Long
    
Private Const ProcessBasicInformation = 0

Private Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type

Private Type SID_AND_ATTRIBUTES
       sID As Long
       Attributes As Long
       Spare(SPARE_LEN) As Byte
End Type

Private Type LPMODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type

Private Type strucPEB
    filler(15) As Byte
    InfoBlockAdderss As Long
End Type

Private Type INFOBLOCK_
    filler(67) As Byte
    wszCmdLineAddress As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Public SeDebugEnabled As Boolean
Public x64 As New Cx64

Function ShowDllsFor(pid As Long, Optional owner As Object)
    frmDlls.ShowDllsFor pid, owner
End Function

Function SelectProcess() As CProcess
    Set SelectProcess = frmListProcess.SelectProcess(GetRunningProcesses())
End Function

Function RunCmdGetOutput(cmdLine As String, Optional blnStdOut As Boolean = True, Optional blnStdErr As Boolean = False) As String
    On Error Resume Next
    RunCmdGetOutput = CmdOutput.GetCommandOutput(cmdLine, blnStdOut, blnStdErr)
End Function

Function GetRunningProcesses() As Collection 'of CProcess classes
    
    Dim m_col As New Collection
    Dim myProcess As PROCESSENTRY32
    Dim mySnapshot As Long
    Dim proc As CProcess
    Dim n As Long
    
    Dim RetVal As Long
    Dim Count As Long
    Dim i As Integer
    Dim lpBuffer As Long
    Dim p As Long
    Dim udtProcessInfo As WTS_PROCESS_INFO
    'Dim itmAdd As ListItem

    
    RetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
   
    If RetVal Then
          p = lpBuffer
          For i = 1 To Count
                CopyMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
                Set proc = New CProcess
                With proc
                    .sessionID = udtProcessInfo.sessionID
                    .pid = udtProcessInfo.ProcessID
                    .path = GetStringFromLP(udtProcessInfo.pProcessName)
                    .fullpath = GetProcessPath(.pid)
                    .ParentPID = GetParentProcessId(.pid)
                    .cmdLine = GetProcessCmdLine(.pid, .fullpath)
                    GetUserName udtProcessInfo.pUserSid, proc
                    .is64Bit = (x64.IsProcess_x64(.pid) = r_64bit)
                    m_col.Add proc, "pid:" & .pid
                End With
                         
                p = p + LenB(udtProcessInfo)
          Next i
    
          'Set itmAdd = Nothing
          WTSFreeMemory lpBuffer   'Free your memory buffer
   Else
        Set m_col = Legacy_GetRunningProcesses() 'win2k
   End If
   
   Set GetRunningProcesses = m_col
    
End Function

Function Legacy_GetRunningProcesses() As Collection 'of CProcess classes

    Dim m_col As New Collection
    Dim myProcess As PROCESSENTRY32
    Dim mySnapshot As Long
    Dim proc As CProcess
    Dim n As Long

    myProcess.dwSize = Len(myProcess)
    mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

    ProcessFirst mySnapshot, myProcess

    If myProcess.th32ProcessID <> 0 Then
        Set proc = New CProcess
        With proc
            .path = myProcess.szexeFile
             If InStr(.path, Chr(0)) > 0 Then
                .path = Mid(.path, 1, InStr(.path, Chr(0)) - 1)
             End If
            .fullpath = GetProcessPath(myProcess.th32ProcessID)
            .pid = myProcess.th32ProcessID
            .ParentPID = myProcess.th32ParentProcessID
            .cmdLine = GetProcessCmdLine(.pid, .fullpath)
            .User = GetProcessUser(.pid)
            .is64Bit = (x64.IsProcess_x64(.pid) = r_64bit)
        End With

        m_col.Add proc, "pid:" & myProcess.th32ProcessID
    End If

    While ProcessNext(mySnapshot, myProcess)

        If myProcess.th32ProcessID <> 4 Then
            Set proc = New CProcess
            With proc
                .path = myProcess.szexeFile
                 If InStr(.path, Chr(0)) > 0 Then
                    .path = Mid(.path, 1, InStr(.path, Chr(0)) - 1)
                 End If
                .fullpath = GetProcessPath(myProcess.th32ProcessID)
                .pid = myProcess.th32ProcessID
                .ParentPID = myProcess.th32ParentProcessID
                .cmdLine = GetProcessCmdLine(.pid, .fullpath)
                .User = GetProcessUser(.pid) 'OpenProcessToken Fails for Local/Network Service
                'If Len(.User) = 0 Then Stop
                .is64Bit = (x64.IsProcess_x64(.pid) = r_64bit)
            End With

            n = InStr(proc.path, Chr(0))
            If n > 1 Then proc.path = Mid(proc.path, 1, n - 1)
            proc.path = Replace(proc.path, "\??\", Empty)
            proc.path = Replace(proc.path, "\SystemRoot", Environ("Windir"))

            m_col.Add proc, "pid:" & myProcess.th32ProcessID
        End If

    Wend

    Set Legacy_GetRunningProcesses = m_col

End Function

'using static PEB address will fail for XP SP 2 and above...
Function GetProcessCmdLine(pid As Long, Optional fullpath As String) As String

    'Exit Function 'fix me sometime
    
   'Converted from C Source by Matt pietrek in 1997 MSDJ
   '1. find the Process Environment Block
   
   Dim hProcess As Long
   Dim PBI As PROCESS_BASIC_INFORMATION
   Dim r As Long
   Dim dwSize As Long
   Dim peb As strucPEB
   Dim udt() As Byte
   ReDim udt(Len(peb))
   
   hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, False, pid)
   'pbi.PebBaseAddress = &H7FFDF000 'unsafe...
    
   r = NtQueryInformationProcess(hProcess, ProcessBasicInformation, PBI, Len(PBI), Len(PBI))
   
   If ReadProcessMemory(hProcess, PBI.PebBaseAddress, udt(0), Len(peb), dwSize) < 0 Then
       'MsgBox Err.LastDllError
       CloseHandle hProcess
       Exit Function
   End If
    
   'Debug.Print "Peb: " & hexdump(StrConv(udt, vbUnicode))
    
   CopyMemory peb, udt(0), Len(peb)

   '2. from this PEB, get the address of the block containing a pointer to the CmdLine
   Dim Block As INFOBLOCK_
   ReDim udt(Len(Block))
    
   If ReadProcessMemory(hProcess, peb.InfoBlockAdderss, udt(0), Len(Block), dwSize) < 0 Then
          'MsgBox Err.LastDllError
          CloseHandle hProcess
          Exit Function
   End If

   CopyMemory Block, udt(0), Len(Block)
      
    '3. get the CmdLine
   Dim wszCmdLine(500) As Byte
   Dim cmdLine As String
     
   If ReadProcessMemory(hProcess, Block.wszCmdLineAddress, wszCmdLine(0), UBound(wszCmdLine), dwSize) < 0 Then
            'MsgBox Err.LastDllError
            CloseHandle hProcess
            Exit Function
   End If
    
    Dim i As Long
    
   For i = 0 To UBound(wszCmdLine) - 1
        If wszCmdLine(i) = 0 And wszCmdLine(i + 1) = 0 Then Exit For
        If wszCmdLine(i) <> 0 Then cmdLine = cmdLine & Chr(wszCmdLine(i))
   Next
     
   cmdLine = Replace(cmdLine, """" & fullpath & """", Empty, , , vbTextCompare)
   cmdLine = Replace(cmdLine, fullpath, Empty, , , vbTextCompare)
     
   GetProcessCmdLine = Trim(cmdLine)
   
   CloseHandle hProcess
   
End Function

Function GetProcessModules(pid As Long) As Collection

    On Error GoTo hell
    
    Dim hMod() As Long
    Dim dwSize As Long
    Dim hProcess As Long
    Dim i As Long
    Dim lmi As LPMODULEINFO
    Dim mfn As String
    Dim nLen As Long
    Dim mCol As New Collection
    Dim cmod As CModule
    
    If x64.IsProcess_x64(pid) = r_64bit Then
        Set GetProcessModules = x64.GetProcessModules(pid)
        Exit Function
    End If
    
    ReDim hMod(200)
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    EnumProcessModules hProcess, hMod(0), 1000, dwSize
     
    ReDim Preserve hMod(dwSize / 4 - 1)
    
    For i = 0 To UBound(hMod)
        mfn = Space(260)
        nLen = GetModuleFileNameExA(hProcess, hMod(i), mfn, 260)
        mfn = Left(mfn, nLen)
        
        Call GetModuleInformation(hProcess, hMod(i), lmi, Len(lmi))
         
        Set cmod = New CModule
        cmod.path = mfn
        cmod.Base = lmi.lpBaseOfDll
        cmod.Size = lmi.SizeOfImage
         
        On Error Resume Next
        mCol.Add cmod, cmod.path
    Next
    
    CloseHandle hProcess
    Set GetProcessModules = mCol
    
    Exit Function
hell:
End Function

'this will fail for Local Service and Network Service processes even with Admin and sedebug...
'use the data in cprocess instead if possible..
Function GetProcessUser(pid As Long) As String
   'Microsoft Visual Basic Developer Support
   'This posting is provided “AS IS” with no warranties, and confers no rights.
   '(c) 2002 Microsoft Corporation.  All rights reserved
   Dim hProcess As Long
   Dim hAccessProcessToken As Long
   Dim sid_name_use As Integer
   Dim User As String
   Dim domain As String
   Dim lu As Long
   Dim ld As Long
   Dim success As Long
   Dim ErrNo As Integer
   Dim InfoStructure As SID_AND_ATTRIBUTES
   Dim LenInfo As Long
   Dim sUser As String, sDomain As String

   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
   success = OpenProcessToken(hProcess, TOKEN_QUERY, hAccessProcessToken)

   If success Then
   
       success = GetTokenInformation(hAccessProcessToken, TOKEN_USER, InfoStructure, SPARE_LEN + 4, LenInfo)

       If success = 0 Then
           ErrNo = Err.LastDllError
       Else

           User = String$(100, 0)
           domain = String$(100, 0)
           lu = Len(User)
           ld = Len(domain)
           success = LookupAccountSid(vbNullString, InfoStructure.sID, User, lu, domain, ld, sid_name_use)

           If success = 0 Then
               ErrNo = Err.LastDllError
           Else
               sUser = Left$(User, InStr(User, Chr$(0)) - 1)
               sDomain = Left$(domain, InStr(domain, Chr$(0)) - 1)
               GetProcessUser = sDomain & ":" & sUser
           End If
         
       End If

   Else

       ErrNo = Err.LastDllError

   End If

   CloseHandle hProcess
    
End Function

Function DumpProcess(pid As Long, fpath As String) As Boolean
    
    If x64.IsProcess_x64(pid) = r_64bit Then
        DumpProcess = x64.DumpProcess(pid, fpath)
        Exit Function
    End If
    
    Dim cmod As CModule
    Dim col As Collection
    
    Set col = GetProcessModules(pid)
    Set cmod = col(1)

    DumpProcess = DumpProcessMemory(pid, cmod.Base, cmod.Size, fpath)
    
End Function

Function DumpMemory(pid As Long, hex_start As String, hex_size As String, ByVal fpath As String) As Boolean
    
    Dim s As Long, l As Long
    
    If x64.IsProcess_x64(pid) = r_64bit Then
        DumpMemory = x64.DumpMemory(pid, Trim(hex_start), Trim(hex_size), fpath)
    Else
        s = CLng("&h" & Trim(hex_start))
        l = CLng("&h" & Trim(hex_size))
        DumpMemory = DumpProcessMemory(pid, s, l, fpath)
    End If
    
End Function

Function DumpProcessMemory(pid As Long, start As Long, Length As Long, fpath As String) As Boolean
   
    Dim pHandle As Long
    Dim b() As Byte
    Dim f As Long
    On Error GoTo hell
    
    f = FreeFile
    
    ReDim b(Length + 1)
        
    pHandle = OpenProcess(PROCESS_VM_READ, False, pid)
    If pHandle = 0 Then Exit Function
    
    ReadProcessMemory pHandle, start, b(0), Length, Length
    CloseHandle pHandle
    
    If Dir(fpath) <> "" Then Kill fpath
    
    Open fpath For Binary As f
    Put f, , b()
    Close f
    DumpProcessMemory = True
    
    Exit Function
hell:
    DumpProcessMemory = False
    
End Function

Function ReadMemory(pid As Long, start As Long, Length As Long) As String
   
    Dim pHandle As Long
    Dim b() As Byte
    Dim f As Long
   
    f = FreeFile
    
    ReDim b(Length + 1)
        
    pHandle = OpenProcess(PROCESS_VM_READ, False, pid)
    ReadProcessMemory pHandle, start, b(0), Length, Length
    CloseHandle pHandle
    
    ReadMemory = StrConv(b(), vbUnicode, &H409)
    
End Function

'less restrictions on GetProcessImageFileName ?
Function GetProcessPath(pid As Long) As String
    Dim hProc As Long
    Dim hMods() As Long, ret As Long, retMax As Long
    Dim sPath As String
    Dim n As Long
    
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, pid) '2k and XP
    If hProc <> 0 Then
        ReDim hMods(900)
        ret = EnumProcessModules(hProc, hMods(0), 900, retMax)
        sPath = Space$(260)
        ret = GetModuleFileNameExA(hProc, hMods(0), sPath, 260)
        GetProcessPath = Left$(sPath, ret)
        GetProcessPath = Replace(GetProcessPath, "\??\", Empty)
        GetProcessPath = Replace(GetProcessPath, "\SystemRoot", Environ("Windir"))
        n = InStr(GetProcessPath, Chr(0))
        If n > 1 Then GetProcessPath = Mid(GetProcessPath, 1, n - 1)
        
    End If
    
    If hProc <> 0 Then CloseHandle hProc
    
    If Len(GetProcessPath) = 0 Then 'handles more permission issues in Vista+
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
        If hProc = 0 Then Exit Function
        sPath = String(1024, " ")
        ret = GetProcAddress(LoadLibrary("psapi.dll"), "GetProcessImageFileNameA")
        If ret = 0 Then
            ret = GetProcAddress(LoadLibrary("kernel32.dll"), "GetProcessImageFileNameA") 'Vista+
        End If
        If ret = 0 Then Exit Function
        ret = GetProcessImageFileNameA(hProc, sPath, 1024) 'this returns NT style names \DeviceHardDiskVolume1\...
        If ret <> 0 Then
            GetProcessPath = Left$(sPath, ret)
            n = InStr(GetProcessPath, Chr(0))
            If n > 1 Then GetProcessPath = Mid(GetProcessPath, 1, n - 1)
            sPath = GetDosPathForNtDevicePath(GetProcessPath)
            If Len(sPath) > 0 Then GetProcessPath = sPath
        End If
        CloseHandle hProc
    End If
    
End Function

Function EnumDrivers() As Collection
  Dim entries As Long
  Dim numBytes As Long
  Dim BUFSIZE As Long
  Dim buf() As Byte
  Dim smi As SYSTEM_MODULE_INFORMATION
  Dim offset As Long
  Dim i As Long
  Dim tmp As String
  
  'get num entries to calc buffer size
  NtQuerySystemInformation SYSMODINFO_SPECIFIER, entries, 4, numBytes
    
  'allocate buffer for num strucs returned
  BUFSIZE = Len(smi) * (entries + 1)
  ReDim buf(BUFSIZE)

  'load buffer with all structures
  NtQuerySystemInformation SYSMODINFO_SPECIFIER, buf(0), BUFSIZE, numBytes
     
  'first 4 bytes are num Entries returned
  CopyMemory entries, buf(0), 4
    
  offset = 4
  
  Dim ret As New Collection
  
  On Error Resume Next
  For i = 1 To entries
        CopyMemory smi, buf(offset), Len(smi)
        
        BUFSIZE = InStr(smi.ImageName, Chr(0))
        If BUFSIZE > 0 Then
            tmp = Mid(smi.ImageName, 1, BUFSIZE - 1)
        Else
            tmp = smi.ImageName
        End If
        
        ret.Add tmp, tmp
        offset = offset + Len(smi)
  Next
    
  Set EnumDrivers = ret
  
End Function

Function TerminateProces(pid As Long) As Boolean
    Dim pHandle As Long
    pHandle = OpenProcess(PROCESS_TERMINATE, 0, pid)
    TerminateProces = CBool(TerminateProcess(pHandle, 0))
    CloseHandle pHandle
End Function

Function GetSeDebug() As Boolean
    Dim hToken As Long, hProcess As Long, lRet As Long
    Dim tkp As TOKEN_PRIVILEGES
    
    Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
    Const TOKEN_QUERY As Long = &H8

    hProcess = GetCurrentProcess()
    OpenProcessToken hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeDebugPrivilege", tkp.TheLuid

    tkp.PrivilegeCount = 1
    tkp.Attributes = 2 'SE_PRIVILEGE_ENABLED
    
    If AdjustTokenPrivileges(hToken, False, tkp, Len(tkp), tkp, lRet) = 0 Then Exit Function
        
    If GetLastError = 0 Then GetSeDebug = True

End Function

Private Sub Class_Initialize()
    SeDebugEnabled = GetSeDebug()
End Sub

Function GetMemoryMap(pid As Long) As Collection 'of CMem

    Dim hProc As Long
    Dim x As Long
    Dim meminfo As MEMORY_BASIC_INFORMATION
    Dim cur As Long
    Dim ret As New Collection
    Dim n As String
    Dim cMem As CMemory
    
    hProc = OpenProcess(CLng(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ), False, pid)
    If hProc = 0 Then
        Set GetMemoryMap = ret
        x = GetLastError
        Exit Function
    End If
           
    Do While cur < &H7FFEFFFF
         x = VirtualQueryEx(hProc, cur, meminfo, LenB(meminfo))
         
         If meminfo.State <> MEM_FREE Then
            Set cMem = New CMemory
            cMem.pid = pid
            cMem.Size = meminfo.RegionSize
            cMem.Base = cur
            cMem.AllocBase = meminfo.AllocationBase
            cMem.InitialProtection = meminfo.InitialProtect
            cMem.MemType = meminfo.Type
            cMem.Protection = meminfo.Protect
            cMem.State = meminfo.State
                       
            If cur > 0 Then
               n = String(256, Chr(0))
               x = GetModuleFileNameExA(hProc, cur, n, 255)
               If x <> 0 Then cMem.ModuleName = ToNull(n)
            End If
            
            ret.Add cMem
            
        End If
        
        cur = cur + meminfo.RegionSize
        DoEvents
        'sleep 10
        
    Loop
       
    CloseHandle hProc
    Set GetMemoryMap = ret
    
End Function

Private Function ToNull(v As String) As String
    Dim x As Integer
    ToNull = v
    x = InStr(v, Chr(0))
    If x > 0 Then ToNull = Left(v, x - 1)
End Function

Private Function GetStringFromLP(ByVal StrPtr As Long) As String
   Dim b As Byte
   Dim tempStr As String
   Dim bufferStr As String
   Dim Done As Boolean

   Done = False
   Do
      ' Get the byte/character that StrPtr is pointing to.
      CopyMemory b, ByVal StrPtr, 1
      If b = 0 Then  ' If you've found a null character, then you're done.
         Done = True
      Else
         tempStr = Chr$(b)  ' Get the character for the byte's value
         bufferStr = bufferStr & tempStr 'Add it to the string
                
         StrPtr = StrPtr + 1  ' Increment the pointer to next byte/char
      End If
   Loop Until Done
   GetStringFromLP = bufferStr
End Function

Private Function GetUserName(sID As Long, p As CProcess) As String
    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    LookupAccountSid vbNullString, sID, retname, 255, retdomain, 255, 0
    p.domain = Left$(retdomain, InStr(retdomain, vbNullChar) - 1)
    p.User = Left$(retname, InStr(retname, vbNullChar) - 1)
    GetUserName = p.domain & "\" & p.User
End Function

Public Function GetParentProcessId(ByVal pid As Long) As Long
    Dim PBI As PROCESS_BASIC_INFORMATION
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, False, pid)
    Call NtQueryInformationProcess(hProcess, ProcessBasicInformation, PBI, Len(PBI), ByVal 0&)
    GetParentProcessId = PBI.InheritedFromUniqueProcessId
    CloseHandle hProcess
End Function

Function GetProcess(pid As Long, cp As CProcess) As Boolean
    Dim c As Collection
    
    Set c = GetRunningProcesses()
    
    For Each cp In c
        If cp.pid = pid Then
            GetProcess = True
            Exit Function
        End If
    Next
    
    Set cp = Nothing
    
End Function



Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

'64 bit safe
Public Function InjectDLL(pid As Long, dllPath As String, Optional injLog As String, Optional cp As CProcess) As Boolean

    Dim hProcess As Long
    Dim lpfnLoadLib As Long
    Dim ret As Long
    Dim lpdllPath As Long
    Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFO
    Dim hThread As Long
    Dim writeLen As Long
    Dim b() As Byte
    Dim buflen As Long
    'Dim li As ListItem
    Dim c As Collection
    Dim found As Boolean
    Dim rv As Boolean
    
    Const PAGE_READWRITE = 4
    Const CREATE_SUSPENDED = &H4
    Const MEM_COMMIT = &H1000
    
    injLog = Empty
    
    If Not FileExists(dllPath) Then
        injLog = "DllPath not found: " & dllPath
        GoTo hell
    End If
    
    If Not GetProcess(pid, cp) Then
        injLog = "pid " & pid & " not found"
        GoTo hell
    End If
    
    If cp.is64Bit Then
        If x64.isExe_x64(dllPath) <> r_64bit Then
            injLog = "Can not inject a 32 bit dll into a 64 bit process."
            GoTo hell
        End If
        rv = x64.x64Inject(pid, dllPath, injLog)
        GoTo hell
    End If
    
    If x64.isExe_x64(dllPath) <> r_32bit Then
            injLog = "Can not inject a 64 bit dll in a 32 bit process."
            GoTo hell
    End If
     
    
    b() = StrConv(dllPath & Chr(0), vbFromUnicode)
    buflen = UBound(b) + 1
 
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    injLog = injLog & "Opening PID: " & pid & " Process Handle=" & hProcess & vbCrLf
    If hProcess = -1 Then GoTo hell
     
    lpdllPath = VirtualAllocEx(hProcess, ByVal 0, buflen, MEM_COMMIT, ByVal PAGE_READWRITE)
    injLog = injLog & "Remote Allocation base: " & Hex(lpdllPath) & vbCrLf
    If lpdllPath = 0 Then GoTo hell
    
    ret = WriteProcessBytes(hProcess, ByVal lpdllPath, b(0), buflen, writeLen)
    injLog = injLog & "WriteProcessMemory=" & ret & " BufLen=" & buflen & " Bytes Written: " & writeLen & vbCrLf
    If ret = 0 Then GoTo hell
     
    lpfnLoadLib = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
    injLog = injLog & "LoadLibraryA = " & Hex(lpfnLoadLib) & vbCrLf
    
    'DebugBreak
    ret = CreateRemoteThread(hProcess, ByVal 0, 0, lpfnLoadLib, lpdllPath, 0, hThread)
    injLog = injLog & "CreateRemoteThread = " & ret & " ThreadID: " & Hex(hThread) & vbCrLf
    If ret = -1 Then GoTo hell
    
    Sleep 300
    rv = True
    
hell:
   CloseHandle hProcess
   InjectDLL = rv
    
End Function

'64 bit safe..
Public Function StartProcessWithDLL(exePath As String, dllPath As String, Optional injLog As String, Optional cp As CProcess) As Boolean

    Dim hProcess As Long
    Dim lpfnLoadLib As Long
    Dim ret As Long
    Dim lpdllPath As Long
    Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFO
    Dim hThread As Long
    Dim writeLen As Long
    Dim b() As Byte
    Dim buflen As Long
    Dim rv As Boolean
    
    Const PAGE_READWRITE = 4
    Const CREATE_SUSPENDED = &H4
    Const MEM_COMMIT = &H1000
    
    injLog = Empty
    
    If Not FileExists(dllPath) Then
        injLog = "DllPath not found: " & dllPath
        GoTo hell
    End If
    
    If Not FileExists(exePath) Then
        injLog = "exePath not found: " & exePath
        GoTo hell
    End If
    
    Set cp = New CProcess
    cp.fullpath = exePath
    cp.path = exePath
    
    If x64.isExe_x64(exePath) = r_64bit Then
        If x64.isExe_x64(dllPath) <> r_64bit Then
            injLog = "Can not start a 32 bit dll in a 64 bit executable."
            GoTo hell
        End If
        rv = x64.x64StartWithDll(exePath, dllPath, injLog)
        If rv Then cp.pid = parsePid(injLog)
        GoTo hell
    End If
    
    If x64.isExe_x64(dllPath) <> r_32bit Then
            injLog = "Can not start a 64 bit dll in a 32 bit executable."
            GoTo hell
    End If
        
    b() = StrConv(dllPath & Chr(0), vbFromUnicode)
    buflen = UBound(b) + 1

    ret = CreateProcess(0&, exePath, 0&, 0&, 1&, CREATE_SUSPENDED, 0&, 0&, si, pi)
    injLog = injLog & "Create Process Suspended: " & ret & IIf(ret = 0, " Failed", " PID: " & pi.dwProcessId) & vbCrLf
    cp.pid = pi.dwProcessId
    If ret = 0 Then GoTo hell
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pi.dwProcessId)
    injLog = injLog & "OpenProcess Handle=" & hProcess & vbCrLf
    If hProcess = -1 Then GoTo hell
    
    lpdllPath = VirtualAllocEx(hProcess, ByVal 0, buflen, MEM_COMMIT, ByVal PAGE_READWRITE)
    injLog = injLog & "Remote Allocation base: " & Hex(lpdllPath) & vbCrLf
    If lpdllPath = 0 Then GoTo hell
    
    ret = WriteProcessBytes(hProcess, ByVal lpdllPath, b(0), buflen, writeLen)
    injLog = injLog & "WriteProcessMemory=" & ret & " BufLen=" & buflen & " Bytes Written: " & writeLen & vbCrLf
    If ret = 0 Then GoTo hell
     
    lpfnLoadLib = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
    injLog = injLog & "LoadLibraryA = " & Hex(lpfnLoadLib) & vbCrLf
    
    'DebugBreak
    ret = CreateRemoteThread(hProcess, ByVal 0, 0, lpfnLoadLib, lpdllPath, 0, hThread)
    injLog = injLog & "CreateRemoteThread = " & ret & " ThreadID: " & Hex(hThread) & vbCrLf
    If ret = -1 Then GoTo hell
    
    Sleep 300
    ret = ResumeThread(pi.hThread)
    injLog = injLog & "ResumeThread = " & ret & vbCrLf
    If ret = -1 Then GoTo hell

    rv = True
    
hell:
    CloseHandle hProcess
    StartProcessWithDLL = rv
    
End Function

Private Function parsePid(injLog As String) As Long
    On Error Resume Next
    Dim a As Long, b As Long, tmp
    a = InStr(injLog, "pid")
    If a > 0 Then
        a = a + 3
        b = InStr(a, injLog, ",")
        If b > 0 Then
            tmp = Mid(injLog, a, b - a)
            tmp = Trim(Replace(tmp, "=", Empty))
            parsePid = CLng(tmp)
        End If
    End If
End Function

'-----------------------------------------------------------
'http://www.vbaccelerator.com/home/VB/Tips/Mapping_NT_Device_Names/article.asp
'-----------------------------------------------------------
Private Function GetNtDeviceNameForDrive( _
   ByVal sDrive As String) As String
Dim bDrive() As Byte
Dim bResult() As Byte
Dim lR As Long
Dim sDeviceName As String

   If Right(sDrive, 1) = "\" Then
      If Len(sDrive) > 1 Then
         sDrive = Left(sDrive, Len(sDrive) - 1)
      End If
   End If
   bDrive = sDrive
   ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
   ReDim bResult(0 To MAX_PATH * 2 + 1) As Byte
   lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bResult(0)), MAX_PATH)
   If (lR > 2) Then
      sDeviceName = bResult
      sDeviceName = Left(sDeviceName, lR - 2)
      GetNtDeviceNameForDrive = sDeviceName
   End If
   
End Function

Private Function GetDrives() As Collection
Dim colDrives As New Collection
Dim lSize As Long
Dim lR As Long
Dim iLastPos As Long
Dim iPos As Long
Dim sDrive As String
Dim sDriveStrings As String

   lSize = GetLogicalDriveStringsA(0, ByVal 0&)
   sDriveStrings = String(lSize + 1, 0)
   lR = GetLogicalDriveStringsA(lSize, ByVal sDriveStrings)
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, sDriveStrings, vbNullChar)
      If Not (iPos = 0) Then
         sDrive = Mid$(sDriveStrings, iLastPos, iPos - iLastPos)
         iLastPos = iPos + 1
      Else
         sDrive = Mid$(sDriveStrings, iLastPos)
      End If
      If Len(sDrive) > 0 Then
         colDrives.Add sDrive
      End If
   Loop While Not (iPos = 0)
   Set GetDrives = colDrives
   
End Function

Private Function GetDosPathForNtDevicePath(ByVal sDevicePath As String) As String
Dim sFoundDrive As String
Dim colDrives As Collection
Dim vDrive As Variant
Dim tmp As String

   For Each vDrive In GetDrives()
      tmp = GetNtDeviceNameForDrive(vDrive)
      If InStr(1, sDevicePath, tmp, vbTextCompare) > 0 Then
         GetDosPathForNtDevicePath = Replace(sDevicePath, tmp, vDrive, , 1, vbTextCompare)
         GetDosPathForNtDevicePath = Replace(GetDosPathForNtDevicePath, "\\", "\")
         Exit Function
      End If
   Next
   
End Function
'-------------------------------------------------




Public Function InjectData(pid As Long, data As String, Optional Base As Long = 0, Optional injLog As String, Optional cp As CProcess) As Boolean

    Dim hProcess As Long
    Dim lpfnLoadLib As Long
    Dim ret As Long
    Dim lpdllPath As Long
    Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFO
    Dim hThread As Long
    Dim writeLen As Long
    Dim b() As Byte
    Dim buflen As Long
    Dim c As Collection
    Dim found As Boolean
    Dim rv As Boolean
    
    Const PAGE_READWRITE = 4
    Const CREATE_SUSPENDED = &H4
    Const MEM_COMMIT = &H1000
    
    injLog = Empty
    
    If Not GetProcess(pid, cp) Then
        injLog = "pid " & pid & " not found"
        GoTo hell
    End If
    
    If cp.is64Bit Then
        injLog = "Not implemented for x64 Processes"
        GoTo hell
    End If
    
    b() = StrConv(data, vbFromUnicode, LANG_US)
    buflen = UBound(b) + 1
 
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    injLog = injLog & "Opening PID: " & pid & " Process Handle=" & hProcess & vbCrLf
    If hProcess = -1 Then GoTo hell
     
    lpdllPath = VirtualAllocEx(hProcess, ByVal Base, buflen, MEM_COMMIT, ByVal PAGE_EXECUTE_READWRITE)
    injLog = injLog & "Remote Allocation base: " & Hex(lpdllPath) & vbCrLf
    If lpdllPath = 0 Then GoTo hell
    
    ret = WriteProcessBytes(hProcess, ByVal lpdllPath, b(0), buflen, writeLen)
    injLog = injLog & "WriteProcessMemory=" & ret & " BufLen=" & buflen & " Bytes Written: " & writeLen & vbCrLf
    If ret = 0 Then GoTo hell
     
    rv = True
    
hell:
   CloseHandle hProcess
   InjectData = rv
    
End Function

Function OpenProcessHandle(pid As Long) As Long
    OpenProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
End Function

Function CloseProcessHandle(h As Long)
    CloseHandle h
End Function

Public Function WriteProcessLong(ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal sValue As Long) As Boolean
    Dim ret As Long
    If WriteProcessMemory(hProcess, lpBaseAddress, ByVal VarPtr(sValue), 4, ret) <> 0 Then WriteProcessLong = True
End Function

Function WriteProcessByte(hProc As Long, val As Byte, address As Long) As Boolean
    Dim ret As Long, written As Long
    If WriteProcessBytes(hProc, ByVal address, ByVal val, 1, written) <> 0 Then WriteProcessByte = True
End Function

Function ReadProcessLng(hProc As Long, address As Long, outVal As Long) As Boolean
   
    Dim b(4) As Byte, tmp As Long
    outVal = 0
    
    If ReadProcessMemory(hProc, address, b(0), 4, 0) > 0 Then
        CopyMemory tmp, b(0), 4
        outVal = tmp
        ReadProcessLng = True
    End If
 
End Function


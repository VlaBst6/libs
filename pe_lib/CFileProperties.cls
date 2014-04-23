VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
 

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
End Type


Private Const LANG_BULGARIAN = &H2
Private Const LANG_CHINESE = &H4
Private Const LANG_CROATIAN = &H1A
Private Const LANG_CZECH = &H5
Private Const LANG_DANISH = &H6
Private Const LANG_DUTCH = &H13
Private Const LANG_ENGLISH = &H9
Private Const LANG_FINNISH = &HB
Private Const LANG_FRENCH = &HC
Private Const LANG_GERMAN = &H7
Private Const LANG_GREEK = &H8
Private Const LANG_HUNGARIAN = &HE
Private Const LANG_ICELANDIC = &HF
Private Const LANG_ITALIAN = &H10
Private Const LANG_JAPANESE = &H11
Private Const LANG_KOREAN = &H12
Private Const LANG_NEUTRAL = &H0
Private Const LANG_NORWEGIAN = &H14
Private Const LANG_POLISH = &H15
Private Const LANG_PORTUGUESE = &H16
Private Const LANG_ROMANIAN = &H18
Private Const LANG_RUSSIAN = &H19
Private Const LANG_SLOVAK = &H1B
Private Const LANG_SLOVENIAN = &H24
Private Const LANG_SPANISH = &HA
Private Const LANG_SWEDISH = &H1D
Private Const LANG_TURKISH = &H1F

Function QuickInfo(fileName As String)
    Dim f As FILEPROPERTIE
    
    f = FileInfo(fileName)
    
    QuickInfo = "CompanyName      " & f.CompanyName & vbCrLf & _
                "FileDescription  " & f.FileDescription & vbCrLf & _
                "FileVersion      " & f.FileVersion & vbCrLf & _
                "InternalName     " & f.InternalName & vbCrLf & _
                "LegalCopyright   " & f.LegalCopyright & vbCrLf & _
                "OriginalFilename " & f.OrigionalFileName & vbCrLf & _
                "ProductName      " & f.ProductName & vbCrLf & _
                "ProductVersion   " & FileInfo.ProductVersion
                

End Function

Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then
        Exit Function
    End If
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim bytBuffer() As Byte
    Dim bytBuff() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strVersionInfo(7) As String
    Dim strTemp As String
    Dim intTemp As Integer
           
    ReDim bytBuff(500)
    
    ' size
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen > 0 Then
    
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
       
       If lngRc <> 0 Then
          lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
          If lngRc <> 0 Then
             'lngVerPointer is a pointer to four 4 bytes of Hex number,
             'first two bytes are language id, and last two bytes are code
             'page. However, strLangCharset needs a  string of
             '4 hex digits, the first two characters correspond to the
             'language id and last two the last two character correspond
             'to the code page id.
             MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
             lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
             strLangCharset = Hex(lngHexNumber)
             'now we change the order of the language id and code page
             'and convert it into a string representation.
             'For example, it may look like 040904E4
             'Or to pull it all apart:
             '04------        = SUBLANG_ENGLISH_USA
             '--09----        = LANG_ENGLISH
             ' ----04E4 = 1252 = Codepage for Windows:Multilingual
             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop
             
             If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
               strLangCharset2 = "English (US)"
             End If

             If Mid(strLangCharset, 2, 2) = LANG_BULGARIAN Then strLangCharset2 = "Bulgarian"
             If Mid(strLangCharset, 2, 2) = LANG_FRENCH Then strLangCharset2 = "French"
             If Mid(strLangCharset, 2, 2) = LANG_NEUTRAL Then strLangCharset2 = "Neutral"

             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop

             ' assign propertienames
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             
             Dim n As Long
             
             ' loop and get fileproperties
             For intTemp = 0 To 7
                strBuffer = String$(800, 0)
                strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(intTemp)
                lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   ' get and format data
                   lstrcpy strBuffer, lngVerPointer
                   n = InStr(strBuffer, Chr(0)) - 1
                   If n > 0 Then
                        strBuffer = Mid$(strBuffer, 1, n)
                        strVersionInfo(intTemp) = strBuffer
                   End If
                 Else
                   ' property not found
                   strVersionInfo(intTemp) = ""
                End If
             Next intTemp
             
          End If
       End If
    End If
    
    ' assign array to user-defined-type
    FileInfo.CompanyName = strVersionInfo(0)
    FileInfo.FileDescription = strVersionInfo(1)
    FileInfo.FileVersion = strVersionInfo(2)
    FileInfo.InternalName = strVersionInfo(3)
    FileInfo.LegalCopyright = strVersionInfo(4)
    FileInfo.OrigionalFileName = strVersionInfo(5)
    FileInfo.ProductName = strVersionInfo(6)
    FileInfo.ProductVersion = strVersionInfo(7)
    FileInfo.LanguageID = strLangCharset2
    
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function





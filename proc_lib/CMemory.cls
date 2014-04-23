VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public pid As Long
Public Base As Long
Public Size As Long
Public MemType As MemoryTypes
Public State As MemoryTypes
Public Protection As MemAccess
Public InitialProtection As MemAccess
Public ModuleName As String
Public AllocBase As Long

Public Function IsAddressInRange(address As Long) As Boolean
    Dim maxOffset As Long
    maxOffset = Me.Base + Me.Size
        
    If address >= Me.Base And address <= maxOffset Then
        AddressInRange = True
        Exit Function
    End If
    
End Function

Public Function MemTypeAsString() As String
    MemTypeAsString = GetTypeAsString(MemType)
End Function

Public Function StateAsString() As String
    StateAsString = GetTypeAsString(State)
End Function

Public Function ProtectionAsString(Optional verbose As Boolean = False) As String
    ProtectionAsString = GetProtectAsString(Protection, verbose)
End Function

Public Function InitProtectionAsString() As String
    InitProtectionAsString = GetProtectAsString(InitialProtection, False)
End Function


Private Function GetTypeAsString(x As Long) As String
    
    Select Case x
       Case &H1000: GetTypeAsString = "COMMIT"
       Case &H2000: GetTypeAsString = "RESERVE"
       Case &H4000: GetTypeAsString = "DECOMMIT"
       Case &H8000: GetTypeAsString = "RELEASE"
       Case &H10000: GetTypeAsString = "FREE"
       Case &H20000: GetTypeAsString = "PRIVATE"
       Case &H40000: GetTypeAsString = "MAPPED"
       Case &H80000: GetTypeAsString = "RESET"
       Case &H100000: GetTypeAsString = "TOP_DOWN"
       Case &H1000000: GetTypeAsString = "IMAGE"
       Case Else: If x <> 0 Then GetTypeAsString = "0x" & Hex(MemType)
    End Select
    
End Function


Private Function GetProtectAsString(x As Long, verbose As Boolean) As String

    Select Case x
       Case &H1:   GetProtectAsString = "NOACCESS"
       Case &H2:   GetProtectAsString = IIf(Not verbose, "R", "READONLY")
       Case &H4:   GetProtectAsString = IIf(Not verbose, "RW", "READWRITE")
       Case &H8:   GetProtectAsString = "WRITECOPY"
       Case &H10:  GetProtectAsString = IIf(Not verbose, "E", "EXECUTE")
       Case &H20:  GetProtectAsString = IIf(Not verbose, "RE", "EXECUTE_READ")
       Case &H40:  GetProtectAsString = IIf(Not verbose, "RWE", "EXECUTE_READWRITE")
       Case &H80:  GetProtectAsString = IIf(Not verbose, "EWC", "EXECUTE_WRITECOPY")
       Case &H100: GetProtectAsString = "GUARD"
       Case &H200: GetProtectAsString = "NOCACHE"
       Case Else:  If x <> 0 Then GetProtectAsString = "0x" & Hex(Protect)
    End Select

End Function

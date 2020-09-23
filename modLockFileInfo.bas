Attribute VB_Name = "modLockFileInfo"
'// This code was originally downloaded from:
'http://blog.csdn.net/chenhui530/archive/2007/12/13/1932917.aspx
'// All credits go to the original author.

'// Unnecessary code has been removed and some code modified.


Option Explicit
Public strProcessName As String

Public appname As String

Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
(ByVal lpFileName As String, ByVal nBufferLength As Long, _
ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Private Enum SYSTEM_INFORMATION_CLASS
    SystemBasicInformation
    SystemProcessorInformation
    SystemPerformanceInformation
    SystemTimeOfDayInformation
    SystemPathInformation
    SystemProcessInformation
    SystemCallCountInformation
    SystemDeviceInformation
    SystemProcessorPerformanceInformation
    SystemFlagsInformation
    SystemCallTimeInformation
    SystemModuleInformation
    SystemLocksInformation
    SystemStackTraceInformation
    SystemPagedPoolInformation
    SystemNonPagedPoolInformation
    SystemHandleInformation
    SystemObjectInformation
    SystemPageFileInformation
    SystemVdmInstemulInformation
    SystemVdmBopInformation
    SystemFileCacheInformation
    SystemPoolTagInformation
    SystemInterruptInformation
    SystemDpcBehaviorInformation
    SystemFullMemoryInformation
    SystemLoadGdiDriverInformation
    SystemUnloadGdiDriverInformation
    SystemTimeAdjustmentInformation
    SystemSummaryMemoryInformation
    SystemMirrorMemoryInformation
    SystemPerformanceTraceInformation
    SystemObsolete0
    SystemExceptionInformation
    SystemCrashDumpStateInformation
    SystemKernelDebuggerInformation
    SystemContextSwitchInformation
    SystemRegistryQuotaInformation
    SystemExtendServiceTableInformation
    SystemPrioritySeperation
    SystemVerifierAddDriverInformation
    SystemVerifierRemoveDriverInformation
    SystemProcessorIdleInformation
    SystemLegacyDriverInformation
    SystemCurrentTimeZoneInformation
    SystemLookasideInformation
    SystemTimeSlipNotification
    SystemSessionCreate
    SystemSessionDetach
    SystemSessionInformation
    SystemRangeStartInformation
    SystemVerifierInformation
    SystemVerifierThunkExtend
    SystemSessionProcessInformation
    SystemLoadGdiDriverInSystemSpace
    SystemNumaProcessorMap
    SystemPrefetcherInformation
    SystemExtendedProcessInformation
    SystemRecommendedSharedDataAlignment
    SystemComPlusPackage
    SystemNumaAvailableMemory
    SystemProcessorPowerInformation
    SystemEmulationBasicInformation
    SystemEmulationProcessorInformation
    SystemExtendedHandleInformation
    SystemLostDelayedWriteInformation
    SystemBigPoolInformation
    SystemSessionPoolTagInformation
    SystemSessionMappedViewInformation
    SystemHotpatchInformation
    SystemObjectSecurityMode
    SystemWatchdogTimerHandler
    SystemWatchdogTimerInformation
    SystemLogicalProcessorInformation
    SystemWow64SharedInformation
    SystemRegisterFirmwareTableInformationHandler
    SystemFirmwareTableInformation
    SystemModuleInformationEx
    SystemVerifierTriageInformation
    SystemSuperfetchInformation
    SystemMemoryListInformation
    SystemFileCacheInformationEx
    MaxSystemInfoClass  '// MaxSystemInfoClass should always be the last enum
End Enum

Private Type SYSTEM_HANDLE
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    pObject As Long
    GrantedAccess As Long
End Type

Private Type SYSTEM_HANDLE_INFORMATION
    uCount As Long
    aSH() As SYSTEM_HANDLE
End Type

Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Private Type CLIENT_ID
    UniqueProcess As Long
    UniqueThread  As Long
End Type

Private Enum OBJECT_INFORMATION_CLASS
    ObjectBasicInformation = 0
    ObjectNameInformation
    ObjectTypeInformation
    ObjectAllTypesInformation
    ObjectHandleInformation
End Enum

Private Type UNICODE_STRING
    uLength As Integer
    uMaximumLength As Integer
    pBuffer(3) As Byte
End Type

Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2
Private Const DUPLICATE_SAME_ATTRIBUTES = &H4
Private Const PROCESS_DUP_HANDLE As Long = (&H40)

Private Const STATUS_INFO_LEN_MISMATCH = &HC0000004
Private Const HEAP_ZERO_MEMORY = &H8

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
                                ByVal pSystemInformation As Long, _
                                ByVal SystemInformationLength As Long, _
                                ByRef ReturnLength As Long) As Long
                                
Private Declare Function NtQueryObject Lib "NTDLL.DLL" (ByVal ObjectHandle As Long, _
                                ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, _
                                ByVal ObjectInformation As Long, ByVal ObjectInformationLength As Long, _
                                ReturnLength As Long) As Long
                                
Private Declare Function NtDuplicateObject Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, _
                                ByVal SourceHandle As Long, _
                                ByVal TargetProcessHandle As Long, _
                                ByRef TargetHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByVal HandleAttributes As Long, _
                                ByVal Options As Long) As Long
                                
Private Declare Function NtOpenProcess Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientID As CLIENT_ID) As Long
                                
Private Declare Function NtClose Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long

Public strFile As String

Private Function NT_SUCCESS(ByVal nStatus As Long) As Boolean
    NT_SUCCESS = (nStatus >= 0)
End Function

Public Function GetFileFullPath(ByVal hFile As Long) As String
    Dim hHeap As Long, dwSize As Long, objName As UNICODE_STRING, pName As Long
    Dim ntStatus As Long, i As Long, strDrives As String, strArray() As String
    Dim dwDriversSize As Long, strDrive As String, strTmp As String, strTemp As String
    On Error GoTo ErrHandle
    hHeap = GetProcessHeap
    pName = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, &H1000)
    ntStatus = NtQueryObject(hFile, ObjectNameInformation, pName, &H1000, dwSize)
    If (NT_SUCCESS(ntStatus)) Then
        i = 1
        Do While (ntStatus = STATUS_INFO_LEN_MISMATCH)
            pName = HeapReAlloc(hHeap, HEAP_ZERO_MEMORY, pName, &H1000 * i)
            ntStatus = NtQueryObject(hFile, ObjectNameInformation, pName, &H1000, ByVal 0)
            i = i + 1
        Loop
    End If
    HeapFree hHeap, 0, pName
    strTemp = String(512, Chr(0))
    lstrcpyW strTemp, pName + Len(objName)
    strTemp = StrConv(strTemp, vbFromUnicode)
    strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    strDrives = String(512, Chr(9))
    dwDriversSize = GetLogicalDriveStrings(512, strDrives)
    If dwDriversSize Then
        strArray = Split(strDrives, Chr(0))
        For i = 0 To UBound(strArray)
            If strArray(i) <> "" Then
                strDrive = Left(strArray(i), 2)
                strTmp = String(260, Chr(0))
                Call QueryDosDevice(strDrive, strTmp, 256)
                strTmp = Left(strTmp, InStr(strTmp, Chr(0)) - 1)
                If InStr(LCase(strTemp), LCase(strTmp)) = 1 Then
                    GetFileFullPath = strDrive & Mid(strTemp, Len(strTmp) + 1, Len(strTemp) - Len(strTmp))
                    Exit Function
                End If
            End If
        Next
    End If
ErrHandle:
End Function

Public Function UnLockFile(ByVal strFileName As String) As Boolean
    Dim ntStatus As Long
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim lngHandles As Long
    Dim i As Long
    Dim objInfo As SYSTEM_HANDLE_INFORMATION
    Dim lngType As Long
    Dim hProcessToDup As Long, hFileHandle As Long
    Dim hFile As Long, blnIsOk As Boolean
    Dim strSubPath As String, strTmp As String
    strSubPath = Mid(strFileName, 3, Len(strFileName) - 2)
    hFile = CreateFile("NUL", &H80000000, 0, ByVal 0&, 3, 0, 0)
    If hFile = -1 Then
        UnLockFile = False
        Exit Function
    End If
    objOa.Length = Len(objOa)
    ntStatus = 0
    Dim bytBuf() As Byte
    Dim nSize As Long
    nSize = 1
    Do
        ReDim bytBuf(nSize)
        ntStatus = NtQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), nSize, 0&)
        If (Not NT_SUCCESS(ntStatus)) Then
            If (ntStatus <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Function
            End If
        Else
            Exit Do
        End If
        nSize = nSize * 2
        ReDim bytBuf(nSize)
    Loop
    lngHandles = 0
    CopyMemory objInfo.uCount, bytBuf(0), 4
    lngHandles = objInfo.uCount
    ReDim objInfo.aSH(lngHandles - 1)
    Call CopyMemory(objInfo.aSH(0), bytBuf(4), Len(objInfo.aSH(0)) * lngHandles)
    For i = 0 To lngHandles - 1
        If objInfo.aSH(i).HandleValue = hFile And objInfo.aSH(i).UniqueProcessId = GetCurrentProcessId Then
            lngType = objInfo.aSH(i).ObjectTypeIndex
            Exit For
        End If
    Next
    NtClose hFile
    'blnIsOk = True
    For i = 0 To lngHandles - 1
        If objInfo.aSH(i).ObjectTypeIndex = lngType Then
            objCid.UniqueProcess = objInfo.aSH(i).UniqueProcessId
            ntStatus = NtOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, objOa, objCid)
            If hProcessToDup <> 0 Then
                ntStatus = NtDuplicateObject(hProcessToDup, objInfo.aSH(i).HandleValue, GetCurrentProcess, hFileHandle, 0, 0, DUPLICATE_SAME_ATTRIBUTES)
                If (NT_SUCCESS(ntStatus)) Then
                    ntStatus = MyGetFileType(hFileHandle)
                    If ntStatus Then
                        strTmp = GetFileFullPath(hFileHandle)
                    Else
                        strTmp = ""
                    End If
                    NtClose hFileHandle
                    If InStr(LCase(strTmp), LCase(strFileName)) Then
                        If Not CloseRemoteHandle(objInfo.aSH(i).UniqueProcessId, objInfo.aSH(i).HandleValue) Then
                            blnIsOk = False
                        Else
                        blnIsOk = True
                            '// Add the code to the Listview
                            'lvw.ListItems.Add , , GetProcessPath(GetProcessCommandLine(objInfo.aSH(i).UniqueProcessId))
                            appname = GetProcessPath(GetProcessCommandLine(objInfo.aSH(i).UniqueProcessId))
                            'lvw.ListItems(lvw.ListItems.Count).SubItems(1) = strFile
                            'lvw.ListItems(lvw.ListItems.Count).SubItems(2) = objInfo.aSH(i).UniqueProcessId
                            'lvw.ListItems(lvw.ListItems.Count).SubItems(3) = objInfo.aSH(i).HandleValue
                        End If
                    End If
                End If
            End If
        End If
    Next
    UnLockFile = blnIsOk
End Function

Private Function GetProcessPath(strData As String) As String
    Dim x As Long
    x = InStr(strData, """ ")
    If x Then
        GetProcessPath = Mid$(strData, 2, x - 2)
    End If
End Function

Private Function GetProcessCommandLine(ByVal dwProcessId As Long) As String
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim ntStatus As Long, hKernel As Long, strName As String
    Dim hProcess As Long, dwAddr As Long, dwRead As Long
    objOa.Length = Len(objOa)
    objCid.UniqueProcess = dwProcessId
    ntStatus = NtOpenProcess(hProcess, &H10, objOa, objCid)
    If hProcess = 0 Then
        GetProcessCommandLine = ""
        Exit Function
    End If
    hKernel = GetModuleHandle("kernel32")
    dwAddr = GetProcAddress(hKernel, "GetCommandLineA")
    CopyMemory dwAddr, ByVal dwAddr + 1, 4
    If ReadProcessMemory(hProcess, ByVal dwAddr, dwAddr, 4, dwRead) Then
        strName = String(260, Chr(0))
        If ReadProcessMemory(hProcess, ByVal dwAddr, ByVal strName, 260, dwRead) Then
            strName = Left(strName, InStr(strName, Chr(0)) - 1)
            NtClose hProcess
            GetProcessCommandLine = strName
            Exit Function
        End If
    End If
    NtClose hProcess
End Function

Public Function CloseRemoteHandle(ByVal dwProcessId, ByVal hHandle As Long) As Boolean
    Dim hMyProcess  As Long, hRemProcess As Long, blnResult As Long, hMyHandle As Long
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim ntStatus As Long, hProcess As Long
    objCid.UniqueProcess = dwProcessId
    objOa.Length = Len(objOa)
    hMyProcess = GetCurrentProcess()
    ntStatus = NtOpenProcess(hRemProcess, PROCESS_DUP_HANDLE, objOa, objCid)
    If hRemProcess Then
        ntStatus = NtDuplicateObject(hRemProcess, hHandle, GetCurrentProcess, hMyHandle, 0, 0, DUPLICATE_CLOSE_SOURCE Or DUPLICATE_SAME_ACCESS)
        If (NT_SUCCESS(ntStatus)) Then
            blnResult = NtClose(hMyHandle)
            If blnResult >= 0 Then
                strProcessName = GetProcessCommandLine(dwProcessId)
                If InStr(LCase(strProcessName), "explorer.exe") = 0 And dwProcessId <> GetCurrentProcessId Then
                    objCid.UniqueProcess = dwProcessId
                    ntStatus = NtOpenProcess(hProcess, 1, objOa, objCid)
                    If hProcess <> 0 Then TerminateProcess hProcess, 0
                End If
            End If
        End If
        Call NtClose(hRemProcess)
    End If
    CloseRemoteHandle = blnResult >= 0
End Function

Private Function MyGetFileType(ByVal hFile As Long) As Long
    Dim hRemProcess As Long, hThread As Long, lngResult As Long, pfnThreadRtn As Long, hKernel As Long
    Dim dwEax As Long, dwTimeOut As Long
    hRemProcess = GetCurrentProcess
    hKernel = GetModuleHandle("kernel32")
    If hKernel = 0 Then
        MyGetFileType = 0
        Exit Function
    End If
    pfnThreadRtn = GetProcAddress(hKernel, "GetFileType")
    If pfnThreadRtn = 0 Then
        FreeLibrary hKernel
        MyGetFileType = 0
        Exit Function
    End If
    hThread = CreateRemoteThread(hRemProcess, ByVal 0&, 0&, ByVal pfnThreadRtn, ByVal hFile, 0, ByVal 0&)
    dwEax = WaitForSingleObject(hThread, 100)
    If dwEax = &H102 Then
        Call GetExitCodeThread(hThread, dwTimeOut)
        Call TerminateThread(hThread, dwTimeOut)
        NtClose hThread
        MyGetFileType = 0
        Exit Function
    End If
    If hThread = 0 Then
        FreeLibrary hKernel
        MyGetFileType = False
        Exit Function
    End If
    GetExitCodeThread hThread, lngResult
    MyGetFileType = lngResult
    NtClose hThread
    NtClose hRemProcess
    FreeLibrary hKernel
End Function

Public Function showappname() As String
'Dim parts As String
'Static buf As String * 261, dummy As String * 20
'parts = Split(strProcessName, " ")
'parts = GetFullPathName(strProcessName, 260, buf, dummy)
'parts = GetProcessPath(GetProcessCommandLine(objInfo.aSH(i).UniqueProcessId))
showappname = appname 'strProcessName
End Function


'This function will get the original FileTitle
Public Function GetFileTitle(ByVal sFilename As String) As String
Dim lPos As Long
    'Returns the position of the last occurrence of one string within another
    lPos = InStrRev(sFilename, "\")
    If lPos > 0 Then
        'If lPos is < then the number of chars in sFilename
        If lPos < Len(sFilename) Then
            'Then trim the Path from the FileTitle
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            'If not then set the Function = ""
            GetFileTitle = ""
        End If
      Else
        GetFileTitle = sFilename
    End If
End Function

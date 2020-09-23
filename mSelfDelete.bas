Attribute VB_Name = "mSelfDelete"

'
' mSelfDelete - given sufficient permissions, deletes the exe file of the process that calls this routine
'
'
'Assembly of the injected code:
'        call        _call                   ; call the next instruction
'_call:  pop         ebp                     ; pop the return address into ebp
'        mov         eax,dword ptr [ebp+23h] ; get hParent parameter into eax
'        push        eax                     ; push the hParent parameter
'        push        0FFFFFFFFh              ; push milliseconds parameter, -1 (INFINITE)
'        push        eax                     ; push the hParent parameter
'        call        dword ptr [ebp+27h]     ; call WaitForSingleObject
'        call        dword ptr [ebp+2Bh]     ; call CloseHandle
'_retry: lea         eax,[ebp+3Bh]           ; get the address of szFileName in eax
'        push        eax                     ; push the szFileName parameter
'        call        dword ptr [ebp+2Fh]     ; call DeleteFileA
'        test        eax,eax                 ; check the return value
'        jne         _quit                   ; success
'        push        7Fh                     ; push the milliseconds parameter (127)
'        call        dword ptr [ebp+33h]     ; call Sleep
'        jmp         _retry                  ; try again
'_quit:  call        dword ptr [ebp+37h]     ; call ExitProcess. Note that ExitProcess doesn't return
'

Option Explicit

'Thread context
'Note that I've only declared the members needed because it would be a complete
'nightmare to accurately convert >600 members from 'C' to VB
Private Type tCONTEXT
    ContextFlags            As Long
    space1(1 To 140)        As Byte
    SegFs                   As Long
    space2(1 To 568)        As Byte
End Type

'Local Descriptor Table bytes
Private Type tLDT_BYTES
    BaseMid                 As Byte
    Flags1                  As Byte
    Flags2                  As Byte
    BaseHi                  As Byte
End Type

'Local Descriptor Table
Private Type tLDT_ENTRY
    LimitLow                As Integer
    BaseLow                 As Integer
    HighWord                As tLDT_BYTES
End Type

'Payload code/data
Private Type tPAYLOAD
    Code1                   As Currency     '8 bytes of machine code
    Code2                   As Currency     '8 bytes of machine code
    Code3                   As Currency     '8 bytes of machine code
    Code4                   As Currency     '8 bytes of machine code
    Code5                   As Currency     '8 bytes of machine code
    hParent                 As Long         'handle to the parent process (us)
    fnWaitForSingleObject   As Long         'pointer to WaitForSingleObject
    fnCloseHandle           As Long         'pointer to CloseHandle
    fnDeleteFile            As Long         'pointer to DeleteFileA
    fnSleep                 As Long         'pointer to Sleep
    fnExitProcess           As Long         'pointer to ExitProcess
    szFileName(1 To 260)    As Byte         'filename to delete
End Type

'CreateProcess startup info
Private Type STARTUPINFO
    cb                      As Long
    lpReserved              As Long
    lpDesktop               As Long
    lpTitle                 As Long
    dwX                     As Long
    dwY                     As Long
    dwXSize                 As Long
    dwYSize                 As Long
    dwXCountChars           As Long
    dwYCountChars           As Long
    dwFillAttribute         As Long
    dwFlags                 As Long
    wShowWindow             As Integer
    cbReserved2             As Integer
    lpReserved2             As Byte
    hStdInput               As Long
    hStdOutput              As Long
    hStdError               As Long
End Type

'CreateProcess Process info
Private Type PROCESS_INFORMATION
    hProcess                As Long
    hThread                 As Long
    dwProcessId             As Long
    dwThreadId              As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourcehProcess As Long, ByVal hSourceHandle As Long, ByVal hTargethProcess As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As tCONTEXT) As Long
Private Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As tLDT_ENTRY) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Long, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long

Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)

Public Sub SelfDelete()
    If InIDE Then
        Stop 'SelfDelete will delete VB6.exe if run from the IDE
    Else
        Const FLAGS As Long = &H44 'CREATE_SUSPENDED Or IDLE_PRIORITY_CLASS
        Dim si      As STARTUPINFO
        Dim pi      As PROCESS_INFORMATION
        
        si.cb = Len(si)
        
        'Start notepad in a suspended state
        If CreateProcess(0, "notepad.exe", 0, 0, 0, FLAGS, 0, 0, si, pi) Then
            Dim Payload   As tPAYLOAD
            Dim hKernel32 As Long
            Dim nEntry    As Long
            Dim nProt     As Long
            
            'get a handle to kernel32.dll
            hKernel32 = GetModuleHandle("KERNEL32")
            
            With Payload
                'fill in the machine code
                .Code1 = 501120126496119.2168@
                .Code2 = 283445226384721.9235@
                .Code3 = -4947442599386.5729@
                .Code4 = 918115899267069.9349@
                .Code5 = 398737444671975.1679@
    
                'fill in the function pointers
                .fnCloseHandle = GetProcAddress(hKernel32, "CloseHandle")
                .fnDeleteFile = GetProcAddress(hKernel32, "DeleteFileA")
                .fnExitProcess = GetProcAddress(hKernel32, "ExitProcess")
                .fnSleep = GetProcAddress(hKernel32, "Sleep")
                .fnWaitForSingleObject = GetProcAddress(hKernel32, "WaitForSingleObject")
                
                'duplicate a handle to our process
                Call DuplicateHandle(GetCurrentProcess(), GetCurrentProcess(), pi.hProcess, .hParent, 0, 0, 0)
                
                'fill in the path/filename of our exe to delete
                Call GetModuleFileNameA(0, VarPtr(.szFileName(1)), 260)
                
                'get notepad's entry point address
                nEntry = GetEntryPoint(pi.hProcess, pi.hThread)
                
                'make the entry point code read/write/executable
                Call VirtualProtectEx(pi.hProcess, nEntry, Len(Payload), &H40, nProt)
                
                'write the payload into notepad over the entry point code
                Call WriteProcessMemory(pi.hProcess, nEntry, VarPtr(Payload), Len(Payload), 0)
                
                'resume the suspended notepad process (resumes at the entry point address)
                Call ResumeThread(pi.hThread)
                
                'close handles
                Call CloseHandle(pi.hThread)
                Call CloseHandle(pi.hProcess)
                
                'as soon as this process exits, the code we injected into notepad will delete our exe file
            End With
        End If
    End If
End Sub

'get the entry point address in the notepad proxy process
Private Function GetEntryPoint(ByVal hProcess As Long, ByVal hThread As Long) As Long
    Const FLAGS  As Long = &H10017 'CONTEXT_FULL Or CONTEXT_DEBUG_REGISTERS
    Dim Context  As tCONTEXT
    Dim LdtEntry As tLDT_ENTRY
    Dim nAddr    As Long
    Dim nImgBase As Long
    Dim nRead    As Long
    
    'get the suspended thread's context
    Context.ContextFlags = FLAGS
    Call GetThreadContext(hThread, Context)
    
    'Retrieve a descriptor table entry for the 'f' segment selector
    Call GetThreadSelectorEntry(hThread, Context.SegFs, LdtEntry)
    
    'convert the descriptor table entry to a physical address
    PutMem1 VarPtr(nAddr) + 3, LdtEntry.HighWord.BaseHi
    PutMem1 VarPtr(nAddr) + 2, LdtEntry.HighWord.BaseMid
    PutMem2 VarPtr(nAddr), LdtEntry.BaseLow

    'get the address of the Process Environment Block
    Call ReadProcessMemory(hProcess, nAddr + 48, VarPtr(nAddr), 4, nRead)
    
    'get the image base address
    Call ReadProcessMemory(hProcess, nAddr + 8, VarPtr(nImgBase), 4, nRead)
    
    'get the offset to option header
    Call ReadProcessMemory(hProcess, nImgBase + 60, VarPtr(nAddr), 4, nRead)
    
    'get the entry point address
    Call ReadProcessMemory(hProcess, nImgBase + nAddr + 40, VarPtr(nAddr), 4, nRead)
    
    'actual code entry point address = image base + entry point
    GetEntryPoint = nImgBase + nAddr
End Function

'return whether we're running under the VB IDE
Private Function InIDE(Optional ByRef nValue As Long = 0) As Boolean
    If nValue = 0 Then
        nValue = 1
        Debug.Assert True Or InIDE(nValue)
        InIDE = nValue <> 1
    Else
        nValue = 0
    End If
End Function

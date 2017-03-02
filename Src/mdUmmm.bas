Attribute VB_Name = "mdUmmm"
'=========================================================================
' $Header: /BuildTools/UMMM/Src/mdUmmm.bas 24    16.12.16 12:07 Wqw $
'
'   Unattended Make My Manifest Project
'   Copyright (c) 2009-2016 wqweto@gmail.com
'
' $Log: /BuildTools/UMMM/Src/mdUmmm.bas $
' 
' 24    16.12.16 12:07 Wqw
' REF: impl dll command for std dll redirection
'
' 23    13.09.16 16:57 Wqw
' REF: fix off by one error on trust info level
'
' 21    28.10.15 12:50 Wqw
' REF: deduplicate api progid too
'
' 20    26.06.15 16:22 Wqw
' REF: typelib version in registry is stored in hex
'
' 19    25.06.15 19:55 Wqw
' REF: don't output BOM
'
' 16    27.04.15 22:37 Wqw
' REF: additional controls progid based on tli name
'
' 15    30.01.15 19:46 Wqw
' REF: impl win81 per-monitor dpi awareness
'
' 14    14.11.14 20:02 Wqw
' REF: win10 support
'
' 13    14.11.14 19:48 Wqw
' REF: impl var arg for supported oses
'
' 12    9.12.11 18:00 Wqw
' REF: dump only dispatch kind interfaces
'
' 11    29.11.11 16:21 Wqw
' REF: fixed dispatch vs dual interface proxy/stub clsid
'
' 10    22.02.11 15:49 Wqw
' REF: LookupArray uses like operator
'
' 8     17.02.11 11:14 Wqw
' REF: retval of bool functions
'
' 7     16.02.10 17:40 Wqw
' REF: in dump classes if no typelib found did not close file tag
'
' 6     16.02.10 15:52 Wqw
' REF: Main calls Ummm
'
' 5     16.02.10 13:56 Wqw
' REF: console application, error handling, impl .NET COM assemblies
' dependency
'
' 1     29.9.09 18:35 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const CP_UTF8                   As Long = 65001
Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const ERROR_SUCCESS             As Long = 0
Private Const SAM_READ                  As Long = &H20019
Private Const REG_SZ                    As Long = 1
Private Const REG_EXPAND_SZ             As Long = 2
Private Const REG_DWORD                 As Long = 4
Private Const STD_OUTPUT_HANDLE         As Long = -11&

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (pCLSID As Any, lpszProgID As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_OLEMISC       As String = "recomposeonresize|onlyiconic|insertnotreplace|static|cantlinkinside|canlinkbyole1|islinkobject|insideout|activatewhenvisible|renderingisdeviceindependent|invisibleatruntime|alwaysrun|actslikebutton|actslikelabel|nouiactivate|alignable|simpleframe|setclientsitefirst|imemode|ignoreativatewhenvisible|wantstomenumerge|supportsmultilevelundo"
Private Const STR_LIBFLAG       As String = "restricted|control|hidden|hasdiskimage"
Private Const STR_ATTRIB_MISCSTATUS As String = "miscStatus|miscStatusContent|miscStatusThumbnail|miscStatusIcon|miscStatusDocprint"
Private Const STR_UTF_BOM       As String = "﻿"
Private Const STR_PSOAINTERFACE As String = "{00020424-0000-0000-C000-000000000046}"
Private Const STR_PSDISPATCH    As String = "{00020420-0000-0000-C000-000000000046}"

Private m_oFSO              As Object
Private m_cClasses          As Collection
Private m_cInterfaces       As Collection
Private m_sBasePath         As String

Private Enum DVASPECT
    DVASPECT_CONTENT = 1
    DVASPECT_THUMBNAIL = 2
    DVASPECT_ICON = 4
    DVASPECT_DOCPRINT = 8
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    ConsolePrint "critical: " & Error & " (" & sFunc & ")" & vbCrLf
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Sub Main()
    Ummm Command$
End Sub

Public Function Ummm(sParams As String) As Boolean
    Const FUNC_NAME     As String = "Ummm"
    Dim vArgs           As Variant
    
    On Error GoTo EH
    Set m_oFSO = CreateObject("Scripting.FileSystemObject")
    vArgs = pvSplitArgs(sParams)
    If UBound(vArgs) >= 0 Then
        With m_oFSO.OpenTextFile(At(vArgs, 1, vArgs(0) & ".manifest"), 2, True, 0)
            .Write pvToUtf8(pvProcess(C_Str(vArgs(0))))
        End With
        '--- success
        Ummm = True
    Else
        ConsolePrint "usage: UMMM <ini_file|dll_file> [output_file]" & vbCrLf
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvProcess(sFile As String) As String
    Const FUNC_NAME     As String = "pvProcess"
    Dim cOutput         As Collection
    Dim vElem           As Variant
    Dim vRow            As Variant
    
    On Error GoTo EH
    Set m_cClasses = New Collection
    Set m_cInterfaces = New Collection
    Set cOutput = New Collection
    cOutput.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    cOutput.Add "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"" xmlns:asmv3=""urn:schemas-microsoft-com:asm.v3"">"
    Select Case pvGetCliRva(sFile)
    Case -1
        '--- not an executable image -> ini file
        For Each vElem In Split(m_oFSO.OpenTextFile(sFile, 1, False, 0).ReadAll(), vbCrLf)
            If IsEmpty(vElem) Then
                GoTo QH
            End If
            vRow = pvSplitArgs(CStr(vElem))
            Select Case LCase$(At(vRow, 0))
            Case "identity"
                '--- identity <exe_file> [name] [description]
                '---   exe_file quoted if containing spaces
                pvDumpIdentity pvCanonicalPath(At(vRow, 1)), At(vRow, 2), At(vRow, 3), cOutput
            Case "dependency"
                '--- dependency {<lib_name>|<assembly_dll>} [version] [/update]
                '---   lib_name in { comctl, vc90crt, vc90mfc }
                pvDumpDependency At(vRow, 1), At(vRow, 2), cOutput
            Case "file"
                '--- file <file_name> [interfaces] [classes]
                '---   file_name can be relative to base path from exe_file
                '---   interfaces are | separated, with or w/o leading underscore
                pvDumpClasses At(vRow, 1), At(vRow, 3), cOutput
                pvDumpInterfaces At(vRow, 1), At(vRow, 2), cOutput
            Case "interface"
                '--- interface <file_name> <interfaces>
                pvDumpInterfaces At(vRow, 1), At(vRow, 2), cOutput
            Case "dll"
                '--- dll <file_name> [dll_name]
                pvDumpDll At(vRow, 1), At(vRow, 2), cOutput
            Case "trustinfo"
                '--- trustinfo [level] [uiaccess]
                '---   level in { 1, 2, 3 }
                '---   uiaccess is true/false or 0/1
                pvDumpTrustInfo C_Lng(At(vRow, 1, "1")), C_Bool(At(vRow, 2)), cOutput
            Case "dpiaware"
                '--- dpiaware [on_off] [per_monitor]
                '---   on_off is true/false or 0/1
                pvDumpDpiAware C_Bool(At(vRow, 1)), C_Bool(At(vRow, 2)), cOutput
            Case "supportedos"
                '--- supportedos <os_types>
                '---   os_types are | separated OSes from { vista, win7, win8, win81 } or guids
                pvDumpSupportedOs vRow, cOutput
            End Select
        Next
    Case 0
        '--- native (COM) dll
        pvDumpIdentity pvCanonicalPath(sFile), vbNullString, vbNullString, cOutput
        pvDumpClasses pvCanonicalPath(sFile), vbNullString, cOutput
        pvDumpInterfaces pvCanonicalPath(sFile), "*", cOutput
    Case Else
        '--- .net assembly
        pvDumpDependency pvCanonicalPath(sFile), vbNullString, cOutput
    End Select
QH:
    cOutput.Add "</assembly>"
    For Each vElem In cOutput
        pvProcess = pvProcess & vElem & vbCrLf
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpIdentity(sFile As String, sName As String, sDesc As String, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpIdentity"
    
    On Error GoTo EH
    If LenB(sFile) <> 0 Then
        If LenB(sName) = 0 Then
            sName = pvGetStringFileInfo(pvCanonicalPath(sFile), "CompanyName")
            sName = IIf(LenB(sName) <> 0, sName & ".", vbNullString) & pvGetStringFileInfo(pvCanonicalPath(sFile), "InternalName")
        End If
        If LenB(sDesc) = 0 Then
            sDesc = pvGetStringFileInfo(pvCanonicalPath(sFile), "FileDescription")
        End If
        cOutput.Add Printf("    <assemblyIdentity name=""%1"" processorArchitecture=""X86"" type=""win32"" version=""%2"" />", pvXmlEscape(Zn(sName, "noname")), Zn(m_oFSO.GetFileVersion(sFile), "1.0.0.0"))
    End If
    If LenB(Trim$(sDesc)) <> 0 Then
        cOutput.Add Printf("    <description>%1</description>", pvXmlEscape(sDesc))
    End If
    m_sBasePath = Left$(sFile, InStrRev(sFile, "\") - 1)
    ChDrive Left$(m_sBasePath, 1)
    ChDir m_sBasePath
    '--- success
    pvDumpIdentity = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpDependency(sLibName As String, sVersion As String, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpDependency"
    Dim sOutput         As String
    Dim sTempFile       As String
    Dim sManifest       As String
    Dim lPos            As Long
    Dim oShell          As Object ' WshShell
    
    On Error GoTo EH
    Select Case LCase$(sLibName)
    Case "comctl"
        sOutput = Printf("<assemblyIdentity language=""*"" name=""Microsoft.Windows.Common-Controls"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" type=""win32"" version=""%1"" />", Zn(sVersion, "6.0.0.0"))
    Case "vc90crt"
        sOutput = Printf("<assemblyIdentity type=""win32"" name=""Microsoft.VC90.CRT"" version=""%1"" processorArchitecture=""X86"" publicKeyToken=""1fc8b3b9a1e18e3b"" />", Zn(sVersion, "9.0.21022.8")) ' 9.0.30729.1
    Case "vc90mfc"
        sOutput = Printf("<assemblyIdentity type=""win32"" name=""Microsoft.VC90.MFC"" version=""%1"" processorArchitecture=""X86"" publicKeyToken=""1fc8b3b9a1e18e3b"" />", Zn(sVersion, "9.0.21022.8"))
    Case Else
        If pvFileExists(sLibName) Then
            '--- dump assembly manifest
            Set oShell = CreateObject("WScript.Shell")
            sTempFile = pvGetTempFileName
            oShell.Run Printf("mt.exe -nologo -managedassemblyname:""%1"" -nodependency -out:""%2""", sLibName, sTempFile), 0, True
            '--- read manifest
            On Error Resume Next
            sManifest = m_oFSO.OpenTextFile(sTempFile, 1, False, 0).ReadAll()
            On Error GoTo EH
            If Left$(sManifest, 3) = STR_UTF_BOM Then
                sManifest = pvFromUtf8(Mid$(sManifest, 4))
            Else
                sManifest = pvFromUtf8(sManifest)
            End If
            '--- extract assembly identity
            lPos = InStr(1, sManifest, "<assemblyIdentity", vbTextCompare)
            If lPos > 0 Then
                sOutput = Mid$(sManifest, lPos, InStr(lPos, sManifest, ">") - lPos) & " />"
                '--- update assembly
                If Left$(LCase$(sVersion), 2) = "/u" Or Left$(LCase$(sVersion), 2) = "-u" Then
                    oShell.Run Printf("mt.exe -nologo -manifest ""%2"" -outputresource:""%1"";2", sLibName, sTempFile), 0, True
                End If
            End If
            On Error Resume Next
            Kill sTempFile
            On Error GoTo EH
        End If
    End Select
    If LenB(sOutput) <> 0 Then
        cOutput.Add "    <dependency>"
        cOutput.Add "        <dependentAssembly>"
        cOutput.Add "            " & sOutput
        cOutput.Add "        </dependentAssembly>"
        cOutput.Add "    </dependency>"
        '--- success
        pvDumpDependency = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpClasses(sFile As String, sClasses As String, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpClasses"
    Const STR_MISCSTATUS As String = " miscStatusContent=""recomposeonresize,cantlinkinside,insideout,activatewhenvisible,setclientsitefirst"""
    Dim oTLI            As TypeLibInfo
    Dim sLibName        As String
    Dim oClass          As CoClassInfo
    Dim sProgID         As String
    Dim sVerIndProgID   As String
    Dim sApiProgID      As String
    Dim sThreading      As String
    Dim sRegValue       As String
    Dim sMiscStatus     As String
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim bMultiProgID    As Boolean
    
    On Error GoTo EH
    If Not pvFileExists(pvCanonicalPath(sFile)) Then
        ConsolePrint "error: file %1 not found" & vbCrLf, sFile
        GoTo QH
    End If
    cOutput.Add Printf("    <file name=""%1"">", pvPathDifference(m_sBasePath, sFile))
    '--- note: TypeLibInfoFromFile is corrupting registry by partially registering
    '---   typelib if relative filename is used!!!
    On Error Resume Next
    Set oTLI = TypeLibInfoFromFile(pvCanonicalPath(sFile))
    On Error GoTo EH
    If oTLI Is Nothing Then
        ConsolePrint "warning: no type library found in %1" & vbCrLf, sFile
    Else
        vSplit = Split(sClasses, "|")
        With oTLI
            sLibName = pvRegGetValue("TypeLib\" & .Guid & "\" & Hex(.MajorVersion) & "." & Hex(.MinorVersion))
            cOutput.Add Printf("        <typelib tlbid=""%1"" version=""%2"" flags=""%3"" helpdir=""%4"" />", .Guid, .MajorVersion & "." & .MinorVersion, _
                pvGetFlags(.AttributeMask, Split(STR_LIBFLAG, "|")), _
                vbNullString) ' pvPathDifference(m_sBasePath, .HelpFile)
            For Each oClass In .CoClasses
                If LenB(sClasses) > 0 Then
                    If Not pvLookupArray(vSplit, oClass.Name) Then
                        Set oClass = Nothing
                    End If
                End If
                If Not oClass Is Nothing Then
                    With oClass
                    sVerIndProgID = vbNullString
                    sApiProgID = vbNullString
                    If Not pvSearchCollection(m_cClasses, .Guid) Then
                        If LenB(sLibName) <> 0 Then
                            If LenB(pvRegGetValue("CLSID\" & .Guid & "\InprocServer32")) <> 0 Then
                                sVerIndProgID = pvRegGetValue("CLSID\" & .Guid & "\VersionIndependentProgID", , pvRegGetValue("CLSID\" & .Guid & "\ProgID"))
                                '--- Recent COMDLG32.OCX has 2 coclasses w/ same ProgID
                                If pvSearchCollection(m_cClasses, sVerIndProgID) Then
                                    ConsolePrint "warning: ProgID %1 already used for CLSID %2 (%3)" & vbCrLf, sVerIndProgID, m_cClasses(sVerIndProgID), .Guid
                                    sVerIndProgID = vbNullString
                                    sProgID = vbNullString
                                Else
                                    sProgID = pvRegGetValue(sVerIndProgID & "\CurVer", , sVerIndProgID)
                                End If
                                sApiProgID = pvGetProgID(.Guid)
                                If pvSearchCollection(m_cClasses, sApiProgID) Then
                                    ConsolePrint "warning: ProgID %1 already used for CLSID %2 (%3)" & vbCrLf, sApiProgID, m_cClasses(sApiProgID), .Guid
                                    sApiProgID = sProgID
                                End If
                                sThreading = pvRegGetValue("CLSID\" & .Guid & "\InprocServer32", "ThreadingModel")
                                sMiscStatus = vbNullString
                                For lIdx = 0 To Log(DVASPECT_DOCPRINT) / Log(2)
                                    sRegValue = pvRegGetValue("CLSID\" & .Guid & "\MiscStatus" & IIf(lIdx > 0, "\" & lIdx, vbNullString))
                                    If LenB(sRegValue) <> 0 Then
                                        sMiscStatus = sMiscStatus & Printf(" %1=""%2""", Split(STR_ATTRIB_MISCSTATUS, "|")(lIdx), pvGetFlags(C_Lng(sRegValue), Split(STR_OLEMISC, "|")))
                                    End If
                                Next
                                bMultiProgID = LCase$(sVerIndProgID) <> LCase$(sProgID) Or (LCase$(sApiProgID) <> LCase$(sVerIndProgID) And LCase$(sApiProgID) <> LCase$(sProgID))
                                cOutput.Add Printf("        <comClass clsid=""%1"" tlbid=""%2""%3%4%5%6>", _
                                    .Guid, .Parent.Guid, _
                                    IIf(LenB(sProgID) <> 0, " progid=""" & pvXmlEscape(sProgID) & """", vbNullString), _
                                    IIf(LenB(sThreading) <> 0, " threadingModel=""" & sThreading & """", vbNullString), _
                                    sMiscStatus, _
                                    IIf(Not bMultiProgID, " /", vbNullString))
                                If bMultiProgID Then
                                    If LCase$(sVerIndProgID) <> LCase$(sProgID) Then
                                        cOutput.Add Printf("            <progid>%1</progid>", pvXmlEscape(sVerIndProgID))
                                    End If
                                    If LCase$(sApiProgID) <> LCase$(sVerIndProgID) And LCase$(sApiProgID) <> LCase$(sProgID) Then
                                        cOutput.Add Printf("            <progid>%1</progid>", pvXmlEscape(sApiProgID))
                                    End If
                                    cOutput.Add "        </comClass>"
                                End If
                            End If
                        Else
                            If .AttributeMask And (TYPEFLAG_FCANCREATE Or TYPEFLAG_FCONTROL) <> 0 Then
                                sVerIndProgID = .Parent.Name & "." & .Name
                                If pvSearchCollection(m_cClasses, sVerIndProgID) Then
                                    ConsolePrint "warning: ProgID %1 already used for CLSID %2 (%3)" & vbCrLf, sVerIndProgID, m_cClasses(sVerIndProgID), .Guid
                                    sVerIndProgID = vbNullString
                                End If
                                cOutput.Add Printf("        <comClass clsid=""%1"" tlbid=""%2""%3%4%5%6>", _
                                    .Guid, .Parent.Guid, _
                                    IIf(LenB(sVerIndProgID) <> 0, " progid=""" & pvXmlEscape(sVerIndProgID) & """", vbNullString), _
                                    " threadingModel=""Apartment""", _
                                    IIf((.AttributeMask And TYPEFLAG_FCONTROL) <> 0, STR_MISCSTATUS, vbNullString), _
                                    " /")
                            End If
                        End If
                        m_cClasses.Add Array(sVerIndProgID, sFile), .Guid
                        If LenB(sVerIndProgID) <> 0 Then
                            m_cClasses.Add .Guid, sVerIndProgID
                        End If
                        If LenB(sApiProgID) <> 0 And sApiProgID <> sVerIndProgID Then
                            m_cClasses.Add .Guid, sApiProgID
                        End If
                    Else
                        If LenB(sLibName) <> 0 Then
                            If LenB(pvRegGetValue("CLSID\" & .Guid & "\InprocServer32")) <> 0 Then
                                sVerIndProgID = pvRegGetValue("CLSID\" & .Guid & "\VersionIndependentProgID", , pvRegGetValue("CLSID\" & .Guid & "\ProgID"))
                            End If
                        Else
                            If .AttributeMask And (TYPEFLAG_FCANCREATE Or TYPEFLAG_FCONTROL) <> 0 Then
                                sVerIndProgID = .Parent.Name & "." & .Name
                            End If
                        End If
                        If LenB(sVerIndProgID) <> 0 Then
                            ConsolePrint "warning: coclass %1 GUID is duplicate of %2 (%3) in %4" & vbCrLf, sVerIndProgID, m_cClasses(.Guid)(0), .Guid, m_cClasses(.Guid)(1)
                        End If
                    End If
                    End With
                End If
            Next
        End With
    End If
    cOutput.Add "    </file>"
    '--- success
    pvDumpClasses = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpInterfaces(sFile As String, sInterfaces As String, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpInterfaces"
    Dim oTLI            As TypeLibInfo
    Dim oInterface      As InterfaceInfo
    Dim sProgID         As String
    Dim vSplit          As Variant
    Dim vElem           As Variant
    
    On Error GoTo EH
    If LenB(sInterfaces) = 0 Then
        GoTo QH
    End If
    If Not pvFileExists(pvCanonicalPath(sFile)) Then
        ConsolePrint "error: file %1 not found" & vbCrLf, sFile
        GoTo QH
    End If
    '--- note: TypeLibInfoFromFile is corrupting registry by partially registering
    '---   typelib if relative filename is used!!!
    On Error Resume Next
    Set oTLI = TypeLibInfoFromFile(pvCanonicalPath(sFile))
    On Error GoTo EH
    If oTLI Is Nothing Then
        ConsolePrint "warning: no type library found in %1" & vbCrLf, sFile
        GoTo QH
    End If
    vSplit = Split(sInterfaces, "|")
    With oTLI
        For Each oInterface In oTLI.Interfaces
            With oInterface
                If oInterface.TypeKind = TKIND_DISPATCH Then
                    For Each vElem In vSplit
                        If .Name Like vElem _
                                Or Left$(.Name, 1) = "_" And Mid$(.Name, 2) Like vElem _
                                Or Left$(.Name, 2) = "__" And Mid$(.Name, 3) Like vElem Then
                            sProgID = .Parent.Name & "." & .Name
                            If Not pvSearchCollection(m_cInterfaces, .Guid) Then
                                If (.AttributeMask And TYPEFLAG_FDISPATCHABLE) <> 0 Then
                                    cOutput.Add Printf("    <comInterfaceExternalProxyStub name=""%1"" iid=""%2"" tlbid=""%3"" proxyStubClsid32=""%4"" />", pvXmlEscape(.Name), .Guid, .Parent.Guid, IIf((.AttributeMask And TYPEFLAG_FDUAL) <> 0, STR_PSOAINTERFACE, STR_PSDISPATCH))
                                    m_cInterfaces.Add Array(sProgID, sFile), .Guid
                                Else
                                    ConsolePrint "warning: interface %1 is not dispatch-based, no proxy/stub tags generated" & vbCrLf, sProgID
                                End If
                            Else
                                ConsolePrint "warning: interface %1 GUID is duplicate of %2 (%3) in %4" & vbCrLf, sProgID, m_cInterfaces(.Guid)(0), .Guid, m_cInterfaces(.Guid)(1)
                            End If
                            Exit For
                        End If
                    Next
                End If
            End With
        Next
    End With
    '--- success
    pvDumpInterfaces = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpDll(sPath As String, sName As String, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpDll"
        
    On Error GoTo EH
    If LenB(sName) = 0 Then
        sName = Mid$(sPath, InStrRev(sPath, "\") + 1)
    End If
    cOutput.Add Printf("    <file name=""%1"" loadFrom=""%2"" />", sName, sPath)
    '--- success
    pvDumpDll = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpTrustInfo(ByVal lLevel As Long, ByVal bUiAccess As Boolean, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpTrustInfo"
    Dim sLevel          As String
    
    On Error GoTo EH
    sLevel = At(Array("asInvoker", "highestAvailable", "requireAdministrator"), lLevel - 1, "asInvoker")
    cOutput.Add "    <trustInfo xmlns=""urn:schemas-microsoft-com:asm.v3"">"
    cOutput.Add "        <security>"
    cOutput.Add "            <requestedPrivileges>"
    cOutput.Add Printf("                <requestedExecutionLevel level=""%1""%2 />", sLevel, IIf(bUiAccess, " uiAccess=""true""", vbNullString))
    cOutput.Add "            </requestedPrivileges>"
    cOutput.Add "        </security>"
    cOutput.Add "    </trustInfo>"
    '--- success
    pvDumpTrustInfo = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpDpiAware(ByVal bAware As Boolean, ByVal bPerMonitor As Boolean, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpDpiAware"
    
    On Error GoTo EH
    '--- note: win7 check only presense of dpiAware element not its value: if present -> app is DPI-aware
    '---   more info at https://msdn.microsoft.com/en-ca/magazine/dn574798.aspx
    If bAware Then
        cOutput.Add "    <asmv3:application>"
        cOutput.Add "        <asmv3:windowsSettings xmlns=""http://schemas.microsoft.com/SMI/2005/WindowsSettings"">"
        cOutput.Add Printf("            <dpiAware>%1</dpiAware>", bAware & IIf(bPerMonitor, "/PM", vbNullString))
        cOutput.Add "        </asmv3:windowsSettings>"
        cOutput.Add "    </asmv3:application>"
    End If
    '--- success
    pvDumpDpiAware = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvDumpSupportedOs(vRow As Variant, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvDumpSupportedOs"
    Dim lIdx            As Long
    Dim sGuid           As String
    
    On Error GoTo EH
    cOutput.Add "    <compatibility xmlns=""urn:schemas-microsoft-com:compatibility.v1"">"
    cOutput.Add "        <application>"
    For lIdx = 1 To UBound(vRow)
        Select Case LCase$(At(vRow, lIdx))
        Case "vista"
            sGuid = "{e2011457-1546-43c5-a5fe-008deee3d3f0}"
        Case "win7"
            sGuid = "{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"
        Case "win8"
            sGuid = "{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"
        Case "win81"
            sGuid = "{1f676c76-80e1-4239-95bb-83d0f6d0da78}"
        Case "win10"
            sGuid = "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"
        Case Else
            '--- this has to be properly escaped attribute value
            sGuid = At(vRow, lIdx)
        End Select
        If LenB(sGuid) <> 0 Then
            cOutput.Add Printf("            <supportedOS Id=""%1"" />", sGuid)
        End If
    Next
    cOutput.Add "        </application>"
    cOutput.Add "    </compatibility>"
    '--- success
    pvDumpSupportedOs = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetFlags(ByVal lMask As Long, vFlags As Variant) As String
    Const FUNC_NAME     As String = "pvGetFlags"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 0 To UBound(vFlags)
        If LenB(vFlags(lIdx)) <> 0 Then
            If (lMask And 2 ^ lIdx) <> 0 Then
                If LenB(pvGetFlags) <> 0 Then
                    pvGetFlags = pvGetFlags & ","
                End If
                pvGetFlags = pvGetFlags & vFlags(lIdx)
            End If
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvSplitArgs(sText As String) As Variant
    Const FUNC_NAME     As String = "pvSplitArgs"
    Dim oMatches        As Object
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = """([^""]*(?:""""[^""]*)*)""|([^ ]+)"
        Set oMatches = .Execute(sText)
        If oMatches.Count > 0 Then
            ReDim vRetVal(0 To oMatches.Count - 1) As String
            For lIdx = 0 To oMatches.Count - 1
                With oMatches(lIdx)
                    vRetVal(lIdx) = Replace$(.SubMatches(0) & .SubMatches(1), """""", """")
                End With
            Next
        Else
            vRetVal = Split(vbNullString)
        End If
    End With
    pvSplitArgs = vRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvPathDifference(sBase As String, sFolder As String) As String
    Const FUNC_NAME     As String = "pvPathDifference"
    Dim vBase           As Variant
    Dim vFolder         As Variant
    Dim lIdx            As Long
    Dim lJ              As Long
    
    On Error GoTo EH
    If LCase$(Left$(sBase, 2)) <> LCase$(Left$(sFolder, 2)) Then
        pvPathDifference = sFolder
    Else
        vBase = Split(sBase, "\")
        vFolder = Split(sFolder, "\")
        For lIdx = 0 To UBound(vFolder)
            If lIdx <= UBound(vBase) Then
                If LCase$(vBase(lIdx)) <> LCase$(vFolder(lIdx)) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        If lIdx > UBound(vBase) Then
'            pvPathDifference = "."
        Else
            For lJ = lIdx To UBound(vBase)
                pvPathDifference = pvPathDifference & IIf(Len(pvPathDifference) > 0, "\", "") & ".."
            Next
        End If
        For lJ = lIdx To UBound(vFolder)
            pvPathDifference = pvPathDifference & IIf(Len(pvPathDifference) > 0, "\", "") & vFolder(lJ)
        Next
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCanonicalPath(sPath As String) As String
    On Error Resume Next
    pvCanonicalPath = sPath
    pvCanonicalPath = m_oFSO.GetAbsolutePathName(sPath)
    On Error GoTo 0
End Function

Public Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Private Function pvXmlEscape(sText As String) As String
    pvXmlEscape = Replace(Replace(Replace(Replace(Replace(sText, _
            "&", "&amp;"), _
            "<", "&lt;"), _
            ">", "&gt;"), _
            """", "&quot;"), _
            "'", "&apos;")
End Function

Private Function pvToUtf8(sText As String) As String
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), vbNullString, 0, 0, 0)
    If lSize > 0 Then
        pvToUtf8 = String(lSize, 0)
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), pvToUtf8, lSize, 0, 0)
    End If
End Function

Private Function pvFromUtf8(sText As String) As String
    Dim lSize           As Long
    
    lSize = MultiByteToWideChar(CP_UTF8, 0, StrPtr(StrConv(sText, vbFromUnicode)), Len(sText), 0, 0)
    If lSize > 0 Then
        pvFromUtf8 = String(lSize, 0)
        Call MultiByteToWideChar(CP_UTF8, 0, StrPtr(StrConv(sText, vbFromUnicode)), Len(sText), StrPtr(pvFromUtf8), lSize)
    End If
End Function

Private Function pvRegGetValue(sKey As String, Optional sValue As String, Optional sDefault As String) As String
    Const FUNC_NAME     As String = "pvRegGetValue"
    Dim hKey            As Long
    Dim lSize           As Long
    Dim lType           As Long
    Dim sString         As String
    Dim lDWord          As Long
    
    On Error GoTo EH
    pvRegGetValue = sDefault
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey, 0, SAM_READ, hKey) = ERROR_SUCCESS Then
        If RegQueryValueEx(hKey, sValue, 0, lType, ByVal 0&, lSize) = ERROR_SUCCESS Then
            If lType = REG_SZ Or lType = REG_EXPAND_SZ Then
                If lSize > 0 Then
                    sString = String(lSize - 1, 0)
                    If RegQueryValueEx(hKey, sValue, 0, lType, ByVal sString, lSize) = ERROR_SUCCESS Then
                        pvRegGetValue = sString
                    End If
                End If
            ElseIf lType = REG_DWORD Then
                If RegQueryValueEx(hKey, sValue, 0, lType, lDWord, 4) = ERROR_SUCCESS Then
                    pvRegGetValue = lDWord
                End If
            End If
        End If
        Call RegCloseKey(hKey)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvSearchCollection(Col As Object, Index As Variant) As Boolean
    On Error Resume Next
    IsObject Col(Index)
    pvSearchCollection = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error Resume Next
    At = sDefault
    At = C_Str(vData(lIdx))
    On Error GoTo 0
End Function

Public Function C_Lng(ByVal v As Variant) As String
    On Error Resume Next
    C_Lng = CLng(v)
    On Error GoTo 0
End Function

Public Function C_Str(ByVal v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
    On Error GoTo 0
End Function

Public Function C_Bool(ByVal v As Variant) As Boolean
    On Error Resume Next
    C_Bool = CBool(v)
    On Error GoTo 0
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Private Function pvFileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        pvFileExists = True
    End If
End Function

Private Function pvGetTempFileName() As String
    Dim sFile           As String
    
    sFile = String(2000, 0)
    Call GetTempFileName(Environ$("TEMP"), "UMMM", 0, sFile)
    If InStr(sFile, Chr$(0)) > 0 Then
        pvGetTempFileName = Left$(sFile, InStr(sFile, Chr$(0)) - 1)
    Else
        pvGetTempFileName = "C:\UMMM.tmp"
    End If
End Function

Private Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long
    Dim hOut            As Long
    
    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(&H101))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    ConsolePrint = Replace(sText, ChrW$(&H101), "%")
    '--- output
    hOut = GetStdHandle(STD_OUTPUT_HANDLE)
    If hOut = 0 Then
        Debug.Print ConsolePrint
    Else
        ReDim baBuffer(1 To Len(ConsolePrint)) As Byte
        If CharToOemBuff(ConsolePrint, baBuffer(1), UBound(baBuffer)) Then
            Call WriteFile(hOut, baBuffer(1), UBound(baBuffer), dwDummy, ByVal 0&)
        End If
    End If
End Function

Private Function pvLookupArray(vSplit As Variant, sName As String) As Boolean
    Dim vElem           As Variant
    
    For Each vElem In vSplit
        If sName Like vElem Then
            pvLookupArray = True
            Exit Function
        End If
    Next
End Function

Private Function pvGetStringFileInfo(sFile As String, sKey As String) As String
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim lPtr            As Long
    Dim lCharset        As Long
    
    lSize = GetFileVersionInfoSize(sFile, 0)
    ReDim baBuffer(0 To lSize)
    If GetFileVersionInfo(sFile, 0, UBound(baBuffer), baBuffer(0)) = 0 Then
        GoTo QH
    End If
    If VerQueryValue(baBuffer(0), "\VarFileInfo\Translation", lPtr, lSize) = 0 Then
        GoTo QH
    End If
    Call CopyMemory(ByVal VarPtr(lCharset) + 2, ByVal lPtr, 2)
    Call CopyMemory(ByVal VarPtr(lCharset) + 0, ByVal lPtr + 2, 2)
    lSize = 0
    If VerQueryValue(baBuffer(0), "\StringFileInfo\" & Right$(String$(8, "0") & Hex(lCharset), 8) & "\" & sKey, lPtr, lSize) = 0 Then
        GoTo QH
    End If
    If lSize > 0 Then
        pvGetStringFileInfo = String$(lSize - 1, 0)
        Call CopyMemory(ByVal pvGetStringFileInfo, ByVal lPtr, lSize)
    End If
    
    ' Strip out null termination (ASCII zero)
    pvGetStringFileInfo = Replace(pvGetStringFileInfo, Chr$(0), "")
QH:
End Function

Private Function pvPeekFile(sFile As String, ByVal lOffset As Long) As Long
    Dim nFile           As Integer
    
    nFile = FreeFile
    Open sFile For Binary Access Read As nFile
    Seek nFile, lOffset + 1
    Get nFile, , pvPeekFile
    Close nFile
End Function

Private Function pvGetCliRva(sFile As String) As Long
    Dim lPeOffset       As Long
    
    pvGetCliRva = -1
    If (pvPeekFile(sFile, 0) And &HFFFF&) <> Asc("Z") * &H100 + Asc("M") Then
        GoTo QH
    End If
    lPeOffset = pvPeekFile(sFile, &H3C)
    If pvPeekFile(sFile, lPeOffset) <> Asc("E") * &H100 + Asc("P") Then
        GoTo QH
    End If
    pvGetCliRva = pvPeekFile(sFile, lPeOffset + 4 + 20 + 208)
QH:
End Function

Private Function pvGetProgID(sClsID As String) As String
    Dim aCLSID(0 To 3)  As Long
    Dim lPtr            As Long
    
    Call CLSIDFromString(StrPtr(sClsID), aCLSID(0))
    Call ProgIDFromCLSID(aCLSID(0), lPtr)
    If lPtr <> 0 Then
        pvGetProgID = String$(lstrlenW(lPtr), 0)
        Call CopyMemory(ByVal StrPtr(pvGetProgID), ByVal lPtr, LenB(pvGetProgID))
        Call CoTaskMemFree(lPtr)
    End If
End Function

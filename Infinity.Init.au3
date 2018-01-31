#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SystemRound.256.ico
#AutoIt3Wrapper_Outfile_x64=..\System.Init.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=InfinityPE Managment
#AutoIt3Wrapper_Res_Fileversion=1.0.1.85
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=InfinityResearchAndDevelopment 2017
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Au3Stripper=n
#Au3Stripper_Parameters=/debug
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.15.0 (Beta)
 Author:         BiatuAutMiahn[@outlook.com]

 Script Function:
	Initialize & Manage InfinityPE

#ce ----------------------------------------------------------------------------
#Include <Array.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <WinAPISys.au3>
#include <WinAPIFiles.au3>
#include <WinAPIProc.au3>
#include <WinAPIDiag.au3>
#include <WinAPIGdi.au3>
#include <File.au3>
#include <SecurityConstants.au3>
#include <Security.au3>
#include "Includes\Services.au3"
#include "Includes\SetupAPI.au3"

If $Cmdline[0]<>1 Then Exit 1
;_ArrayDisplay($CmdLine)
Global $sSystemDrive=_WinAPI_GetFullPathName(@WindowsDir&"\..")
Global $sSystemLog=$sSystemDrive&"\System.Log"
;Global $hSystemLog=FileOpen($sSystemLog,1)
Global $sNetWorkgroup="INFINITYCOMMS"
Global $sNetComputerName=@ComputerName
Global $sHost=_VCheck()
Global $sBootSrcWim=""
Global $sBootSrcRW=""
Global $sBootSrcType=""
Global $aDevices[1][3]

Global $iUIDWM,$iUIProgressLast,$idUIProgress,$idUIStatus,$hUI
Global Const $DIRECTORY_QUERY = 0x0001
Global Const $DIRECTORY_TRAVERSE = 0x0002
Global Const $sTagUNICODESTRING = "USHORT Length;USHORT MaximumLength;PTR Buffer;"
Global Const $sTagOBJECT_ATTRIBUTES = "ULONG Length;HANDLE RootDirectory;PTR ObjectName;ULONG Attributes;PTR SecurityDescriptor;PTR SecurityQualityOfService"
_CM_Enable_Privileges()

Switch $CmdLine[1]
    Case "Init"
        Run("X:\Programs.64\GetDeps\GetDeps.exe","X:\Programs.64\GetDeps",@SW_SHOW,0x10000)
        _Exec("PECMD.exe DISP W1024 H768 B32 F60")
        _ChangeWallpaper(@SystemDir&"\winpe.jpg",2,0)
        $iTotP=100/15
        _InitUI()
        _Log("Machine Type: "&$sHost)
        $sNetComputerName="InfinityPE-"&StringLower(Hex(Random(0,65535,1),4))
        _StatusUI("Executing wpeinit.exe...",$iTotP*0)
        RunWait(@SystemDir&'\wpeinit.exe')
        ;_Exec("Ux.exe")
;~         _StatusUI("Configuring Workgroup",$iTotP*1)
;~         $iRet=_SetWorkgroup($sNetWorkgroup)
;~         If @error Then
;~             _Log("Failed -> "&_WinAPI_GetLastErrorMessage()&","&@Extended,2)
;~         Else
;~             _Log("Success",2)
;~         EndIf

;~         _StatusUI("Configuring Computer Name",$iTotP*2)
;~         _Log("Configuring Computer Name -> "&$sNetComputerName,1)
;~         $iRet=_SetComputerName($sNetComputerName)
;~         If @error Then
;~             _Log("Failed -> "&_WinAPI_GetLastErrorMessage()&","&@Extended,2)
;~         Else
;~             _Log("Success",2)
;~         EndIf
        ;InitSwap
        ;InitPersist
        _DriverNoSign()
        _Log("Initializing Services",1)
        RegDelete("HKLM\SYSTEM\CurrentControlSet\Control\MiniNT")
        _InitService("eventsystem",$iTotP*3)
        _InitService("eventlog",$iTotP*3)
        _InitService("wecsvc",$iTotP*3)
        ;RegWrite("HKLM\SYSTEM\CurrentControlSet\Control\MiniNT")
        _InitService("swprv",$iTotP*4)
        _InitService("vss",$iTotP*5)
        _StatusUI("Initializng ms_ndisuio network protocol",$iTotP*6)
        _Exec("netcfg.exe -c p -i ms_ndisuio")
        _StatusUI("Initializng ms_sstp network protocol",$iTotP*6)
        _Exec("netcfg.exe -c p -i ms_sstp")
        _StatusUI("Installing ms_nativewifip network service",$iTotP*7)
        _Exec("netcfg.exe -c s -i ms_nativewifip")
        _Exec("netcfg.exe -c p -i ms_vwifi")
        _Exec("netcfg.exe -c s -i ms_vwifi")
        _Exec("drvload.exe X:\Windows\Inf\netvwififlt.inf")
        _Exec('Drvload.exe "'&@SystemDir&'\DriverStore\FileRepository\netvwifibus.inf_amd64_neutral_9d0740f32ce81d24\netvwifibus.inf"')

        _InitService("lmhosts",$iTotP*8)
        _InitService("lanmanserver",$iTotP*9)
        _InitService("sens",$iTotP*10)
        _InitService("profsvc",$iTotP*10)

        _Log("Configuring Explorer",1)
        _StatusUI("Configuring Shell",$iTotP*11)
        _ShellCfg();RunWait(@ComSpec&" /c ShellConfig.cmd")
        _StatusUI("Registering Fonts",$iTotP*12)
        _Exec(@SystemDir&'\AddFonts.exe')

        _StatusUI("Refreshing Program Attributes",$iTotP*13)
        _RefreshFolderAttribs(@ProgramsCommonDir)
        _RefreshFolderAttribs(@ProgramsDir)
        _RefreshFolderAttribs($sSystemDrive&"\Programs")
        _RefreshFolderAttribs($sSystemDrive&"\Programs.32")
        _RefreshFolderAttribs($sSystemDrive&"\Programs.64")
        If $sHost="VirtualBox" Then
            _StatusUI("Initializing VirtualBox Services",$iTotP*15)
            _InitVBox()
        EndIf
        _Log("Initializing Drivers",1)
        RunWait("cmd.exe")
;~         _CM_Scan_Device_Changes("")
        _Log("Configuring Logon",1)
        _StatusUI("Configuring Logon",$iTotP*14)
        _Exec('net user InfinityUser "" /add')
        _Exec('net localgroup Administrators InfinityUser /add')
        _Exec('net localgroup Users InfinityUser /add')
        _Exec("secedit.exe /configure /db " & @WindowsDir & "\security\database\unattend.sdb /cfg " & @WindowsDir & "\security\templates\unattend.inf /log " & @WindowsDir & "\security\logs\unattend.log")
        RegWrite("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName", "REG_SZ", "InfinityUser")
        RegWrite("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon", "REG_SZ", "1")
        RegWrite("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultPassword", "REG_SZ", "")
        RegWrite("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "ForceAutoLogon", "REG_DWORD", "1")
;~         RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce", "ReConn", "REG_SZ", "System.Init.exe Logon")
        RegWrite("HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce", "ReConn", "REG_SZ", "System.Init.exe Logon")

;~         _Exec("Ux.exe")
        ;Get Driverless Devices
;~         Run("cmd.exe")

        ;NET Load

        ;Begin Logon
        _Log("Initializing Logon",1)
        _StatusUI("Initializing Logon",$iTotP*15)
        DllCall("wtsapi32.dll", "int", "WTSDisconnectSession", "ptr", 0, "dword", -1, "ptr", 0)
    Case "Logon"
        _DriverNoSign()
        ;Enable NumLock
        Send("{NUMLOCK on}")
        _Log("Configuring Explorer",1)
        _ShellCfg();RunWait(@ComSpec&" /c ShellConfig.cmd")
        _Exec("psexec.exe -i -s "&@ScriptFullPath&" Shell")
        Local $iTotP=100/4
        _InitUI()
        _StatusUI("Initializing Network Drive",$iTotP*1)
        _Log("Initializing Network Drive",1)
        $hTimer=TimerInit()
        Do
            Sleep(1000)
            If TimerDiff($hTimer)>10000 Then
                ExitLoop
            EndIf
        Until _WinAPI_IsNetworkAlive()
        _Exec("net use T: \\192.168.0.2\reminst\InfinityPE /USER:INFINITYCOMMS\Biatu *")
        While ProcessExists("Explorer.exe")
            Sleep(1000)
            ProcessClose("Explorer.exe")
        WEnd
        _StatusUI("Starting VirtualBox Tray...",$iTotP*2)
        Run("X:\Windows\System32\VBoxTray.exe",@SystemDir)
        _StatusUI("Fixing Themes...",$iTotP*3)
        _Exec("net stop UxSms")
        _Exec("net start UxSms")
        _WinAPI_DwmEnableComposition(1)
        _Exec("Ux.exe")
        _StatusUI("Starting Shell...",$iTotP*4)
        Run("Explorer.exe")
    Case "Shell"
        _DriverNoSign()
        Local $iTotP=100/8
        _InitUI()
        ;MsgBox(64,"","")
        ;Run("X:\Programs.64\GetDeps\GetDeps.exe","X:\Programs.64\GetDeps")
        ;Sleep(5000)
        While ProcessExists("runonce.exe")
            Sleep(1000)
            ProcessClose("runonce.exe")
        WEnd
        _StatusUI("Configuring Shell",$iTotP*1)
        _ShellCfg()
        Run("Explorer.exe")
        Sleep(4000)
        While ProcessExists("Explorer.exe")

            ProcessClose("Explorer.exe")
        WEnd
        RegWrite("HKCU\Software\Policies\Microsoft\Windows NT\Driver Signing","BehaviorOnFailedVerify","reg_dword",0)
        RegWrite("HKU\.DEFAULT\Software\Policies\Microsoft\Windows NT\Driver Signing","BehaviorOnFailedVerify","reg_dword",0)
        RegWrite("HKLM\Software\Policies\Microsoft\Windows NT\Driver Signing","BehaviorOnFailedVerify","reg_dword",0)
        ;GFX Load
        _StatusUI("Initializing Graphics",$iTotP*2)
		If Not _WinAPI_DwmIsCompositionEnabled() Then _Exec("net stop UxSms")
;~ 		Run("explorer.exe")
;~ 		ProcessWait("explorer.exe")
        If Not _WinAPI_DwmIsCompositionEnabled() Then
            _InitGFX()
            _Exec("net start UxSms")
            _WinAPI_DwmEnableComposition(1)
        EndIf
        ;NetworkPersist
        _DriverNoSign()
        _StatusUI("Initializing Network Driver",$iTotP*3)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\nete1g3e.inf_amd64_neutral_7f08406e40c6ede2\nete1g3e.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\netwxw00.inf_amd64_neutral_ed186b01f1d63c77\netwxw00.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\netwsw01.inf_amd64_neutral_6b0c7e460b50118c\netwsw01.inf",1)

        ;Aud Load
        _StatusUI("Initializing Audio Driver",$iTotP*4)
;~         _Exec("drvload.exe "&@WindowsDir&"\inf\vmaudio.inf",15)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\wdmaudio.inf_amd64_neutral_423894ded0ba8fdf\wdmaudio.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\wdma_usb.inf_amd64_neutral_7bb325bca8ea1218\wdma_usb.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\hdaudbus.inf_amd64_neutral_4b99fffee061ff26\hdaudbuf.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\hdaudss.inf_amd64_neutral_330a593eb888237c\hdaudss.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\hdaudio.inf_amd64_neutral_ce7bc199c85ae0a0\hdaudio.inf",1)

        ;_Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\machine.inf_amd64_neutral_a2f120466549d68b\machine.inf",1)
        _FindDeviceRecur(_CM_Locate_DevNode(""))
        ;_ArrayDisplay($aDevices)
        _FileWriteFromArray($sSystemDrive&"\Drivers.ini",$aDevices)
        _StatusUI("Scanning for Other Devices...",$iTotP*5)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\hal.inf_amd64_neutral_a768502c7cb64799\hal.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\ramdisk.inf_amd64_neutral_798b5d4dd3f22a07\ramdisk.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\battery.inf_amd64_neutral_0ff3dd8f4047f37b\battery.inf",1)
        ;MsgBox(64,"","")
        _Log("Initializing Applications",1)
        _StatusUI("Registering MMC",$iTotP*6)
        _Log("MMC",2)
        _Exec("regsvr32 /s X:\Windows\System32\mmcndmgr.dll")
        _StatusUI("Registering MSI Installer",$iTotP*7)
        _Log("MSI",2)
        _Exec("regsvr32 /s X:\Windows\System32\msi.dll")
        _Exec("regsvr32 /s X:\Windows\SysWow64\msi.dll")
        _StatusUI("Starting MiniTool Partition Wizard Services",$iTotP*8)
        _Log("MiniTool Partition Wizard",2)
        _InitService("pwdrvio")
        _InitService("pwdspio")
        _StatusUI("Notepad++",$iTotP*9)
        _Log("Notepad++",2)
        _Exec("Regsvr32 /s X:\Programs.64\Notepad++\NppShell_06.dll")
        _StatusUI("7-zip",$iTotP*10)
        _Log("7-Zip",2)
        _Exec("Regsvr32 /s X:\Programs.64\7-Zip\7-zip.dll")
        _Exec("Regsvr32 /s X:\Programs.64\7-Zip\7-zip32.dll")
        _StatusUI("PdaNet",$iTotP*11)
        _Log("PdaNet",2)
        _Exec("Regsvr32 /s X:\Programs.32\PdaNet\brw.dll")
        _StatusUI("Firefox",$iTotP*12)
        _Log("Firefox",2)
        _Exec("Regsvr32 /s X:\Programs.64\Firefox\AccessibleHandler.dll")
        _StatusUI("Bulk Rename Utility",$iTotP*13)
        _Log("Bulk Rename Utility",2)
        _Exec("Regsvr32 /s X:\Programs.64\BRU\BRUhere64.dll")
        _StatusUI("HTML Help",$iTotP*14)
        _Log("HTML Help",2)
        _Exec("Regsvr32 /s X:\Windows\System32\hhctrl.ocx")
        _Exec("Regsvr32 /s X:\Windows\System32\itss.dll")
        _Exec("Regsvr32 /s X:\Windows\System32\itircl.dll")
        _Log("PowerISO",2)
        _Exec("Regsvr32 /s X:\Programs.64\PowerISO\PWRISOSH.DLL")
        Sleep(1000)
;
;~         _Log("Clearing Policies..."&RegDelete("HKLM\Software\Policies"))
;~         RegWrite("HKLM\Software\Policies")
;~         RegWrite("HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce", "ReConn", "REG_SZ", "System.Init.exe Shell2")

    ;Case "Apps"
    ;Scan for Programs[.32|.64]\\.*\\.InfinityPE\Program.ini
    ;
    ;Case "Update"
    ;   Get Reg
    ;   Get Arc Paths
    ;       If Fail, Attempt manually
    ;   PhoneHome, Check for new packages
    ;Case "NetDev"
    ;Case "InitDrivers"
        ;Scan For Hardware Devices
        ;
    ;Case "InitPersist"
    ;Case "InitNetPersist"
    Case Else
        Exit 1
EndSwitch

Func _DriverNoSign()
;~     RegWrite("HKCU\Software\Policies\Microsoft\Windows NT\Driver Signing","BehaviorOnFailedVerify","reg_dword",0)
;~     RegWrite("HKCU\Software\Microsoft\Driver Signing","Policy","reg_binary",0x00)
;~     RegWrite("HKCU\Software\Microsoft\Non-Driver Signing","Policy","reg_binary",0x00)
;~     RegWrite("HKU\.DEFAULT\Software\Policies\Microsoft\Windows NT\Driver Signing","BehaviorOnFailedVerify","reg_dword",0)
;~     RegWrite("HKLM\Software\Policies\Microsoft\Windows NT\Driver Signing","BehaviorOnFailedVerify","reg_dword",0)
;~     RegWrite("HKLM\Software\Microsoft\Driver Signing","Policy","reg_binary",0x00)
;~     RegWrite("HKLM\Software\Microsoft\Non-Driver Signing","Policy","reg_binary",0x00)
EndFunc

Func _FindDeviceRecur($hDevInst)
        Local $aChild, $sDescr, $sClass,$iMax
        $aChild=_CM_Enumerate_Children($hDevInst)
        For $i=1 To $aChild[0]
            $hDevInst=_CM_Locate_DevNode($aChild[$i])
            $sClass=_CM_Get_DevNode_Registry_Property($hDevInst,$CM_DRP_CLASS)
            _CM_Get_DevNode_Status_Ex($hDevInst,0)
            $iProb=@extended
            If ($sClass="" And $iProb<>0) Then
                $iMax=UBound($aDevices,1)
                ReDim $aDevices[$iMax+1][3]
                $sDescr=_CM_Get_Device_Display_Name_Ex($hDevInst)
                $aDevices[0][0]=$iMax
                $aDevices[$iMax][0]=$sDescr
                $aDevices[$iMax][1]=$aChild[$i]
                $aDevices[$iMax][2]=$iProb
            EndIf
            _FindDeviceRecur($hDevInst)
        Next
        ;Return $aDevices
EndFunc

Func _ShellCfg()
    FileDelete(@DesktopDir&"\Desktop.ini")
    FileDelete(@DesktopCommonDir&"\Desktop.ini")
    RegWrite("HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced","IconsOnly","REG_DWORD",1)
    RegWrite("HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams","Settings","REG_BINARY",Binary("0x080000000600000001000000e525f165e17b1048ba9dd271c8432ce304000000be01000043000000"))
    RegWrite("HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Defaults","{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}","REG_BINARY",Binary("0x1c000000000000000000000000000000000000000000000000000000f1f1f1f114000000000000000000000000000000d0020000cc0200003153505305d5cdd59c2e1b10939708002b2cf9ae830000002200000000470072006f0075007000420079004b00650079003a0046004d005400490044000000080000004e0000007b00300030003000300030003000300030002d0030003000300030002d0030003000300030002d0030003000300030002d003000300030003000300030003000300030003000300030007d0000000000330000002200000000470072006f00750070004200790044006900720065006300740069006f006e00000013000000010000005b0000000a0000000053006f00720074000000420000001e000000700072006f0070003400320039003400390036003700320039003500000000001c0000000100000030f125b7ef471a10a5f102608c9eebac0a00000001000000250000001400000000470072006f0075007000560069006500770000000b000000000000001b0000000a000000004d006f006400650000001300000006000000230000001200000000490063006f006e00530069007a00650000001300000030000000bd000000100000000043006f006c0049006e0066006f000000420000001e000000700072006f00700034003200390034003900360037003200390035000000000078000000fddfdffd100000000000000000000000040000001800000030f125b7ef471a10a5f102608c9eebac0a0000001001000030f125b7ef471a10a5f102608c9eebac0e0000007800000030f125b7ef471a10a5f102608c9eebac040000007800000030f125b7ef471a10a5f102608c9eebac0c000000500000002f0000001e00000000470072006f0075007000420079004b00650079003a00500049004400000013000000000000001f0000000e00000000460046006c00610067007300000013000000010020413100000020000000004c006f0067006900630061006c0056006900650077004d006f0064006500000013000000020000000000000000000000"))
    RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced","IconsOnly","REG_DWORD",1)
    RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams","Settings","REG_BINARY",Binary("0x080000000600000001000000e525f165e17b1048ba9dd271c8432ce304000000be01000043000000"))
    RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Defaults","{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}","REG_BINARY",Binary("0x1c000000000000000000000000000000000000000000000000000000f1f1f1f114000000000000000000000000000000d0020000cc0200003153505305d5cdd59c2e1b10939708002b2cf9ae830000002200000000470072006f0075007000420079004b00650079003a0046004d005400490044000000080000004e0000007b00300030003000300030003000300030002d0030003000300030002d0030003000300030002d0030003000300030002d003000300030003000300030003000300030003000300030007d0000000000330000002200000000470072006f00750070004200790044006900720065006300740069006f006e00000013000000010000005b0000000a0000000053006f00720074000000420000001e000000700072006f0070003400320039003400390036003700320039003500000000001c0000000100000030f125b7ef471a10a5f102608c9eebac0a00000001000000250000001400000000470072006f0075007000560069006500770000000b000000000000001b0000000a000000004d006f006400650000001300000006000000230000001200000000490063006f006e00530069007a00650000001300000030000000bd000000100000000043006f006c0049006e0066006f000000420000001e000000700072006f00700034003200390034003900360037003200390035000000000078000000fddfdffd100000000000000000000000040000001800000030f125b7ef471a10a5f102608c9eebac0a0000001001000030f125b7ef471a10a5f102608c9eebac0e0000007800000030f125b7ef471a10a5f102608c9eebac040000007800000030f125b7ef471a10a5f102608c9eebac0c000000500000002f0000001e00000000470072006f0075007000420079004b00650079003a00500049004400000013000000000000001f0000000e00000000460046006c00610067007300000013000000010020413100000020000000004c006f0067006900630061006c0056006900650077004d006f0064006500000013000000020000000000000000000000"))
    RegWrite("HKU\.default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced","IconsOnly","REG_DWORD",1)
    RegWrite("HKU\.default\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams","Settings","REG_BINARY",Binary("0x080000000600000001000000e525f165e17b1048ba9dd271c8432ce304000000be01000043000000"))
    RegWrite("HKU\.default\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Defaults","{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}","REG_BINARY",Binary("0x1c000000000000000000000000000000000000000000000000000000f1f1f1f114000000000000000000000000000000d0020000cc0200003153505305d5cdd59c2e1b10939708002b2cf9ae830000002200000000470072006f0075007000420079004b00650079003a0046004d005400490044000000080000004e0000007b00300030003000300030003000300030002d0030003000300030002d0030003000300030002d0030003000300030002d003000300030003000300030003000300030003000300030007d0000000000330000002200000000470072006f00750070004200790044006900720065006300740069006f006e00000013000000010000005b0000000a0000000053006f00720074000000420000001e000000700072006f0070003400320039003400390036003700320039003500000000001c0000000100000030f125b7ef471a10a5f102608c9eebac0a00000001000000250000001400000000470072006f0075007000560069006500770000000b000000000000001b0000000a000000004d006f006400650000001300000006000000230000001200000000490063006f006e00530069007a00650000001300000030000000bd000000100000000043006f006c0049006e0066006f000000420000001e000000700072006f00700034003200390034003900360037003200390035000000000078000000fddfdffd100000000000000000000000040000001800000030f125b7ef471a10a5f102608c9eebac0a0000001001000030f125b7ef471a10a5f102608c9eebac0e0000007800000030f125b7ef471a10a5f102608c9eebac040000007800000030f125b7ef471a10a5f102608c9eebac0c000000500000002f0000001e00000000470072006f0075007000420079004b00650079003a00500049004400000013000000000000001f0000000e00000000460046006c00610067007300000013000000010020413100000020000000004c006f0067006900630061006c0056006900650077004d006f0064006500000013000000020000000000000000000000"))
EndFunc

Func _InitVBox()
    _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\vboxguest.inf_amd64_neutral_c1c3a23271b0e375\vboxguest.inf",1)
    _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\vboxmouse.inf_amd64_neutral_7f412c49ccbbddff\vboxmouse.inf",1)
    _InitGFX()
    _InitService("vboxservice")
EndFunc

Func _InitGFX()
    If $sHost="VirtualBox" Then
        _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\vboxvideowddm.inf_amd64_neutral_4abeb766b91c5f62\vboxvideowddm.inf",1)
    Else
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\nv_lh.inf_amd64_neutral_bc69f20e3115af59\nv_LH.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\atiilhag.inf_amd64_neutral_0a660e899f5038a2\atiilhag.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\igdlh64.inf_amd64_neutral_fc2bea7bf411a3c4\igdlh64.inf",1)
;~         _Exec("drvload.exe "&@SystemDir&"\DriverStore\FileRepository\igdlh.inf_amd64_neutral_54a12b57f547d08e\igdlh.inf",1)
    EndIf
EndFunc

Func _RegDll($sFile,$iWow64=0,$iWait=1,$iHide=1)
    Local $sCmd
    If $iWow64 Then
        $sCmd=@WindowsDir&'\SysWow64\regsvr32.exe'
    Else
        $sCmd=@SystemDir&'\regsvr32.exe'
    EndIf
    If $iWait Then
        RunWait($sCmd&' /s "'&$sFile&'"')
    Else
        Run($sCmd&' /s "'&$sFile&'"')
    EndIf
EndFunc

Func _EnableAero()
;~     Local $hPrev = _WinAPI_GetThreadDesktop(_WinAPI_GetCurrentThreadId())
    Local $hDesktop = _WinAPI_CreateDesktop('SetAero', BitOR($DESKTOP_CREATEWINDOW, $DESKTOP_SWITCHDESKTOP))
    If Not $hDesktop Then
        MsgBox(64, 'Error', 'Unable to create desktop.')
        Exit
    EndIf
;~     _WinAPI_SwitchDesktop($hDesktop)
    Local $pText = _WinAPI_CreateString('SetAero')
    Local $tProcess = DllStructCreate($tagPROCESS_INFORMATION)
    Local $tStartup = DllStructCreate($tagSTARTUPINFO)
    DllStructSetData($tStartup, 'Size', DllStructGetSize($tStartup))
    DllStructSetData($tStartup, 'Desktop', $pText)
    If _WinAPI_CreateProcess('', @SystemDir & '\System.Init.exe Aero', 0, 0, 0, $CREATE_NEW_PROCESS_GROUP, 0, 0, $tStartup, $tProcess) Then
        ProcessWaitClose(DllStructGetData($tProcess, 'ProcessID'))
    EndIf
;~     _WinAPI_SwitchDesktop($hPrev)
    _WinAPI_CloseDesktop($hDesktop)
    _WinAPI_FreeMemory($pText)
EndFunc

Func _Exec($sCmd,$iHide=1,$iNoWait=0)
    Local $vRet
    _Log("Exec: "&$sCmd,1)
	If $iNoWait Then
        If $iHide Then
            Run($sCmd, @SystemDir, @SW_HIDE, 65536)
        Else
            Run($sCmd, @SystemDir)
        EndIf
	Else
        If $iHide Then
            $vRet=RunWait($sCmd, @SystemDir, @SW_HIDE, 65536)
        Else
            $vRet=RunWait($sCmd, @SystemDir)
        EndIf
        _Log("Return -> "&$vRet,2)
	EndIf
EndFunc

Func _InitService($sName,$iProg=-1)
    If $iProg<>-1 Then _StatusUI("Starting "&$sName&" service",$iProg)
    _Log("Starting "&$sName,2)
    If Not _Service_Exists($sName) Then
        _Log("Failed -> Service does not exist.",3)
        Return SetError(1,0,0)
    EndIf
    Local $aRet=_Service_QueryStatus($sName)
    If $aRet[1]=$SERVICE_RUNNING Then
        _Log("Failed -> Already Running.",3)
        Return SetError(2,0,0)
    EndIf
    If $aRet[1]<>$SERVICE_STOPPED Then
        _Log("Failed -> Service not in a stopped state.",3)
        Return SetError(3,0,0)
    EndIf
    _Service_Start($sName)
    $aRet=_Service_QueryStatus($sName)
    Sleep(125)
    If $aRet[1]=$SERVICE_START_PENDING Then
        Local $hTimer=TimerInit()
        Do
            If TimerDiff($hTimer)>=10000 Then
                _Log("Failed -> Timed out",3)
                Return SetError(4,0,0)
            EndIf
            $aRet=_Service_QueryStatus($sName)
            If $aRet[1]=$SERVICE_RUNNING Or $aRet[1]=$SERVICE_STOPPED Then ExitLoop
            Sleep(50)
        Until False
    EndIf
    If $aRet[1]=$SERVICE_RUNNING Then
        _Log("Success",3)
        Return SetError(0,0,1)
    ElseIf $aRet[1]=$SERVICE_STOPPED Then
        _Log("Failed -> Service failed to start.",3)
        Return SetError(5,0,0)
    EndIf
EndFunc

Func _Log($sString,$iTab=0)
    Local $sMsg="["&@YEAR&"."&@MON&"."&@MDAY&","&@HOUR&":"&@MIN&":"&@SEC&"] "&_StringRepeat("  ",$iTab)&$sString
    FileWriteLine($sSystemLog,$sMsg)
    ConsoleWrite($sMsg&@CRLF)
EndFunc

Func _SetWorkgroup($sName)
	Local $tName=DllStructCreate("wchar["&(15+1)&"]")
	DllStructSetData($tName,1,StringUpper($sName))
	Local $aRet=DllCall("netapi32.dll", "dword", "NetJoinDomain","ptr",0,"ptr",DllStructGetPtr($tName),"ptr",0,"ptr",0,"ptr",0, "dword", 0x00000040)
	If $aRet[0]<>0 Then
		Return SetError(_WinAPI_GetLastError(),0,0)
    EndIf
    Return SetError(0,0,1)
EndFunc

Func _SetComputerName($sName)
	Local $sLogonKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	Local $sCtrlKey = "HKLM\SYSTEM\CurrentControlSet"
	If StringRegExp($sName, '|/|:|*|?|"|<|>|.|,|~|!|@|#|$|%|^|&|(|)|;|{|}|_|=|+|[|]|x60' & "|'", 0) = 1 Then Return SetError(1,0,0)
	Local $aRet = DllCall("Kernel32.dll","BOOL","SetComputerNameEx","int",5,"str",$sName)
	If $aRet[0] = 0 Then
        Return SetError(_WinAPI_GetLastError(),0,0)
	EndIf
	RegWrite($sCtrlKey&"\Control\ComputernameActive\Computername","ComputerName","REG_SZ",$sName)
	RegWrite($sCtrlKey&"\Control\Computername\Computername","ComputerName","REG_SZ",$sName)
	RegWrite($sCtrlKey&"\Services\Tcpip\Parameters","Hostname","REG_SZ",$sName)
	RegWrite($sCtrlKey&"\Services\Tcpip\Parameters","NV Hostname","REG_SZ",$sName)
	RegWrite($sLogonKey,"AltDefaultDomainName","REG_SZ",$sName)
	RegWrite($sLogonKey,"DefaultDomainName","REG_SZ",$sName)
	RegWrite("HKEY_USERS\.Default\Software\Microsoft\Windows Media\WMSDK\General","Computername","REG_SZ",$sName)
	RegWrite($sCtrlKey&"\Control\Session Manager\Environment","Computername","REG_SZ",$sName)
	_SetEnv("ComputerName",$sName)
    If @error Then
        Return SetError(@Error,@extended+10,0)
    EndIf
    Return SetError(0,0,1)
EndFunc

Func _SetEnv($sVar, $sData)
	RegWrite("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", $sVar, "REG_SZ", $sData)
	Local $aRet = DllCall("Kernel32.dll", "BOOL", "SetEnvironmentVariable", "str", $sVar, "str", $sData)
	If $aRet[0] = 0 Then
		Return SetError(_WinAPI_GetLastError(),0,0)
	EndIf
	$aRet = DllCall("user32.dll", "lresult", "SendMessageTimeoutW", "hwnd", 65535, "dword", 26, "ptr", 0, "wstr", "Environment", "dword", 2, "dword", 5000, "dword_ptr*", 0)
	If $aRet[0] = 0 Then
		Return SetError(_WinAPI_GetLastError(),1,0)
	EndIf
    Return SetError(0,0,1)
EndFunc

Func DisplayChangeRes($width, $height, $bpp, $freq)
    Local Const $DM_PELSWIDTH = 0x00080000
    Local Const $DM_PELSHEIGHT = 0x00100000
    Local Const $DM_BITSPERPEL = 0x00040000
    Local Const $DM_DISPLAYFREQUENCY = 0x00400000
    Local Const $CDS_TEST = 0x00000002
    Local Const $CDS_UPDATEREGISTRY = 0x00000001
    Local Const $DISP_CHANGE_RESTART = 1
    Local Const $DISP_CHANGE_SUCCESSFUL = 0
    Local Const $HWND_BROADCAST = 0xffff
    Local Const $WM_DISPLAYCHANGE = 0x007E
    Local $DEVMODE = DllStructCreate("byte[32];int[10];byte[32];int[6]")
    Local $b = DllCall("user32.dll", "int", "EnumDisplaySettings", "ptr", 0, "long", 0, "ptr", DllStructGetPtr($DEVMODE))
    If @error Then
        $b = 0
    Else
        $b = $b[0]
    EndIf
    If $b <> 0 Then
        DllStructSetData($DEVMODE, 2, BitOR($DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_BITSPERPEL, $DM_DISPLAYFREQUENCY), 5)
        DllStructSetData($DEVMODE, 4, $width, 2)
        DllStructSetData($DEVMODE, 4, $height, 3)
        DllStructSetData($DEVMODE, 4, $bpp, 1)
        DllStructSetData($DEVMODE, 4, $freq, 5)
        $b = DllCall("user32.dll", "int", "ChangeDisplaySettings", "ptr", DllStructGetPtr($DEVMODE), "int", $CDS_TEST)
        If @error Then
            $b = -1
        Else
            $b = $b[0]
        EndIf
        Select
            Case $b = $DISP_CHANGE_RESTART
                $DEVMODE = 0
                Return 2
            Case $b = $DISP_CHANGE_SUCCESSFUL
                DllCall("user32.dll", "int", "ChangeDisplaySettings", "ptr", DllStructGetPtr($DEVMODE), "int", $CDS_UPDATEREGISTRY)
                DllCall("user32.dll", "int", "SendMessage", "hwnd", $HWND_BROADCAST, "int", $WM_DISPLAYCHANGE, _
                        "int", $bpp, "int", $height * 2 ^ 16 + $width)
                $DEVMODE = 0
                Return 1
            Case Else
                $DEVMODE = 0
                Return $b
        EndSelect
    EndIf
EndFunc  ;==>DisplayChangeRes

Func _VCheck()
    Local $strComputer = ".", $sMake, $sModel, $sBIOSVersion, $bIsVM, $sVMPlatform
    Local $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    If IsObj($colItems) Then
        For $objItem In $colItems
            $sMake = $objItem.Manufacturer
            $sModel = $objItem.Model
        Next
    EndIf

    $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_BIOS", "WQL", 0x10 + 0x20)
    If IsObj($colItems) Then
        For $objItem In $colItems
            $sBIOSVersion = $objItem.SMBIOSBIOSVersion
        Next
    EndIf
    $bIsVM = False
    $sVMPlatform = ""
    If $sModel = "Virtual Machine" Then
        $sVMPlatform = "Hyper-V"
        $bIsVM = True
    ElseIf $sModel = "VMware Virtual Platform" Then
        $sVMPlatform = "VMware"
        $bIsVM = True
    ElseIf $sModel = "VirtualBox" Then
        $bIsVM = True
        $sVMPlatform = "VirtualBox"
    EndIf
    If $bIsVM Then Return $sVMPlatform
    Return "Bare Metal"
EndFunc   ;==>_VCheck

Func _RefreshFolderAttribs($sPath)
    Local $aFiles=_FileListToArrayRec($sPath,"*",2,1)
    If @error Then Return
    For $i=1 To $aFiles[0]
        FileSetAttrib($sPath&"\"&$aFiles[$i],"-RASHNOT")
        If $aFiles[$i]="Startup" Then
            FileSetAttrib($sPath&"\"&$aFiles[$i],"+H")
        Else
            FileSetAttrib($sPath&"\"&$aFiles[$i],"+S")
        EndIf
    Next
    Return
EndFunc

Func _StatusUI($sMsg,$iProg)
    GUICtrlSetData($idUIStatus,"Status: "&$sMsg)
    GUICtrlSetData($idUIProgress,$iProg)
EndFunc
Func _InitUI($iProgress=0)
    Local $iWidth=@DesktopWidth/2
    Local $iHeight=@DesktopHeight/(8+4)
    Local $iPosX=(@DesktopWidth/2)-($iWidth/2)
    Local $iPosY=(@DesktopHeight/2)-$iHeight
    Local Static $iAlpha = 0xABCDEF
    GUIRegisterMsg(132, "WM_NCHITTEST")
    If _WinAPI_DwmIsCompositionEnabled() Then
        $iUIDWM=1
        $hUI=GUICreate("",$iWidth-14,$iHeight-34,$iPosX+1,$iPosY+19,BitOR($ws_overlapped,$ws_popup,$ws_sizebox),BitOR($ws_ex_topmost,$ws_ex_toolwindow,$ws_ex_layered))
        GUISetBkColor($iAlpha)
        _WinAPI_DwmGetColorizationColor()
        _WinAPI_SetLayeredWindowAttributes($hUI,$iAlpha,0,$lwa_colorkey)
        _WinAPI_DwmExtendFrameIntoClientArea($hUI)
        $idUIStatus= GUICtrlCreateLabel("Status: ",16,2,$iWidth)
        $idUIProgress= GUICtrlCreateProgress(4,$iHeight-($iHeight/2)-16,$iWidth-22,$iHeight/4)
    Else
        $iUIDWM=0
        $hUI1=GUICreate("Initializing...", $iWidth+2,$iHeight,$iPosX,$iPosY,0,BitOR($ws_ex_topmost,$ws_ex_toolwindow))
        $idUIStatus= GUICtrlCreateLabel("Status: ",16,8,$iWidth)
        $idUIProgress= GUICtrlCreateProgress(4,$iHeight-($iHeight/2)-10,$iWidth-14,$iHeight/4)
    EndIf
    GUICtrlSetData($idUIProgress,$iProgress)
    $iUIProgressLast=$iProgress
    GUISetState(@SW_SHOW,$hUI)
    AdlibRegister("_RefreshUI",125)
EndFunc

Func wm_nccalcsize($hwnd, $imsg, $wparam, $lparam)
	Return 0
EndFunc

Func wm_nchittest($hwnd, $imsg, $wparam, $lparam)
	Return 0
EndFunc

Func _RefreshUI()
    Local $iActiveDWM=_WinAPI_DwmIsCompositionEnabled()
	If $iUIDWM=0 And $iActiveDWM Then
        GUIDelete($hUI)
        _InitUI($iUIProgressLast)
	ElseIf $iUIDWM=1 And Not $iActiveDWM Then
        GUIDelete($hUI)
        _InitUI($iUIProgressLast)
	EndIf
EndFunc

Func _NtQueryDirectoryArcNames()
    Local $aDrives=_NtQueryDirectoryDrives()
    Local $iDrive=-1
    Local $aArcNames[1][2],$iMax
    Local $aObjects = _NtQueryDirectoryObjects("ArcName")
    Local $sDriveRegEx="\\(Device\\Harddisk\d{1,2})\\(Partition\d{1,2})"
    Local $sDevice
    For $i=0 To UBound($aDrives,1)-1
        Local $aDevList=$aDrives[$i][1]
        If Not IsArray($aDevList) Then ContinueLoop
        For $j=0 To UBound($aObjects,1)-1
            If StringRegExp($aObjects[$j][2],$sDriveRegEx) Then
                Local $aDrive=StringRegExp($aObjects[$j][2],$sDriveRegEx,1)
                If @error Then ContinueLoop
                Local $aObject=_NtQueryDirectoryObjects($aDrive[0])
                If Not IsArray($aObject) Then ContinueLoop
                For $k=0 To UBound($aObject,1)-1
                    If $aObject[$k][0]=$aDrive[1] Then
                        For $l=1 To UBound($aDevList,1)-1
                            If $aDevList[$l]=$aObject[$k][2] Then
                                $iMax=UBound($aArcNames,1)
                                ReDim $aArcNames[$iMax+1][2]
                                $aArcNames[$iMax][0]=$aObjects[$j][0]
                                $aArcNames[$iMax][1]=$aDrives[$i][0]
                                $aArcNames[0][0]+=1
                            EndIf
                        Next
                    EndIf
                Next
                Else
                For $k=1 To UBound($aDevList,1)-1
                    If $aDevList[$k]=$aObjects[$j][2] Then
                        $iMax=UBound($aArcNames,1)
                        ReDim $aArcNames[$iMax+1][2]
                        $aArcNames[$iMax][0]=$aObjects[$j][0]
                        $aArcNames[$iMax][1]=$aDrives[$i][0]
                        $aArcNames[0][0]+=1
                    EndIf
                Next
            EndIf
        Next
    Next
    Return $aArcNames
EndFunc

Func _NtQueryDirectoryDrives()
    Local $aDrives[0][2]
    Local $aObjects = _NtQueryDirectoryObjects("GLOBAL??")
    For $i = 0 To UBound($aObjects) - 1
        If StringRegExp($aObjects[$i][0],"^[A-Za-z]\:$") Then
            Local $iMax=UBound($aDrives,1)
            ReDim $aDrives[$iMax+1][2]
            $aDrives[$iMax][1]=$aObjects[$i][2]
            $aDrives[$iMax][0]=$aObjects[$i][0]
        EndIf
    Next
    For $i=0 To UBound($aDrives)-1
        Dim $aEntries[2]
        $aEntries[0]=1
        $aEntries[1]=$aDrives[$i][1]
        For $j=0 To UBound($aObjects)-1
            If $aObjects[$j][0]=$aDrives[$i][0] Then ContinueLoop
            If $aObjects[$j][2]=$aDrives[$i][1] Then
                Local $iMax=UBound($aEntries,1)
                ReDim $aEntries[$iMax+1]
                $aEntries[$iMax]=$aObjects[$j][0]
                $aEntries[0]+=1
            EndIf
        Next
        If $aEntries[0]>1 Then $aDrives[$i][1]=$aEntries
    Next
    Return $aDrives
EndFunc

Func _NtQueryDirectoryObjects($sSource="\")
    If $sSource <> "\" Then $sSource = "\" & $sSource
    Local $tNameSource = _tName($sSource)
    Local $tUnicodeString = _RtlInitUnicodeString($tNameSource)
    Local $pUnicodeString = DllStructGetPtr($tUnicodeString)
    Local $tObject_Attributes = _InitializeObjectAttributes($pUnicodeString)
    Local $pObject_Attributes = DllStructGetPtr($tObject_Attributes)
    Local $hDirectory = _NtOpenDirectoryObject($pObject_Attributes, BitOR($DIRECTORY_TRAVERSE, $DIRECTORY_QUERY))
    Local $tData = 0
    Local $tName = 0
    Local $tType = 0
    Local $index = 0
    Local $bytes = 0
    Local $tBuffer = 0
    Local $aObjects[0]
    Local $taName = 0
    Local $tStr = 0
    Local $tAttr = 0
    Local $hLink = 0
    Local $taTarget = 0
    Local $tTarget = 0
    Local $tSTarget = 0
    $tBuffer = DllStructCreate("byte Data[32767]")
    Local $aRet = DllCall("Ntdll.dll", "LONG", "NtQueryDirectoryObject", "HANDLE", $hDirectory, "ptr", DllStructGetPtr($tBuffer), "ULONG", 32767, "BOOL", False, "BOOL", True, "ULONG*", 0, "ULONG*", 0)
    If @error Or $aRet[0] < 0 Then Return
    $index = $aRet[6]
    $bytes = $aRet[7]
    For $i = 0 To $index - 1
        $tData = DllStructCreate($sTagUNICODESTRING & $sTagUNICODESTRING, DllStructGetPtr($tBuffer) + ($i * 16))
        $tName = DllStructCreate("wchar wNameString[" & DllStructGetData($tData, 1) & "]", DllStructGetData($tData, 3))
        $tType = DllStructCreate("wChar wTypeString[" & DllStructGetData($tData, 4) & "]", DllStructGetData($tData, 6))
        $taName = _tName($tName.wNameString)
        $tStr = _RtlInitUnicodeString($taName)
        $tAttr = _InitializeObjectAttributes(DllStructGetPtr($tStr), 0, $hDirectory)
        $hLink = _NtOpenSymbolicLinkObject(DllStructGetPtr($tAttr), $GENERIC_READ)
        $taTarget = _tName("")
        $tTarget = _RtlInitUnicodeString($taTarget)
        $tTarget.MaximumLength = 512
        _NtQuerySymbolicLinkObject($hLink, $tTarget)
        $tSTarget = DllStructCreate("wchar wString[" & DllStructGetData($tTarget, 1) & "]", DllStructGetData($tTarget, 3))
        ReDim $aObjects[UBound($aObjects) + 1][3]
        $aObjects[UBound($aObjects) - 1][0] = $tName.wNameString
        $aObjects[UBound($aObjects) - 1][1] = $tType.wTypeString
        $aObjects[UBound($aObjects) - 1][2] = (IsDllStruct($tSTarget) = 1) ? $tSTarget.wString : ""
        $tType = 0
        $tName = 0
        $tData = 0
        $tSTarget = 0
        _WinAPI_CloseHandle($hLink)
    Next
    Return $aObjects
EndFunc   ;==>_GetObjects


Func _RtlInitUnicodeString($tSourceString)
    Local $tUnicodeString = DllStructCreate($sTagUNICODESTRING)
    DllCall("Ntdll.dll", "NONE", "RtlInitUnicodeString", "struct*", $tUnicodeString, "struct*", $tSourceString)
    If @error Then SetError(@error, 0, 0)
    Return $tUnicodeString
EndFunc   ;==>_RtlInitUnicodeString

Func _tName($String)
    Local $t = DllStructCreate("wchar String[512]")
    DllStructSetData($t, 1, $String)
    Return $t
EndFunc   ;==>_tName


Func _InitializeObjectAttributes($pObjectName, $ulAttributes = 0, $hRootDirectory = Null, $pSecurityDescriptor = Null)
    Local $tObject_Attributes = DllStructCreate($sTagOBJECT_ATTRIBUTES)
    DllStructSetData($tObject_Attributes, 1, DllStructGetSize($tObject_Attributes))
    DllStructSetData($tObject_Attributes, 2, $hRootDirectory)
    DllStructSetData($tObject_Attributes, 3, $pObjectName)
    Return $tObject_Attributes
EndFunc   ;==>_InitializeObjectAttributes


Func _NtQuerySymbolicLinkObject($hLinkHandle, $tLinkTarget)
    Local $aRet = DllCall("Ntdll.dll", "LONG", "NtQuerySymbolicLinkObject", "HANDLE", $hLinkHandle, "struct*", $tLinkTarget, "ULONG*", 0)
    If @error Then SetError(@error, 0, 0)
    If $aRet[0] <> 0 Then SetError(1, 0, 0)
EndFunc   ;==>_NtQuerySymbolicLinkObject

Func _NtOpenDirectoryObject($pObjectAttr, $AccessMask)
    Local $aRet = DllCall("Ntdll.dll", "LONG", "NtOpenDirectoryObject", "HANDLE*", 0, "ULONG", $AccessMask, "PTR", $pObjectAttr)
    If @error Then SetError(@error, 0, 0)
    If $aRet[0] < 0 Or $aRet[1] = 0 Then SetError(1, 0, 0)
    Return $aRet[1]

EndFunc   ;==>_NtOpenDirectoryObject

Func _NtOpenSymbolicLinkObject($pObjectAttr, $AccessMask)
    Local $aRet = DllCall("Ntdll.dll", "LONG", "NtOpenSymbolicLinkObject", "HANDLE*", 0, "ULONG", $AccessMask, "PTR", $pObjectAttr)
    If @error Then Return SetError(@error, 0, 0)
    If $aRet[0] < 0 Or $aRet[1] = 0 Then Return SetError(1, 0, 0)
    Return $aRet[1]
EndFunc   ;==>_NtOpenSymbolicLinkObject

; #FUNCTION# ====================================================================================================================
; Name...........: _ChangeWallpaper
; Description ...: Change Windows Wallpaper
; Syntax.........: _ChangeWallpaper($sImage, [$iStyle])
; Parameters ....: $sImage   - The path of the .bmp file
;                   $$iStyle  - The numeric value of desidered style
;                              0 Tiled
;                              1 Centered
;                              2 Stretched
;                              3 Fit (Windows 7 and later)
;                              4 Fill (Windows 7 and later)
;                               5 Screen Width
;                   $bResize   - Automatically resize th image if has a higher resolution than screen
; Return values .: On Success - Return the new file name.
;                  On Failure -
;                               @error = 1 The image doesn't exist
;                               @error = 2 The image is not a .bmp file
;                               @error = 3 Invalid style
;                               @error = 4 Style not supported by OS
;                               @error = 5 Unable to change the wallpaper
; Author ........: Nessie
; ===============================================================================================================================

Func _ChangeWallpaper($sImage, $iStyle = 0, $bResize = True)
    If Not FileExists($sImage) Then Return SetError(1, 0, "")

    Local $sImageExt = StringLower(StringRight($sImage, 4))

    Local $fWinVer = _GetWinVersion()
    If $sImageExt <> '.bmp' And $fWinVer < 6.0 Then
        Return SetError(2, 0, "")
    Else
        If $fWinVer >= 6.0 And $sImageExt <> '.bmp' And $sImageExt <> '.jpg' Then
            Return SetError(2, 0, "")
        EndIf
    EndIf
    If $iStyle < 0 Or $iStyle > 5 Then Return SetError(3, 0, "")
    If $fWinVer < 6.0 Then ; More info http://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx
        If $iStyle > 2 Then Return SetError(4, 0, "")
    EndIf
    Local $sWallpaperKey = "HKEY_CURRENT_USER\Control Panel\Desktop"
    Local $iTileWallPaper, $iWallpaperStyle
    Switch $iStyle
        Case 0
            $iTileWallPaper = 1
            $iWallpaperStyle = 0
        Case 1
            $iTileWallPaper = 0
            $iWallpaperStyle = 0
        Case 2
            $iTileWallPaper = 0
            $iWallpaperStyle = 2
        Case 3
            $iTileWallPaper = 0
            $iWallpaperStyle = 6
        Case 4
            $iTileWallPaper = 0
            $iWallpaperStyle = 10
        Case 5
            $iTileWallPaper = 0
            $iWallpaperStyle = 4
    EndSwitch

    RegWrite($sWallpaperKey, "TileWallPaper", "REG_SZ", $iTileWallPaper)
    If @error Then Return SetError(5, 0, "")
    RegWrite($sWallpaperKey, "WallpaperStyle", "REG_SZ", $iWallpaperStyle)
    If @error Then Return SetError(5, 0, "")

    ;Thanks to guinness for his advice
    ; Idea from here: http://www.autoitscript.com/forum/topic/19370-autoit-wrappers/page__st__280#entry652536
    ; $SPI_SETDESKWALLPAPER, $SPIF_UPDATEINIFILE and $SPIF_SENDCHANGE can be found on APIConstants.au3 included on WinAPIEx by Yashied
    ;Return _WinAPI_SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, DllStructGetPtr($tBuffer), BitOR($SPIF_UPDATEINIFILE, $SPIF_SENDCHANGE))
    Local $tBuffer = DllStructCreate('wchar Text[' & StringLen($sImage) + 1 & ']')
    DllStructSetData($tBuffer, 'Text', $sImage)
    Return _WinAPI_SystemParametersInfo(0x0014, 0, DllStructGetPtr($tBuffer), BitOR(0x0001, 0x0002))
    If @error Then Return SetError(5, 0, "")

    Return True
EndFunc   ;==>_ChangeWallpaper

Func _GetWinVersion()
    Local $sRet = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\', 'CurrentVersion')
    If @error Then Return SetError(0, 0, "")
    Return $sRet
EndFunc   ;==>_GetWinVersion

Func _ExitWindowsEx($iCode=0x00000005)
#cs
Private Const EWX_LOGOFF As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_POWEROFF As Long = 8
Private Const EWX_FORCE As Long = 4
Private Const EWX_FORCEIFHUNG As Long = 16
#ce
    Local $hToken = _Security__OpenProcessToken(_WinAPI_GetCurrentProcess(), $TOKEN_ALL_ACCESS)
    If $hToken Then
    ; $hToken it this process' token with $TOKEN_ALL_ACCESS access
    ; Enable SeDebugPrivilege for this token
    If _Security__SetPrivilege($hToken, $SE_SHUTDOWN_NAME, True) Then
    ;... Do whatever with this token now and here...
    ; MsgBox(262144, "TokenPrivileges", $SE_SHUTDOWN_NAME & " enabled!")

    DllCall("user32.dll", "int", "ExitWindowsEx", "uint",BitOR(0x00000000, 0x00000004), "dword", 0)

    ; Disable
    _Security__SetPrivilege($hToken, $SE_SHUTDOWN_NAME, False)
    ; MsgBox(262144, "TokenPrivileges", $SE_SHUTDOWN_NAME & " disabled!")
    EndIf
    ; Close handle when done
    _WinAPI_CloseHandle($hToken)
    EndIf
EndFunc

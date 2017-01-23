Attribute VB_Name = "mdl_IOTVB10"
Global Const NODEVICE% = -1
Global Const TERMIN = 1
Global Const TERMOUT = 2
Global Const TURNON = 1
Global Const TURNOFF = 0
Global Const ALL = -1
Global Const WHILESRQ = -2
Global Const UNTILRSV = -3

'These Constants are used with the Arm and Disarm functions

Global Const acSRQ = &H400
Global Const acPeripheral = &H200
Global Const acController = &H100
Global Const acTrigger = &H80
Global Const acClear = &H40
Global Const acTalk = &H20
Global Const acListen = &H10
Global Const acIdle = &H8
Global Const acByteIn = &H4
Global Const acByteOut = &H2
Global Const acChange = &H1

' These Constants are the returned values in the compStat parameter for
' Enter and Output functions

Global Const ccCount = &H1
Global Const ccBuffer = &H2
Global Const ccTerm = &H4
Global Const ccEnd = &H8
Global Const ccChange = &H10
Global Const ccStop = &H20
Global Const ccDone = &H4000
Global Const ccError = &H8000

'This Type is used with the Enter, Output, and Term functions

Type terms
   eoi       As Integer
   nChar     As Integer
   EightBits As Integer
   term1     As Integer
   term2     As Integer
End Type

'This Type is used with the Status function

Type IeeeStatus
   SC          As Integer
   CA          As Integer
   Primaddr    As Integer
   Secaddr     As Integer
   SRQ         As Integer
   addrChange  As Integer
   talker      As Integer
   listener    As Integer
   triggered   As Integer
   cleared     As Integer
   transfer    As Integer
   byteIn      As Integer
   byteOut     As Integer
End Type

' Conversion routines to account for differences between the BASIC
' structures and the structures expected by the DLL

' This is a list of C routines directly accessible from Visual Basic

Declare Function Abort Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function Arm Lib "Drvr488.Dll" (ByVal devHandle%, ByVal armcond%) As Integer
Declare Function AutoRemote Lib "Drvr488.Dll" (ByVal devHandle%, ByVal flag%) As Integer
Declare Function Buffered Lib "Drvr488.Dll" (ByVal devHandle%) As Long
Declare Function BusAddress Lib "Drvr488.Dll" (ByVal devHandle%, ByVal prim%, ByVal sec%) As Integer
Declare Function CheckListener Lib "Drvr488.Dll" (ByVal devHandle%, ByVal prim%, ByVal sec%) As Integer
Declare Function ioclear Lib "Drvr488.Dll" Alias "Clear" (ByVal devHandle%) As Integer
Declare Function ClearList Lib "Drvr488.Dll" (listFirstElement%) As Integer
Declare Function ClockFrequency Lib "Drvr488.Dll" (ByVal devHandle%, ByVal freq%) As Integer
Declare Function ioclose Lib "Drvr488.Dll" Alias "Close" (ByVal devHandle%) As Integer
Declare Function ControlLine Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function Disarm Lib "Drvr488.Dll" (ByVal devHandle%, ByVal armcond%) As Integer
Declare Function DmaChannel Lib "Drvr488.Dll" (ByVal devHandle%, ByVal channel%) As Integer
Declare Function EnterXdll Lib "Drvr488.Dll" Alias "EnterX" (ByVal devHandle%, dat As Any, ByVal count&, ByVal forceAddress%, term As Any, ByVal async%, compStat%) As Long
Declare Function EnterXIdll Lib "Drvr488.Dll" Alias "EnterX" (ByVal devHandle%, dat As Integer, ByVal count&, ByVal forceAddress%, term As Any, ByVal async%, compStat%) As Long
Declare Function ioError Lib "Drvr488.Dll" Alias "Error" (ByVal devHandle%, ByVal display%) As Integer
Declare Function Finish Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function FindListeners Lib "Drvr488.Dll" (ByVal devHandle%, ByVal primary%, listFirstElement%, ByVal limit%) As Integer
Declare Function GetError Lib "Drvr488.Dll" (ByVal devHandle%, ByVal errorText$) As Integer
Declare Function GetErrorList Lib "Drvr488.Dll" (listFirstElement%, ByVal errorText$, errorHandle%) As Integer
Declare Function Hello Lib "Drvr488.Dll" (ByVal devHandle%, ByVal helloText$) As Integer
Declare Function IntLevel Lib "Drvr488.Dll" (ByVal devHandle%, ByVal channel%) As Integer
Declare Function IOAddress Lib "Drvr488.Dll" (ByVal devHandle%, ByVal addr%) As Integer
Declare Function KeepDevice Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function Listen Lib "Drvr488.Dll" (ByVal devHandle%, ByVal prim%, ByVal sec%) As Integer
Declare Function ioLocal Lib "Drvr488.Dll" Alias "Local" (ByVal devHandle%) As Integer
Declare Function LocalList Lib "Drvr488.Dll" (listFirstElement%) As Integer
Declare Function Lol Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function MakeDevice Lib "Drvr488.Dll" (ByVal devHandle%, ByVal nam$) As Integer
Declare Function MyListenAddr Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function MyTalkAddr Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function openname Lib "Drvr488.Dll" (ByVal nam As String) As Integer
Declare Function OutputXdll Lib "Drvr488.Dll" Alias "OutputX" (ByVal devHandle%, dat As Any, ByVal count&, ByVal last%, ByVal forceAddress%, term As Any, ByVal async%, compStat As Any) As Long
Declare Function OutputXIdll Lib "Drvr488.Dll" Alias "OutputX" (ByVal devHandle%, dat As Integer, ByVal count&, ByVal last%, ByVal forceAddress%, term As Any, ByVal async%, compStat As Any) As Long
Declare Function PassControl Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function PPoll Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function PPollConfig Lib "Drvr488.Dll" (ByVal devHandle%, ByVal ppresponse%) As Integer
Declare Function PPollDisable Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function PPollDisableList Lib "Drvr488.Dll" (listFirstElement%) As Integer
Declare Function PPollUnconfig Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function remote Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function RemoteList Lib "Drvr488.Dll" (listFirstElement%) As Integer
Declare Function RemoveDevice Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function Request Lib "Drvr488.Dll" (ByVal devHandle%, ByVal spstatus%) As Integer
Declare Function ioReset Lib "Drvr488.Dll" Alias "Reset" (ByVal devHandle%) As Integer
Declare Function ioResume Lib "Drvr488.Dll" Alias "Resume" (ByVal devHandle%, ByVal monitor%) As Integer
Declare Function SendCmd Lib "Drvr488.Dll" (ByVal devHandle%, dat As Any, ByVal count%) As Integer
Declare Function SendData Lib "Drvr488.Dll" (ByVal devHandle%, dat As Any, ByVal count%) As Integer
Declare Function SendEoi Lib "Drvr488.Dll" (ByVal devHandle%, dat As Any, ByVal count%) As Integer
Declare Function SPoll Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function SPollList Lib "Drvr488.Dll" (listFirstElement%, result%, ByVal flag%) As Integer
Declare Function Statusdll Lib "Drvr488.Dll" Alias "Status" (ByVal devHandle%, stat As IeeeStatus) As Integer
Declare Function ioStop Lib "Drvr488.Dll" Alias "Stop" (ByVal devHandle%) As Integer
Declare Function SysController Lib "Drvr488.Dll" (ByVal devHandle%, ByVal flag%) As Integer
Declare Function Talk Lib "Drvr488.Dll" (ByVal devHandle%, ByVal prim%, ByVal sec%) As Integer
Declare Function Termdll Lib "Drvr488.Dll" Alias "Term" (ByVal devHandle%, term As terms, ByVal termType%) As Integer
Declare Function TimeOut Lib "Drvr488.Dll" (ByVal devHandle%, ByVal millisec&) As Integer
Declare Function Trigger Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function TriggerList Lib "Drvr488.Dll" (listFirstElement%) As Integer
Declare Function UnListen Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function UnTalk Lib "Drvr488.Dll" (ByVal devHandle%) As Integer
Declare Function ioWait Lib "Drvr488.Dll" Alias "Wait" (ByVal devHandle%) As Integer


'*******************************************************
'This is the beginning of the application-oriented
'declarations
'*******************************************************

' These declarations are necessary for unloading the
' driver application and its icon
Declare Function FindWindow Lib "USER" (ByVal data1$, ByVal data2$) As Integer
Declare Function SendMessage Lib "USER" (ByVal winHandle%, ByVal Message%, ByVal wp%, ByVal lp As Long) As Integer

' Convert DLL status struct to return to BASIC application
Sub ConvertStatus(stat As IeeeStatus)
stat.byteOut = stat.talker And 255
stat.byteIn = (stat.addrChange And &HFF00) \ 256
stat.transfer = stat.addrChange And 255
stat.cleared = (stat.SRQ And &HFF00) \ 256
stat.triggered = stat.SRQ And 255
stat.listener = (stat.Secaddr And &HFF00) \ 256
stat.talker = stat.Secaddr And 255
stat.addrChange = (stat.Primaddr And &HFF00) \ 256
stat.SRQ = stat.Primaddr And 255
stat.Secaddr = (stat.CA And &HFF00) \ 256
stat.Primaddr = stat.CA And 255
stat.CA = (stat.SC And &HFF00) \ 256
stat.SC = stat.SC And 255
End Sub

' Convert BASIC terminator struct to DLL
Sub ConvertTerm(Terminator As terms)
Terminator.eoi = Terminator.eoi And 255 Or (Terminator.nChar And 255) * 256
If Terminator.EightBits Then
    Terminator.nChar = &HFF00
Else
    Terminator.nChar = 0
End If
Terminator.EightBits = Terminator.term1
Terminator.term1 = Terminator.term2
End Sub

' Short form of EnterX
Function Enter(ByVal devHandle%, dat$) As Long
Enter = EnterXdll(ByVal devHandle%, ByVal dat$, Len(dat$), 1, ByVal 0&, 0, 0)
End Function

' Short form of EnterXI
Function EnterI(ByVal devHandle%, datx%()) As Long
Dim newTerm As terms
newTerm.eoi = 1
Size% = 2 * UBound(datx)
EnterI = EnterXIdll(ByVal devHandle%, datx(0), Size%, 1, newTerm, 0, 0)
End Function

' Short form of EnterX
Function EnterMore(ByVal devHandle%, dat$) As Long
EnterMore = EnterXdll(ByVal devHandle%, ByVal dat$, Len(dat$), 0, ByVal 0&, 0, 0)
End Function

' Short form of EnterXI
Function EnterMoreI(ByVal devHandle%, datx%()) As Long
Dim newTerm As terms
newTerm.eoi = 1
Size% = 2 * UBound(datx)
EnterMoreI = EnterXIdll(ByVal devHandle%, datx(0), Size%, 0, newTerm, 0, 0)
End Function

' Short form of EnterX
Function EnterN(ByVal devHandle%, dat$, ByVal count&) As Long
EnterN = EnterXdll(ByVal devHandle%, ByVal dat$, ByVal count&, 1, ByVal 0&, 0, 0)
End Function

' Short form of EnterXI
Function EnterNI(ByVal devHandle%, dat%, ByVal count&) As Long
Dim newTerm As terms
newTerm.eoi = 1
EnterNI = EnterXIdll(ByVal devHandle%, dat%, ByVal count&, 1, newTerm, 0, 0)
End Function

' Short form of EnterX
Function EnterNMore(ByVal devHandle%, dat$, ByVal count&) As Long
EnterNMore = EnterXdll(ByVal devHandle%, ByVal dat$, ByVal count&, 0, ByVal 0&, 0, 0)
End Function

' Short form of EnterXI
Function EnterNMoreI(ByVal devHandle%, dat%, ByVal count&) As Long
Dim newTerm As terms
newTerm.eoi = 1
EnterNMoreI = EnterXIdll(ByVal devHandle%, dat%, ByVal count&, 0, newTerm, 0, 0)
End Function

' Reformat Term for EnterX
Function Enterx(ByVal devHandle%, dat$, ByVal count&, ByVal forceAddress%, Terminator As terms, ByVal async%, compStat%) As Long
Dim newTerm As terms
newTerm.eoi = Terminator.eoi
newTerm.nChar = Terminator.nChar
newTerm.EightBits = Terminator.EightBits
newTerm.term1 = Terminator.term1
newTerm.term2 = Terminator.term2
Call ConvertTerm(newTerm)
Enterx = EnterXdll(devHandle%, ByVal dat$, count&, forceAddress%, newTerm, async%, compStat%)
End Function

' Reformat Term for EnterXI
Function EnterXI(ByVal devHandle%, dat%, ByVal count&, ByVal forceAddress%, Terminator As terms, ByVal async%, compStat%) As Long
Dim newTerm As terms
newTerm.eoi = Terminator.eoi
newTerm.nChar = Terminator.nChar
newTerm.EightBits = Terminator.EightBits
newTerm.term1 = Terminator.term1
newTerm.term2 = Terminator.term2
Call ConvertTerm(newTerm)
EnterXI = EnterXIdll(devHandle%, dat%, count&, forceAddress%, newTerm, async%, compStat%)
End Function

' Short form of OutputX
Function iooutput(ByVal devHandle%, ByVal dat$) As Long
 iooutput = OutputXdll(ByVal devHandle%, ByVal dat$, Len(dat$), 1, 1, ByVal 0&, 0, 0)
End Function

' Short form of OutputXI
Function OutputI(ByVal devHandle%, datx%()) As Long
Size% = 2 * UBound(datx)
OutputI = OutputXIdll(ByVal devHandle%, datx(0), Size%, 1, 1, ByVal 0&, 0, 0)
End Function

' Short form of OutputX
Function OutputMore(ByVal devHandle%, ByVal dat$) As Long
OutputMore = OutputXdll(ByVal devHandle%, ByVal dat$, Len(dat$), 1, 0, ByVal 0&, 0, 0)
End Function

' Short form of OutputXI
Function OutputMoreI(ByVal devHandle%, datx%()) As Long
Size% = 2 * UBound(datx)
OutputMoreI = OutputXIdll(ByVal devHandle%, datx(0), Size%, 1, 0, ByVal 0&, 0, 0)
End Function

' Short form of OutputX
Function OutputN(ByVal devHandle%, ByVal dat$, ByVal count&) As Long
OutputN = OutputXdll(ByVal devHandle%, ByVal dat$, ByVal count&, 0, 1, ByVal 0&, 0, 0)
End Function

' Short form of OutputXI
Function OutputNI(ByVal devHandle%, dat%, ByVal count&) As Long
OutputNI = OutputXIdll(ByVal devHandle%, dat%, ByVal count&, 0, 1, ByVal 0&, 0, 0)
End Function

' Short form of OutputX
Function OutputNMore(ByVal devHandle%, ByVal dat$, ByVal count&) As Long
OutputNMore = OutputXdll(ByVal devHandle%, ByVal dat$, ByVal count&, 0, 0, ByVal 0&, 0, 0)
End Function

' Short form of OutputXI
Function OutputNMoreI(ByVal devHandle%, dat%, ByVal count&) As Long
OutputNMoreI = OutputXIdll(ByVal devHandle%, dat%, ByVal count&, 0, 0, ByVal 0&, 0, 0)
End Function

' Reformat Term for OutputX
Function OutputX(ByVal devHandle%, ByVal dat$, ByVal count&, ByVal last%, ByVal forceAddress%, Terminator As terms, ByVal async%, compStat%) As Long
Dim newTerm As terms
newTerm.eoi = Terminator.eoi
newTerm.nChar = Terminator.nChar
newTerm.EightBits = Terminator.EightBits
newTerm.term1 = Terminator.term1
newTerm.term2 = Terminator.term2
Call ConvertTerm(newTerm)
OutputX = OutputXdll(devHandle%, ByVal dat$, count&, last%, forceAddress%, newTerm, async%, compStat%)
End Function

' Reformat Term for OutputXI
Function OutputXI(ByVal devHandle%, dat%, ByVal count&, ByVal last%, ByVal forceAddress%, Terminator As terms, ByVal async%, compStat%) As Long
Dim newTerm As terms
newTerm.eoi = Terminator.eoi
newTerm.nChar = Terminator.nChar
newTerm.EightBits = Terminator.EightBits
newTerm.term1 = Terminator.term1
newTerm.term2 = Terminator.term2
Call ConvertTerm(newTerm)
OutputXI = OutputXIdll(devHandle%, dat%, count&, last%, forceAddress%, newTerm, async%, compStat%)
End Function

' Shell for Status to reformat structure
Function Status(ByVal devHandle%, stat As IeeeStatus) As Integer
Status = Statusdll(devHandle%, stat)
Call ConvertStatus(stat)
End Function

' Shell for Term to reformat structure
Function term(ByVal devHandle%, Terminator As terms, ByVal termType%) As Integer
Dim newTerm As terms
newTerm.eoi = Terminator.eoi
newTerm.nChar = Terminator.nChar
newTerm.EightBits = Terminator.EightBits
newTerm.term1 = Terminator.term1
newTerm.term2 = Terminator.term2
Call ConvertTerm(newTerm)
term = Termdll(devHandle%, newTerm, termType%)
End Function


Attribute VB_Name = "WMIConst"
Option Explicit
Public Type BarretteMod
    BBankLabel As String
    BCapacity As Long
    BCaption As String
    BCreationClassName As String
    BDataWidth As String
    BDescription As String
    BDeviceLocator As String
    BFormFactor As String
    BHotSwappable As String
    BInstallDate As String
    BInterleaveDataDepth As String
    BInterleavePosition As String
    BManufacturer As String
    BMemoryType As String
    BModel As String
    BName As String
    BOtherIdentifyingInfo As String
    BPartNumber As String
    BPositionInRow As String
    BPoweredOn As String
    BRemovable As String
    BReplaceable As String
    BSerialNumber As String
    BSKU As String
    BSpeed As String
    BStatus As String
    BTag As String
    BTotalWidth As String
    BTypeDetail As String
    BVersion As String
    ParentAr As String
    ParentSystem As String
End Type
Public Type ArrayMod
    ACaption As String
    ACreationClassName As String
    ADepth As String
    ADescription As String
    AHeight As String
    AHotSwappable As String
    AInstallDate As String
    ALocation As String
    AManufacturer As String
    AMaxCapacity As Long
    AMemoryDevices As Long
    AMemoryErrorCorrection As String
    AModel As String
    AName As String
    AOtherIdentifyingInfo As String
    APartNumber As String
    APoweredOn As String
    ARemovable As String
    AReplaceable As String
    ASerialNumber As String
    ASKU As String
    AStatus As String
    ATag As String
    AUse As String
    AVersion As String
    AWeight As String
    AWidth As String
    TotalCapInstallee As Long
    NbMemoire As Long
    MinMemPosition As Long
    MaxMemPosition As Long
End Type
Public Type SystemMod
    SType As String
    NbMemoire As Long
    TotalCapInstallee As String
    MaxCap As String
    NbArray As Long
End Type
Public Type SlotMod
    DIMCorresp As Long
    ParentAr As String
    ParentSystem As String
End Type

Public Type GeneralInfo
    GModel As String
    GFabriquant As String
    GDescription As String
    Gserial As String
    
End Type
Public Type HardWare
    GModel As String
    Gserial As String
    GUUID As String
    AutoResetCap As String
    AutoResetBoot As String
    BootUpState As String
    HDescrip As String
    FrontPanelReset As String
    Infrared As String
    NetworkServerMode As String
    NbProcesseurs As String
    HMemoire As String
    HStatus As String
    HSystemType As String
    
End Type
Public Type Systeme
    OSName As String
    SPVersion As String
    RegUser As String
    OSSerialN As String
    MemVirtuelle As String
    MemPhysic As String
    SystemDir As String
    Caption  As String
    Domaine As String
    DomaineRole As String
    OwnerName As String
    UserName As String
    TimeZone As String
    AdminPassStatus As String
End Type
Public Type HDDss
    HDescription  As String
    HCaption As String
    HMediaLoaded As String
    HSignature As String
    HInterfaceType As String
    HManufacturer As String
    HMediaType As String
    HModel As String
    HPartitions As String
    HSCSIBus As String
    HSCSILogicalUnit As String
    HSCSIPort As String
    HSCSITargetId As String
    HSize As String
    HSectorsPerTrack As String
    HTotalCylinders As String
    HTotalHeads As String
    HTotalSectors As String
    HTotalTracks As String
    HTracksPerCylinder As String
    HStatus As String
End Type
Public Type Drivess
    HDescription  As String
    HCaption As String
    HFileSystem As String
    HDriveType As String
    HDriveTypeN As Long
    HFreeSpace  As String
    HSize As String
    HVolumeName   As String
    HVolumeSerialNumber   As String
    DLetter As String
End Type
Public Type NetAdapter
    AdapterType  As String
    AdapterTypeID As String
    Description As String
    DeviceID As String
    MACAddress As String
    Manufacturer  As String
    NetConnectionStatus As String
End Type
Public Type PortP
    Description    As String
    PortType  As String
    InternalReferenceDesignator  As String
    ExternalReferenceDesignator  As String
End Type
Public Type Processeur
    AddressWidth     As String
    Architecture   As String
    Caption   As String
    CpuStatus   As String
    CurrentClockSpeed  As String
    CurrentVoltage   As String
    DataWidth  As String
    DeviceID As String
    ExtClock As String
    L2CacheSize As String
    L2CacheSpeed As String
    Level   As String
    Manufacturer As String
    MaxClockSpeed As String
    ProcessorId As String
    Role As String
    SocketDesignation As String
    Version As String
End Type
Public NbSystemMod As Long
Public NbArrayMod As Long
Public NbBarretteMod As Long
Public AllSlotsNm As Long
Public NbProcessor As Long
Public GeneralI As GeneralInfo
Public HardWareI As HardWare
Public SystemeI As Systeme

Public NDiskUn As Long
Public NDiskNoRt As Long
Public NDiskRm As Long
Public NDiskLc As Long
Public NDiskNt As Long
Public NDiskCq As Long
Public NDiskRAM As Long
Public DrvNum As Long
Public NetWorkNm As Long
Public PortNm As Long
Public ProcNum As Long

Public BarretteM(0 To 2000) As BarretteMod
Public ArrayM(0 To 2000) As ArrayMod
Public SystemM(0 To 2000) As SystemMod
Public SlotM(0 To 2000) As SlotMod
Public HDDM(0 To 1000) As HDDss
Public DrivesM(0 To 1000) As Drivess
Public NetAdapterM(0 To 2000) As NetAdapter
Public PortM(0 To 2000) As PortP
Public ProcesseurM(0 To 2000) As Processeur

Public Function NewLineIf(ByVal Str1 As String) As String
    If Str1 <> "" Then NewLineIf = vbNewLine Else NewLineIf = ""
End Function
Public Sub AddToCollect(Collect As Collection, ByVal Str1 As String)
   Dim Ks, existe As Boolean
   existe = False
   For Each Ks In Collect
    If Str1 = Ks Then
        existe = True
        Exit For
    End If
   Next
   If Not existe Then Collect.Add Str1
End Sub
Public Sub ClearInfo()
    Dim i As Long
GeneralI.GDescription = ""
GeneralI.GFabriquant = ""
GeneralI.GModel = ""
GeneralI.Gserial = ""
With HardWareI
    .GModel = ""
    .Gserial = ""
    .GUUID = ""
    .HDescrip = ""
    .HMemoire = ""
    .HSystemType = ""
    .NbProcesseurs = ""
End With
With SystemeI
    .Caption = ""
    .Domaine = ""
    .DomaineRole = ""
    .MemPhysic = ""
    .MemVirtuelle = ""
    .OSName = ""
    .OSSerialN = ""
    .OwnerName = ""
    .OwnerName = ""
    .RegUser = ""
    .UserName = ""
    
End With
For i = 0 To 2000
    BarretteM(i).BBankLabel = ""
    BarretteM(i).BCapacity = 0
    BarretteM(i).BCaption = ""
    BarretteM(i).BCreationClassName = ""
    BarretteM(i).BDataWidth = ""
    BarretteM(i).BDescription = ""
    BarretteM(i).BDeviceLocator = ""
    BarretteM(i).BFormFactor = ""
    BarretteM(i).BHotSwappable = ""
    BarretteM(i).BInstallDate = ""
    BarretteM(i).BInterleaveDataDepth = ""
    BarretteM(i).BInterleavePosition = ""
    BarretteM(i).BManufacturer = ""
    BarretteM(i).BMemoryType = ""
    BarretteM(i).BModel = ""
    BarretteM(i).BName = ""
    BarretteM(i).BOtherIdentifyingInfo = ""
    BarretteM(i).BPartNumber = ""
    BarretteM(i).BPositionInRow = ""
    BarretteM(i).BPoweredOn = ""
    BarretteM(i).BRemovable = ""
    BarretteM(i).BReplaceable = ""
    BarretteM(i).BSerialNumber = ""
    BarretteM(i).BSKU = ""
    BarretteM(i).BSpeed = ""
    BarretteM(i).BStatus = ""
    BarretteM(i).BTag = ""
    BarretteM(i).BTotalWidth = ""
    BarretteM(i).BTypeDetail = ""
    BarretteM(i).BVersion = ""
    BarretteM(i).ParentAr = ""
    BarretteM(i).ParentSystem = ""

    ArrayM(i).ACaption = ""
    ArrayM(i).ACreationClassName = ""
    ArrayM(i).ADepth = ""
    ArrayM(i).ADescription = ""
    ArrayM(i).AHeight = ""
    ArrayM(i).AHotSwappable = ""
    ArrayM(i).AInstallDate = ""
    ArrayM(i).ALocation = ""
    ArrayM(i).AManufacturer = ""
    ArrayM(i).AMaxCapacity = 0
    ArrayM(i).AMemoryDevices = 0
    ArrayM(i).AMemoryErrorCorrection = ""
    ArrayM(i).AModel = ""
    ArrayM(i).AName = ""
    ArrayM(i).AOtherIdentifyingInfo = ""
    ArrayM(i).APartNumber = ""
    ArrayM(i).APoweredOn = ""
    ArrayM(i).ARemovable = ""
    ArrayM(i).AReplaceable = ""
    ArrayM(i).ASerialNumber = ""
    ArrayM(i).ASKU = ""
    ArrayM(i).AStatus = ""
    ArrayM(i).ATag = ""
    ArrayM(i).AUse = ""
    ArrayM(i).AVersion = ""
    ArrayM(i).AWeight = ""
    ArrayM(i).AWidth = ""
    ArrayM(i).TotalCapInstallee = 0
    ArrayM(i).NbMemoire = 0
    ArrayM(i).MinMemPosition = -1
    ArrayM(i).MaxMemPosition = -1
    
    
    SystemM(i).SType = ""
    SystemM(i).NbMemoire = 0
    SystemM(i).TotalCapInstallee = ""
    SystemM(i).MaxCap = ""
    SystemM(i).NbArray = 0
    
    SlotM(i).DIMCorresp = -1
    SlotM(i).ParentAr = ""
    SlotM(i).ParentSystem = ""
Next i
End Sub
Public Sub Clear_Info1000()
Dim i As Long
For i = 0 To 1000
    HDDM(i).HCaption = ""
    HDDM(i).HDescription = ""
    HDDM(i).HInterfaceType = ""
    HDDM(i).HManufacturer = ""
    HDDM(i).HMediaLoaded = ""
    HDDM(i).HMediaType = ""
    HDDM(i).HModel = ""
    HDDM(i).HPartitions = ""
    HDDM(i).HSignature = ""
    HDDM(i).HSize = ""
    HDDM(i).HStatus = ""
    HDDM(i).HTotalCylinders = ""
    HDDM(i).HTotalHeads = ""
    HDDM(i).HTotalSectors = ""
    HDDM(i).HTotalTracks = ""
    HDDM(i).HTracksPerCylinder = ""
    With DrivesM(i)
        .HCaption = ""
        .HDescription = ""
        .HDriveType = ""
        .HFileSystem = ""
        .HFreeSpace = ""
        .HSize = ""
        .HVolumeName = ""
        .HVolumeSerialNumber = ""
    End With
Next i
End Sub


Public Function TransProcMz(ByVal ProcVal As Long) As String
Dim Sk As Single
  Select Case ProcVal
    Case Is < 1024
        TransProcMz = CStr(ProcVal) + " MHz"
    Case Is >= 1024
        Sk = CSng(ProcVal / 1024)
        TransProcMz = CStr(Sk) + " GHz"
  End Select
End Function
Public Function Type_NetConnectionStatus(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_NetConnectionStatus = "Disconnected "
Case 1
Type_NetConnectionStatus = "Connecting "
Case 2
Type_NetConnectionStatus = "Connected "
Case 3
Type_NetConnectionStatus = "Disconnecting "
Case 4
Type_NetConnectionStatus = "Hardware not present "
Case 5
Type_NetConnectionStatus = "Hardware disabled "
Case 6
Type_NetConnectionStatus = "Hardware malfunction "
Case 7
Type_NetConnectionStatus = "Media disconnected "
Case 8
Type_NetConnectionStatus = "Authenticating "
Case 9
Type_NetConnectionStatus = "Authentication succeeded "
Case 10
Type_NetConnectionStatus = "Authentication failed"
Case 11
Type_NetConnectionStatus = "Invalid address "
Case 12
Type_NetConnectionStatus = "Credentials required "

End Select
End Function
Public Function Type_AdapterTypeID(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_AdapterTypeID = "Ethernet 802.3 "
Case 1
Type_AdapterTypeID = "Token Ring 802.5 "
Case 2
Type_AdapterTypeID = "Fiber Distributed Data Interface (FDDI) "
Case 3
Type_AdapterTypeID = "Wide Area Network (WAN) "
Case 4
Type_AdapterTypeID = "LocalTalk "
Case 5
Type_AdapterTypeID = "Ethernet using DIX header format "
Case 6
Type_AdapterTypeID = "ARCNET "
Case 7
Type_AdapterTypeID = "ARCNET (878.2) "
Case 8
Type_AdapterTypeID = "ATM "
Case 9
Type_AdapterTypeID = "Wireless "
Case 10
Type_AdapterTypeID = "Infrared Wireless "
Case 11
Type_AdapterTypeID = "Bpc "
Case 12
Type_AdapterTypeID = "CoWan "
Case 13
Type_AdapterTypeID = "1394 "

End Select
End Function
Public Function Type_DriveType(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_DriveType = "Unknown"
Case 1
Type_DriveType = "No Root Directory"
Case 2
Type_DriveType = "Removable Disk"
Case 3
Type_DriveType = "Local Disk"
Case 4
Type_DriveType = "Network Drive"
Case 5
Type_DriveType = "Compact Disc"
Case 6
Type_DriveType = "RAM Disk"

End Select
End Function
Public Function Type_Architecture(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_Architecture = "x86"
Case 1
Type_Architecture = "MIPS"
Case 2
Type_Architecture = "Alpha"
Case 3
Type_Architecture = "PowerPC"
Case 6
Type_Architecture = "Intel Itanium Processor Family (IPF)"
Case 9
Type_Architecture = "x64"
End Select
End Function
Public Function Type_CpuStatus(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_CpuStatus = "Unknown"
Case 1
Type_CpuStatus = "CPU Enabled"
Case 2
Type_CpuStatus = "CPU Disabled by User via BIOS Setup"
Case 3
Type_CpuStatus = "CPU Disabled By BIOS (POST Error)"
Case 4
Type_CpuStatus = "CPU is Idle"
Case 5
Type_CpuStatus = "Reserved"
Case 6
Type_CpuStatus = "Reserved"
Case 7
Type_CpuStatus = "Other"
End Select
End Function
Public Function Type_AdminPasswordStatus(ByVal TypeN As Long) As String
Select Case TypeN
Case 1
Type_AdminPasswordStatus = "Disabled"
Case 2
Type_AdminPasswordStatus = "Enabled"
Case 3
Type_AdminPasswordStatus = "Not Implemented"
Case 4
Type_AdminPasswordStatus = "Unknown"
End Select
End Function
Public Function Type_PortType(ByVal TypeN As Long) As String
Select Case TypeN

Case 0
Type_PortType = "None"
Case 1
Type_PortType = "Parallel Port XT/AT Compatible"
Case 2
Type_PortType = "Parallel Port PS/2"
Case 3
Type_PortType = "Parallel Port ECP"
Case 4
Type_PortType = "Parallel Port EPP"
Case 5
Type_PortType = "Parallel Port ECP/EPP"
Case 6
Type_PortType = "Serial Port XT/AT Compatible"
Case 7
Type_PortType = "Serial Port 16450 Compatible"
Case 8
Type_PortType = "Serial Port 16550 Compatible"
Case 9
Type_PortType = "Serial Port 16550A Compatible"
Case 10
Type_PortType = "SCSI Port"
Case 11
Type_PortType = "MIDI Port"
Case 12
Type_PortType = "Joy Stick Port"
Case 13
Type_PortType = "Keyboard Port"
Case 14
Type_PortType = "Mouse Port"
Case 15
Type_PortType = "SSA SCSI"
Case 16
Type_PortType = "USB"
Case 17
Type_PortType = "FireWire (IEEE P1394)"
Case 18
Type_PortType = "PCMCIA Type II"
Case 19
Type_PortType = "PCMCIA Type II"
Case 20
Type_PortType = "PCMCIA Type III"
Case 21
Type_PortType = "CardBus"
Case 22
Type_PortType = "Access Bus Port"
Case 23
Type_PortType = "SCSI II"
Case 24
Type_PortType = "SCSI Wide"
Case 25
Type_PortType = "PC-98"
Case 26
Type_PortType = "PC-98-Hireso"
Case 27
Type_PortType = "PC-H98"
Case 28
Type_PortType = "Video Port"
Case 29
Type_PortType = "Audio Port"
Case 30
Type_PortType = "Modem Port"
Case 31
Type_PortType = "Network Port"
Case 32
Type_PortType = "8251 Compatible"
Case 33
Type_PortType = "8251 FIFO Compatible"
End Select
End Function

Public Function Type_ProtocolSupported(ByVal TypeN As Long) As String
Select Case TypeN
Case 1
Type_ProtocolSupported = "Other "
Case 2
Type_ProtocolSupported = "Unknown "
Case 3
Type_ProtocolSupported = "EISA "
Case 4
Type_ProtocolSupported = "ISA "
Case 5
Type_ProtocolSupported = "PCI "
Case 6
Type_ProtocolSupported = "ATA/ATAPI "
Case 7
Type_ProtocolSupported = "Flexible Diskette "
Case 8
Type_ProtocolSupported = "1496 "
Case 9
Type_ProtocolSupported = "SCSI Parallel Interface "
Case 10
Type_ProtocolSupported = "SCSI Fibre Channel Protocol "
Case 11
Type_ProtocolSupported = "SCSI Serial Bus Protocol "
Case 12
Type_ProtocolSupported = "SCSI Serial Bus Protocol-2 (1394) "
Case 13
Type_ProtocolSupported = "SCSI Serial Storage Architecture "
Case 14
Type_ProtocolSupported = "VESA "
Case 15
Type_ProtocolSupported = "PCMCIA "
Case 16
Type_ProtocolSupported = "Universal Serial Bus "
Case 17
Type_ProtocolSupported = "Parallel Protocol "
Case 18
Type_ProtocolSupported = "ESCON "
Case 19
Type_ProtocolSupported = "Diagnostic "
Case 20
Type_ProtocolSupported = "I2C "
Case 21
Type_ProtocolSupported = "Power "
Case 22
Type_ProtocolSupported = "HIPPI "
Case 23
Type_ProtocolSupported = "MultiBus "
Case 24
Type_ProtocolSupported = "VME "
Case 25
Type_ProtocolSupported = "IPI "
Case 26
Type_ProtocolSupported = "IEEE-488 "
Case 27
Type_ProtocolSupported = "RS232 "
Case 28
Type_ProtocolSupported = "IEEE 802.3 10BASE5 "
Case 29
Type_ProtocolSupported = "IEEE 802.3 10BASE2 "
Case 30
Type_ProtocolSupported = "IEEE 802.3 1BASE5 "
Case 31
Type_ProtocolSupported = "IEEE 802.3 10BROAD36 "
Case 32
Type_ProtocolSupported = "IEEE 802.3 100BASEVG "
Case 33
Type_ProtocolSupported = "IEEE 802.5 Token-Ring "
Case 34
Type_ProtocolSupported = "ANSI X3T9.5 FDDI "
Case 35
Type_ProtocolSupported = "MCA "
Case 36
Type_ProtocolSupported = "ESDI "
Case 37
Type_ProtocolSupported = "IDE "
Case 38
Type_ProtocolSupported = "CMD "
Case 39
Type_ProtocolSupported = "ST506 "
Case 40
Type_ProtocolSupported = "DSSI "
Case 41
Type_ProtocolSupported = "QIC2 "
Case 42
Type_ProtocolSupported = "Enhanced ATA/IDE "
Case 43
Type_ProtocolSupported = "AGP "
Case 44
Type_ProtocolSupported = "TWIRP (two-way infrared) "
Case 45
Type_ProtocolSupported = "FIR (fast infrared) "
Case 46
Type_ProtocolSupported = "SIR (serial infrared) "
Case 47
Type_ProtocolSupported = "IrBus "

End Select
End Function
Public Function TransMemK(ByVal MemVal As Long) As String
Dim Sk As Single
On Error Resume Next
  Select Case MemVal
    Case Is < 1024
        TransMemK = CStr(MemVal) + " KB"
    Case 1024 To 1048575
        Sk = CSng(MemVal / 1024)
        TransMemK = CStr(Sk) + " MB"
    Case 1048576 To 1073741823
        Sk = CSng(MemVal / 1048576)
        TransMemK = CStr(Sk) + " GB"
    Case Else
        TransMemK = CStr(MemVal)
  End Select
  
End Function
Public Function Type_MemoryType(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_MemoryType = "Unknown"
Case 1
Type_MemoryType = "Other"
Case 2
Type_MemoryType = " DRAM"
Case 3
 Type_MemoryType = "Synchronous DRAM"
Case 4
 Type_MemoryType = "Cache DRAM"
Case 5
 Type_MemoryType = "EDO"
Case 6
 Type_MemoryType = "EDRAM"
Case 7
 Type_MemoryType = "VRAM"
Case 8
 Type_MemoryType = "SRAM"
Case 9
 Type_MemoryType = "RAM"
Case 10
 Type_MemoryType = "ROM"
Case 11
Type_MemoryType = "Flash"
Case 12
 Type_MemoryType = "EEPROM"
Case 13
 Type_MemoryType = "FEPROM"
Case 14
Type_MemoryType = "EPROM"
Case 15
 Type_MemoryType = "CDRAM"
Case 16
Type_MemoryType = "3DRAM"
Case 17
 Type_MemoryType = "SDRAM"
Case 18
Type_MemoryType = "SGRAM"
Case 19
 Type_MemoryType = "RDRAM"
Case 20
 Type_MemoryType = "DDR"

End Select
End Function
Public Function Type_FrontPanelResetStatus(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_FrontPanelResetStatus = "Disabled"
Case 1
Type_FrontPanelResetStatus = "Enabled"
Case 2
Type_FrontPanelResetStatus = "Not Implemented"
Case 3
Type_FrontPanelResetStatus = "Unknown"
End Select
End Function
Public Function Type_DomainRole(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_DomainRole = "Standalone Workstation"
Case 1
Type_DomainRole = "Member Workstation"
Case 2
Type_DomainRole = "Standalone Server"
Case 3
Type_DomainRole = "Member Server"
Case 4
Type_DomainRole = "Backup Domain Controller"
Case 5
Type_DomainRole = "Primary Domain Controller"
End Select
End Function
Public Function Type_InterleavePosition(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_InterleavePosition = "Non -interleaved"
Case 1
Type_InterleavePosition = "First position"
Case 2
Type_InterleavePosition = "Second position"
Case 3
Type_InterleavePosition = "third position"
Case 4
Type_InterleavePosition = "fourth position"
End Select
End Function
Public Function Type_Location(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_Location = "Reserved"
Case 1
Type_Location = "Other"
Case 2
Type_Location = "Unknown"
Case 3
Type_Location = "System board Or motherboard"
Case 4
Type_Location = "ISA add-on card"
Case 5
Type_Location = "EISA add-on card"
Case 6
Type_Location = "PCI add-on card"
Case 7
Type_Location = "MCA add-on card"
Case 8
Type_Location = "PCMCIA add-on card"
Case 9
Type_Location = "Proprietary add-on card"
Case 10
Type_Location = "NuBus"
Case 11
Type_Location = "PC-98/C20 add-on card"
Case 12
Type_Location = "PC-98/C24 add-on card"
Case 13
Type_Location = "PC-98/E add-on card"
Case 14
Type_Location = "PC-98/Local bus add-on card"

End Select
End Function
Public Function Type_Use(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_Use = "Reserved"
Case 1
Type_Use = "Other"
Case 2
Type_Use = "Unknown"
Case 3
Type_Use = "System memory"
Case 4
Type_Use = "Video memory"
Case 5
Type_Use = "Flash memory"
Case 6
Type_Use = "Non-volatile RAM"
Case 7
Type_Use = "Cache memory"

End Select
End Function
Public Function Type_MemoryErrorCorrection(ByVal TypeN As Long) As String
Select Case TypeN
Case 0
Type_MemoryErrorCorrection = "Reserved"
Case 1
Type_MemoryErrorCorrection = "Other"
Case 2
Type_MemoryErrorCorrection = "Unknown"
Case 3
Type_MemoryErrorCorrection = "None"
Case 4
Type_MemoryErrorCorrection = "Parity"
Case 5
Type_MemoryErrorCorrection = "Single-bit ECC"
Case 6
Type_MemoryErrorCorrection = "Multi-bit ECC"
Case 7
Type_MemoryErrorCorrection = "CRC"
End Select
End Function
Public Function Type_TypeDetail(ByVal TypeN As Long) As String
Select Case TypeN
Case 1
Type_TypeDetail = " Reserved"
Case 2
Type_TypeDetail = "Other"
Case 4
Type_TypeDetail = "Unknown"
Case 8
Type_TypeDetail = "Fast -paged"
Case 16
Type_TypeDetail = "Static column"
Case 32
Type_TypeDetail = "Pseudo-static"
Case 64
Type_TypeDetail = "RAMBUS"
Case 128
Type_TypeDetail = "Synchronous"
Case 256
Type_TypeDetail = "CMOS"
Case 512
Type_TypeDetail = "EDO"
Case 1024
Type_TypeDetail = "Window DRAM"
Case 2048
Type_TypeDetail = "Cache DRAM"
Case 4096
Type_TypeDetail = "Non -volatile"


End Select
End Function
Public Function Type_FormFactor(ByVal TypeN As Long) As String
  Select Case TypeN
Case 0
    Type_FormFactor = "Unknown"
Case 1
    Type_FormFactor = "Other"
Case 2
    Type_FormFactor = "SIP"
Case 3
    Type_FormFactor = "DIP"
Case 4
    Type_FormFactor = "ZIP"
Case 5
    Type_FormFactor = "SOJ"
Case 6
    Type_FormFactor = "Proprietary"
Case 7
    Type_FormFactor = "SIMM"
Case 8
    Type_FormFactor = "DIMM"
Case 9
    Type_FormFactor = "TSOP"
Case 10
    Type_FormFactor = "PGA"
Case 11
    Type_FormFactor = "RIMM"
Case 12
    Type_FormFactor = "SODIMM"
Case 13
    Type_FormFactor = "SRIMM"
Case 14
    Type_FormFactor = "SMD"
Case 15
    Type_FormFactor = "SSMP"
Case 16
    Type_FormFactor = "QFP"
Case 17
    Type_FormFactor = "TQFP"
Case 18
    Type_FormFactor = "SOIC"
Case 19
    Type_FormFactor = "LCC"
Case 20
    Type_FormFactor = "PLCC"
Case 21
    Type_FormFactor = "BGA"
Case 22
    Type_FormFactor = "FPBGA"
Case 23
    Type_FormFactor = "LGA"
  End Select
End Function

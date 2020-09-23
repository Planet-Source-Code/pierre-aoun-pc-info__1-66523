VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PC Info"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picEnCrs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   -20000
      ScaleHeight     =   1665
      ScaleWidth      =   6465
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "In progress..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.ListBox lstMain 
      Height          =   3570
      ItemData        =   "Form1.frx":0742
      Left            =   120
      List            =   "Form1.frx":0761
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   2655
   End
   Begin VB.ListBox lstSlots 
      Height          =   3570
      Left            =   8400
      TabIndex        =   9
      Top             =   960
      Width           =   2775
   End
   Begin VB.ListBox lstArray 
      Height          =   3570
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Width           =   2655
   End
   Begin VB.ListBox lstSystem 
      Height          =   3570
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtMainDesc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4560
      Width           =   10215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      Top             =   90
      Width           =   2055
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   9720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pass:"
      Height          =   195
      Left            =   6120
      TabIndex        =   5
      Top             =   480
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User:"
      Height          =   195
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PC IP or Host Name:"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PC info
'Author : Pierre AOUN
'email: pierre_aoun@hotmail.com
'The original version is French
'-----------------------------------------
Option Explicit
Private Sub ClearAllF()
lstSystem.Clear
lstArray.Clear
lstSlots.Clear
txtMainDesc = ""
End Sub
Private Sub Gettest()
Dim strObject
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from  Win32_PortConnector", , 48)
  
For Each objItem In colItems
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Description  : " & CStr(objItem.Description)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Caption: " & CStr(objItem.Caption)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Name: " & CStr(objItem.Name)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "PortType : " & Type_PortType(objItem.PortType)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Tag  : " & CStr(objItem.Tag)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "InternalReferenceDesignator : " & CStr(objItem.InternalReferenceDesignator)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "ExternalReferenceDesignator : " & CStr(objItem.ExternalReferenceDesignator)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "ConnectorType  : " & CStr(objItem.ConnectorType)
  
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "---------------------------------------------"

Next
End Sub
Private Sub GetPorts1()
Dim strObject
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from  Win32_PortConnector", , 48)
  
For Each objItem In colItems
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Description  : " & CStr(objItem.Description)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Caption: " & CStr(objItem.Caption)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Name: " & CStr(objItem.Name)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "PortType : " & Type_PortType(objItem.PortType)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Tag  : " & CStr(objItem.Tag)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "InternalReferenceDesignator : " & CStr(objItem.InternalReferenceDesignator)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "ExternalReferenceDesignator : " & CStr(objItem.ExternalReferenceDesignator)
    
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "---------------------------------------------"

Next
End Sub
Private Sub GetDrives()
    On Error Resume Next
    Clear_Info1000
    
    Dim strObject, K As String, i As Long, j As Long
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from  Win32_LogicalDisk", , 48)
NDiskUn = 0
NDiskNoRt = 0
NDiskRm = 0
NDiskLc = 0
NDiskNt = 0
NDiskCq = 0
NDiskRAM = 0
DrvNum = 0
lstArray.Clear
lstSystem.Clear
For Each objItem In colItems
    DrvNum = DrvNum + 1
  With DrivesM(DrvNum)
    .HDescription = "Description : " & CStr(objItem.Description)
    .HCaption = "Caption : " & CStr(objItem.Caption)
    .HFileSystem = "File System : " & CStr(objItem.FileSystem)
    .HDriveType = "DriveType: " & Type_DriveType(objItem.DriveType)
    .HDriveTypeN = objItem.DriveType
    .HFreeSpace = "Free Space : " & TransMemK(objItem.FreeSpace / 1024)
    .HSize = "Capacity : " & TransMemK(objItem.Size / 1024)
    .HVolumeName = "Volume name : " & CStr(objItem.VolumeName)
    .HVolumeSerialNumber = "Serial Number : " & CStr(objItem.VolumeSerialNumber)
    .DLetter = CStr(objItem.Caption)
K = CStr(objItem.Caption)
K = K + " " + CStr(objItem.VolumeName)
K = K + "(" + TransMemK(objItem.FreeSpace / 1024) + ")"
lstArray.AddItem K
Select Case .HDriveTypeN
Case 0
NDiskUn = NDiskUn + 1
Case 1
NDiskNoRt = NDiskNoRt + 1
Case 2
NDiskRm = NDiskRm + 1
Case 3
NDiskLc = NDiskLc + 1
Case 4
NDiskNt = NDiskNt + 1
Case 5
NDiskCq = NDiskCq + 1
Case 6
NDiskRAM = NDiskRAM + 1
End Select
End With
Next

 For j = 0 To 6
  For i = 1 To DrvNum
  If DrivesM(i).HDriveTypeN = j Then
    Select Case j
    Case 0
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskUn) + ")"
        Exit For
    Case 1
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskNoRt) + ")"
        Exit For
    Case 2
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskRm) + ")"
        Exit For
    Case 3
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskLc) + ")"
        Exit For
    Case 4
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskNt) + ")"
        Exit For
    Case 5
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskCq) + ")"
        Exit For
    Case 6
        lstSystem.AddItem Type_DriveType(j) + " (" + CStr(NDiskRAM) + ")"
        Exit For
    End Select
  End If
 Next i
Next j
If NDiskUn > 0 Then txtMainDesc = txtMainDesc + "Number of unknown Drives: " + CStr(NDiskUn) + vbNewLine
If NDiskNoRt > 0 Then txtMainDesc = txtMainDesc + "Numbre of Volumes with no root: " + CStr(NDiskNoRt) + vbNewLine
If NDiskRm > 0 Then txtMainDesc = txtMainDesc + "Number of Removable Drives: " + CStr(NDiskRm) + vbNewLine
If NDiskLc > 0 Then txtMainDesc = txtMainDesc + "Number of Local Drives: " + CStr(NDiskLc) + vbNewLine
If NDiskNt > 0 Then txtMainDesc = txtMainDesc + "Number of Network Drives: " + CStr(NDiskNt) + vbNewLine
If NDiskCq > 0 Then txtMainDesc = txtMainDesc + "Number of compact disk Drives: " + CStr(NDiskCq) + vbNewLine
If NDiskRAM > 0 Then txtMainDesc = txtMainDesc + "Number of RAM Drives: " + CStr(NDiskRAM) + vbNewLine


End Sub
Private Sub GetToutMateriel()
Dim strObject
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from  Win32_PnPEntity", , 48)
  
For Each objItem In colItems
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Description  : " & CStr(objItem.Description)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Caption: " & CStr(objItem.Caption)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "ClassGuid : " & CStr(objItem.ClassGuid)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Name: " & CStr(objItem.Name)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "DeviceID : " & CStr(objItem.DeviceID)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Manufacturer : " & CStr(objItem.Manufacturer)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Service : " & CStr(objItem.Service)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Status : " & CStr(objItem.Status)
    
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "---------------------------------------------"

Next
End Sub
Private Sub GetXPort(ByVal Xport As String)
Dim i As Long
Dim Numb As Long
txtMainDesc.Text = ""
Numb = 0
For i = 1 To PortNm
With PortM(i)
If .PortType = Xport Then
    Numb = Numb + 1
    txtMainDesc.Text = txtMainDesc.Text + "Description  : " & .Description + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + "PortType : " + .PortType + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + "Internal Reference : " + .InternalReferenceDesignator + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + "External Reference : " + .ExternalReferenceDesignator + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "---------------------------------------------" + vbNewLine
End If
End With
Next i
txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine _
   + Xport + " = " + CStr(Numb) + " :" + vbNewLine + vbNewLine + txtMainDesc.Text

End Sub
Private Sub GetSerie()
Dim strObject
Dim Numb As Long
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort", , 48)
  Numb = 0
  txtMainDesc.Text = ""
For Each objItem In colItems
    Numb = Numb + 1
    txtMainDesc.Text = txtMainDesc.Text + "Description : " & CStr(objItem.Description) + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + "Caption: " & CStr(objItem.Caption) + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + "ProviderType    : " & CStr(objItem.ProviderType) + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + "Status     : " & CStr(objItem.Status) + vbNewLine
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "---------------------------------------------" + vbNewLine
Next
txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine _
   + "Nombre de Ports Serie = " + CStr(Numb) + " :" + vbNewLine + vbNewLine + txtMainDesc.Text

End Sub
Private Sub GetPorts()
Dim strObject
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim P As Collection, Ks, i As Long
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PortConnector", , 48)
  PortNm = 0
  lstSystem.Clear
  Set P = New Collection
For Each objItem In colItems
   If objItem.PortType <> 0 Then
    PortNm = PortNm + 1
    With PortM(PortNm)
    .Description = objItem.Description
    .PortType = Type_PortType(objItem.PortType)
    .ExternalReferenceDesignator = objItem.ExternalReferenceDesignator
    .InternalReferenceDesignator = objItem.InternalReferenceDesignator
    AddToCollect P, .PortType
    End With
   End If
Next
For Each Ks In P
    lstSystem.AddItem Ks
    txtMainDesc.Text = txtMainDesc.Text + Ks + vbNewLine
Next
End Sub
Private Sub GetParallel()
Dim strObject
Dim Numb As Long
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ParallelPort", , 48)
  txtMainDesc.Text = ""
  Numb = 0
For Each objItem In colItems
    Numb = Numb + 1
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "Description : " & CStr(objItem.Description)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "ProtocolSupported : " & Type_ProtocolSupported(objItem.ProtocolSupported)
    txtMainDesc.Text = txtMainDesc.Text + vbNewLine + "---------------------------------------------"
Next
txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine _
   + "Nombre de Ports Parallele = " + CStr(Numb) + " :" + vbNewLine + vbNewLine + txtMainDesc.Text

End Sub
Private Sub GetDisques()
   On Error Resume Next
    Dim DskNum As Long
    Clear_Info1000
    Dim strObject
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive", , 48)
lstSystem.Clear
DskNum = 0
For Each objItem In colItems
    DskNum = DskNum + 1
    With HDDM(DskNum)
    .HDescription = "Description : " & CStr(objItem.Description)
    .HCaption = "Caption : " & CStr(objItem.Caption)
    .HInterfaceType = "Interface : " & CStr(objItem.InterfaceType)
    .HManufacturer = "manufacturer : " & CStr(objItem.Manufacturer)
    .HMediaLoaded = "installed Media : " & CStr(objItem.MediaLoaded)
    .HMediaType = "Media Type: " & CStr(objItem.MediaType)
    .HModel = "Model : " & CStr(objItem.Model)
    .HPartitions = "Number of Partitions : " & CStr(objItem.Partitions)
    .HSCSIBus = "Bus : " & CStr(objItem.SCSIBus)
    .HSCSILogicalUnit = "Logical Unit : " & CStr(objItem.SCSILogicalUnit)
    .HSCSIPort = "Port : " & CStr(objItem.SCSIPort)
    .HSCSITargetId = "TargetId : " & CStr(objItem.SCSITargetId)
    .HSectorsPerTrack = "Sectors per Track : " & CStr(objItem.SectorsPerTrack)
    .HSignature = "Signature: " & CStr(objItem.Signature)
    .HSize = "Capacity : " & TransMemK(objItem.Size / 1024)
    .HStatus = "Status : " & CStr(objItem.Status)
    .HTotalCylinders = "Cylinders : " & CStr(objItem.TotalCylinders)
    .HTotalHeads = "Heads : " & CStr(objItem.TotalHeads)
    .HTotalSectors = "Sectors : " & CStr(objItem.TotalSectors)
    .HTotalTracks = "Tracks : " & CStr(objItem.TotalTracks)
    .HTracksPerCylinder = "Tracks per Cylindre : " & CStr(objItem.TracksPerCylinder)
    lstSystem.AddItem CStr(objItem.Caption) + " (" + TransMemK(objItem.Size / 1024) + ")"
    End With
Next
txtMainDesc = txtMainDesc + "Nomber of Physical volumes : " + CStr(DskNum)
End Sub
Private Sub GetNetwork()
    Dim strObject, K As String
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
  NetWorkNm = 0
  lstSystem.Clear
For Each objItem In colItems
      NetWorkNm = NetWorkNm + 1
      With NetAdapterM(NetWorkNm)
     .AdapterType = CStr(objItem.AdapterType)
     .AdapterTypeID = Type_AdapterTypeID(objItem.AdapterTypeID)
     .Description = CStr(objItem.Description)
     .DeviceID = CStr(objItem.DeviceID)
     .MACAddress = CStr(objItem.MACAddress)
     .Manufacturer = CStr(objItem.Manufacturer)
     .NetConnectionStatus = Type_NetConnectionStatus(objItem.NetConnectionStatus)
     lstSystem.AddItem .Description
     End With
Next
End Sub
Private Sub GetProcessorInfo()
Dim strObject
    On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
    ProcNum = 0
    
For Each objItem In colItems
    ProcNum = ProcNum + 1
    With ProcesseurM(ProcNum)
    .AddressWidth = "AddressWidth : " & CStr(objItem.AddressWidth) + " Bits"
    .Architecture = "Architecture : " & Type_Architecture(objItem.Architecture)
    .Caption = "Caption : " & objItem.Caption
    .CpuStatus = "Status : " & Type_CpuStatus(objItem.CpuStatus)
    .CurrentClockSpeed = "Current Clock Speed : " & TransProcMz(objItem.CurrentClockSpeed)
    .CurrentVoltage = "Voltage : " & CStr(objItem.CurrentVoltage / 10) + " V"
    .DataWidth = "Data Width  : " & CStr(objItem.DataWidth) + " Bits"
    .DeviceID = "DeviceID: " & CStr(objItem.DeviceID)
    .ExtClock = "Externel Clock : " & TransProcMz(objItem.ExtClock)
    .L2CacheSize = "L2 Cache capacity: " & CStr(objItem.L2CacheSize) + " KB"
    .L2CacheSpeed = "L2 Cache Speed: " & TransProcMz(objItem.L2CacheSpeed)
    .Level = "Lecvel : " & CStr(objItem.Level)
    .Manufacturer = "manufacturer : " & CStr(objItem.Manufacturer)
    .MaxClockSpeed = "Maximum Clock Speed : " & TransProcMz(objItem.MaxClockSpeed)
    .ProcessorId = "Processeur Id: " & CStr(objItem.ProcessorId)
    .Role = "Role : " & CStr(objItem.Role)
    .SocketDesignation = "Socket : " & CStr(objItem.SocketDesignation)
    .Version = "Version : " & CStr(objItem.Version)
    lstSystem.AddItem CStr(objItem.DeviceID) + " (" + TransProcMz(objItem.CurrentClockSpeed) + ")"
    End With
Next
txtMainDesc.Text = txtMainDesc.Text + "Number of processors = " + CStr(ProcNum)
End Sub
Private Function GetMemNumero(ByVal MemTag As String) As Long
GetMemNumero = CLng(Right(MemTag, Len(MemTag) - Len("Physical Memory")))

End Function
Private Function GetArNumero(ByVal ArTag As String) As Long
GetArNumero = CLng(Right(ArTag, Len(ArTag) - Len("Physical Memory Array")))
End Function
Private Function GetSlotInf(ByVal SlotN As Long) As String
Dim i As Long
On Error Resume Next
GetSlotInf = "Slot " + Format(SlotN, "00") + " (Empty)"
For i = 1 To NbBarretteMod
    If GetMemNumero(BarretteM(i).BTag) = SlotN Then
      GetSlotInf = "Slot " + Format(SlotN, "00") + ": " + BarretteM(i).BTag + " (" + TransMemK(BarretteM(i).BCapacity) + ")"
      Exit For
    End If
Next i
End Function
Private Function GetDIM(ByVal SlotN As Long) As Long
Dim i As Long
On Error Resume Next
GetDIM = -1
For i = 1 To NbBarretteMod
    If GetMemNumero(BarretteM(i).BTag) = SlotN Then
      GetDIM = i
      Exit For
    End If
Next i
End Function
Private Sub GetMemory()
Dim strObject
On Error Resume Next
    Call ClearInfo
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    Dim NbAlea As Long, i As Long
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    'get Array
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray", , 48)
    NbAlea = 0
    For Each objItem In colItems
    NbAlea = NbAlea + 1
    ArrayM(NbAlea).ACaption = CStr(objItem.Caption)
    ArrayM(NbAlea).ACreationClassName = CStr(objItem.CreationClassName)
    ArrayM(NbAlea).ADepth = CStr(objItem.Depth)
    ArrayM(NbAlea).ADescription = CStr(objItem.Description)
    ArrayM(NbAlea).AHeight = CStr(objItem.Height)
    ArrayM(NbAlea).AHotSwappable = CStr(objItem.HotSwappable)
    ArrayM(NbAlea).AInstallDate = CStr(objItem.InstallDate)
    ArrayM(NbAlea).ALocation = Type_Location(objItem.Location)
    ArrayM(NbAlea).AManufacturer = CStr(objItem.Manufacturer)
    ArrayM(NbAlea).AMaxCapacity = CLng(objItem.MaxCapacity)
    ArrayM(NbAlea).AMemoryDevices = CLng(objItem.MemoryDevices)
    ArrayM(NbAlea).AMemoryErrorCorrection = Type_MemoryErrorCorrection(objItem.MemoryErrorCorrection)
    ArrayM(NbAlea).AModel = CStr(objItem.Model)
    ArrayM(NbAlea).AName = CStr(objItem.Name)
    ArrayM(NbAlea).AOtherIdentifyingInfo = CStr(objItem.OtherIdentifyingInfo)
    ArrayM(NbAlea).APartNumber = CStr(objItem.PartNumber)
    ArrayM(NbAlea).APoweredOn = CStr(objItem.PoweredOn)
    ArrayM(NbAlea).ARemovable = CStr(objItem.Removable)
    ArrayM(NbAlea).AReplaceable = CStr(objItem.Replaceable)
    ArrayM(NbAlea).ASerialNumber = CStr(objItem.SerialNumber)
    ArrayM(NbAlea).ASKU = CStr(objItem.SKU)
    ArrayM(NbAlea).AStatus = CStr(objItem.Status)
    ArrayM(NbAlea).ATag = CStr(objItem.Tag)
    ArrayM(NbAlea).AUse = Type_Use(objItem.Use)
    ArrayM(NbAlea).AVersion = CStr(objItem.Version)
    ArrayM(NbAlea).AWeight = CStr(objItem.Weight)
    ArrayM(NbAlea).AWidth = CStr(objItem.Width)
Next
NbArrayMod = NbAlea
'Get Memories
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory", , 48)
    NbAlea = 0
    For Each objItem In colItems
    NbAlea = NbAlea + 1
    BarretteM(NbAlea).BBankLabel = CStr(objItem.BankLabel)
    BarretteM(NbAlea).BCapacity = CLng(objItem.Capacity) / 1024
    BarretteM(NbAlea).BCaption = CStr(objItem.Caption)
    BarretteM(NbAlea).BCreationClassName = CStr(objItem.CreationClassName)
    BarretteM(NbAlea).BDataWidth = CStr(objItem.DataWidth)
    BarretteM(NbAlea).BDescription = CStr(objItem.Description)
    BarretteM(NbAlea).BDeviceLocator = CStr(objItem.DeviceLocator)
    BarretteM(NbAlea).BFormFactor = Type_FormFactor(objItem.FormFactor)
    BarretteM(NbAlea).BHotSwappable = CStr(objItem.HotSwappable)
    BarretteM(NbAlea).BInstallDate = CStr(objItem.InstallDate)
    BarretteM(NbAlea).BInterleaveDataDepth = CStr(objItem.InterleaveDataDepth)
    BarretteM(NbAlea).BInterleavePosition = Type_InterleavePosition(objItem.InterleavePosition)
    BarretteM(NbAlea).BManufacturer = CStr(objItem.Manufacturer)
    BarretteM(NbAlea).BMemoryType = Type_MemoryType(CLng(objItem.MemoryType))
    BarretteM(NbAlea).BModel = CStr(objItem.Model)
    BarretteM(NbAlea).BName = CStr(objItem.Name)
    BarretteM(NbAlea).BOtherIdentifyingInfo = CStr(objItem.OtherIdentifyingInfo)
    BarretteM(NbAlea).BPartNumber = CStr(objItem.PartNumber)
    BarretteM(NbAlea).BPositionInRow = CStr(objItem.PositionInRow)
    BarretteM(NbAlea).BPoweredOn = CStr(objItem.PoweredOn)
    BarretteM(NbAlea).BRemovable = CStr(objItem.Removable)
    BarretteM(NbAlea).BReplaceable = CStr(objItem.Replaceable)
    BarretteM(NbAlea).BSerialNumber = CStr(objItem.SerialNumber)
    BarretteM(NbAlea).BSKU = CStr(objItem.SKU)
    BarretteM(NbAlea).BSpeed = CStr(objItem.Speed)
    BarretteM(NbAlea).BStatus = CStr(objItem.Status)
    BarretteM(NbAlea).BTag = CStr(objItem.Tag)
    BarretteM(NbAlea).BTotalWidth = CStr(objItem.TotalWidth)
    BarretteM(NbAlea).BTypeDetail = Type_TypeDetail(objItem.TypeDetail)
    BarretteM(NbAlea).BVersion = CStr(objItem.Version)
Next
NbBarretteMod = NbAlea
'****************************Calcul*****************************
Dim buf As String
' get system Memory
NbAlea = 0
buf = ""
ArrayM(0).MaxMemPosition = -1
AllSlotsNm = 0
For i = 1 To NbArrayMod
    If UCase(buf) <> UCase(ArrayM(i).AUse) Then
        buf = ArrayM(i).AUse
        NbAlea = NbAlea + 1
        SystemM(NbAlea).SType = buf
    End If
AllSlotsNm = AllSlotsNm + ArrayM(i).AMemoryDevices
ArrayM(i).MinMemPosition = ArrayM(i - 1).MaxMemPosition + 1
ArrayM(i).MaxMemPosition = ArrayM(i).MinMemPosition + ArrayM(i).AMemoryDevices - 1
Next i
NbSystemMod = NbAlea
'------------------------------------------
Dim j As Long
For i = 1 To NbBarretteMod
   For j = 1 To NbArrayMod
    If GetMemNumero(BarretteM(i).BTag) <= ArrayM(j).MaxMemPosition _
    And GetMemNumero(BarretteM(i).BTag) >= ArrayM(j).MinMemPosition Then
      BarretteM(i).ParentAr = ArrayM(j).ATag
      BarretteM(i).ParentSystem = ArrayM(j).AUse
    End If
  Next j
Next i
'----------Max Cap--------------

For i = 1 To NbArrayMod
ArrayM(i).TotalCapInstallee = 0
ArrayM(i).NbMemoire = 0
   For j = 1 To NbBarretteMod
    If BarretteM(j).ParentAr = ArrayM(i).ATag Then
      ArrayM(i).TotalCapInstallee = ArrayM(i).TotalCapInstallee + BarretteM(i).BCapacity
      ArrayM(i).NbMemoire = ArrayM(i).NbMemoire + 1
    End If
  Next j
Next i

For i = 1 To NbSystemMod
SystemM(i).TotalCapInstallee = 0
SystemM(i).NbMemoire = 0
   For j = 1 To NbBarretteMod
    If BarretteM(j).ParentSystem = SystemM(i).SType Then
      SystemM(i).TotalCapInstallee = SystemM(i).TotalCapInstallee + BarretteM(j).BCapacity
      SystemM(i).NbMemoire = SystemM(i).NbMemoire + 1
    End If
  Next j
  
  SystemM(i).MaxCap = 0
  SystemM(i).NbArray = 0
  For j = 1 To NbArrayMod
    If ArrayM(j).AUse = SystemM(i).SType Then
      SystemM(i).MaxCap = SystemM(i).MaxCap + ArrayM(j).AMaxCapacity
      SystemM(i).NbArray = SystemM(i).NbArray + 1
    End If
  Next j
  
Next i


For i = 1 To AllSlotsNm
  If GetDIM(i - 1) <> -1 Then
    SlotM(i).DIMCorresp = GetDIM(i - 1)
  End If
  For j = 1 To NbArrayMod
    If i <= ArrayM(j).MaxMemPosition + 1 _
    And i >= ArrayM(j).MinMemPosition + 1 Then
      SlotM(i).ParentAr = ArrayM(j).ATag
      SlotM(i).ParentSystem = ArrayM(j).AUse
    End If
  Next j
  
Next i

End Sub


Sub ListShares(strComputer, strUser, strPassword)
    Dim strObject
    Dim objLocator, objWMIService, objShare
    Dim colShares
    lstSystem.Clear
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
        For Each objShare In colShares
            lstSystem.AddItem objShare.Name & " [" & objShare.Path & "]"
        Next
End Sub



 

Private Sub SystemInfo()
Dim strObject
On Error Resume Next
    Dim objLocator, objWMIService, objItem
    Dim colItems, strComputer, strUser, strPassword
    strComputer = txtPC.Text
    strUser = txtUser
    strPassword = txtPass
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objWMIService = objLocator.ConnectServer(strComputer, "root/cimv2", strUser, strPassword)
    objWMIService.Security_.impersonationlevel = 3
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objItem In colItems
    SystemeI.AdminPassStatus = "Admin Password Status: " & Type_AdminPasswordStatus(objItem.AdminPasswordStatus)
    HardWareI.AutoResetBoot = "Automatic Reset Boot Option: " & CStr(objItem.AutomaticResetBootOption)
    HardWareI.AutoResetCap = "Automatic Reset Capability: " & CStr(objItem.AutomaticResetCapability)
    SystemeI.Caption = "Caption: " & CStr(objItem.Caption)
    SystemeI.TimeZone = "CurrentTimeZone: " & CStr(objItem.CurrentTimeZone)
    HardWareI.HDescrip = "Description: " & CStr(objItem.Description)
    SystemeI.Domaine = "Domain: " & CStr(objItem.Domain)
    SystemeI.DomaineRole = "Domain Role: " & Type_DomainRole(objItem.DomainRole)
    HardWareI.FrontPanelReset = "Front Panel Reset Status: " & Type_FrontPanelResetStatus(objItem.FrontPanelResetStatus)
    HardWareI.Infrared = "Infrared Supported: " & CStr(objItem.InfraredSupported)
    GeneralI.GFabriquant = "Manufacturer: " & CStr(objItem.Manufacturer)
    GeneralI.GModel = "Model: " & objItem.Model
    HardWareI.NetworkServerMode = "Network Server Mode Enabled: " & CStr(objItem.NetworkServerModeEnabled)
    HardWareI.NbProcesseurs = "Number Of Processors: " & CStr(objItem.NumberOfProcessors)
    SystemeI.OwnerName = "Primary Owner Name: " & CStr(objItem.PrimaryOwnerName)
    HardWareI.HStatus = "Status: " & CStr(objItem.Status)
    HardWareI.HSystemType = "System Type: " & CStr(objItem.SystemType)
    HardWareI.HMemoire = "Total Physical Memory: " & TransMemK(objItem.TotalPhysicalMemory / 1024)
    SystemeI.UserName = "UserName: " & CStr(objItem.UserName)
Next

Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct", , 48)
For Each objItem In colItems
    GeneralI.GDescription = "Description: " & objItem.Description
    GeneralI.Gserial = "Serial Number: " & objItem.IdentifyingNumber
    HardWareI.GUUID = "UUID: " & objItem.UUID
Next

Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem", , 48)
For Each objItem In colItems
   With SystemeI
    .OSName = "OS : " & CStr(objItem.Caption)
    .SPVersion = "SP Version : " & CStr(objItem.CSDVersion)
    .RegUser = "Registered User : " & CStr(objItem.RegisteredUser)
    .OSSerialN = "OS Serial Number : " & CStr(objItem.SerialNumber)
    .MemVirtuelle = "Virtual Memory: " & TransMemK(objItem.TotalVirtualMemorySize)
    .MemPhysic = "Phisical Memory: " & TransMemK(objItem.TotalVisibleMemorySize)
    .SystemDir = "System Directory : " & CStr(objItem.SystemDirectory)
  End With
Next
End Sub


Private Sub Form_Load()
Dim K As String
Dim P As Collection
On Error Resume Next

K = GetSetting("Pierre Programs", "PC Info", "User", "")
Decrypt K, P, 50, ""
txtUser.Text = P(1)
txtPass.Text = P(2)
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 8000 Or Me.Height < 6000 Then Exit Sub
Dim WPic As Long
Line1.X2 = Me.Width
txtMainDesc.Width = Me.Width - 360
txtMainDesc.Height = Me.Height - txtMainDesc.Top - 720
WPic = (txtMainDesc.Width - 360) / 4

lstMain.Width = WPic
lstSystem.Width = WPic
lstSystem.Left = lstMain.Width + lstMain.Left + 120
lstArray.Width = WPic
lstArray.Left = lstSystem.Width + lstSystem.Left + 120
lstSlots.Width = WPic
lstSlots.Left = lstArray.Width + lstArray.Left + 120
picEnCrs.Left = (Me.Width - picEnCrs.Width) / 2
picEnCrs.Top = (Me.Height - picEnCrs.Height) / 2
End Sub

Private Sub lstArray_Click()
Dim i As Long
Dim K As String, L As String
On Error Resume Next
txtMainDesc.Text = ""
K = lstMain.Text
lstSlots.Clear
txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine

Select Case UCase(lstMain.Text)
'----------------------------DRIVES------------------------
Case "DRIVES"
    For i = 1 To DrvNum
       If Left(lstArray.Text, 2) = DrivesM(i).DLetter Then Exit For
    Next i
    With DrivesM(i)
        txtMainDesc.Text = txtMainDesc.Text + .HCaption + NewLineIf(.HCaption)
        txtMainDesc.Text = txtMainDesc.Text + .HVolumeName + NewLineIf(.HVolumeName)
        txtMainDesc.Text = txtMainDesc.Text + .HDescription + NewLineIf(.HDescription)
        txtMainDesc.Text = txtMainDesc.Text + .HDriveType + NewLineIf(.HDriveType)
        txtMainDesc.Text = txtMainDesc.Text + .HFileSystem + NewLineIf(.HFileSystem)
        txtMainDesc.Text = txtMainDesc.Text + .HFreeSpace + NewLineIf(.HFreeSpace)
        txtMainDesc.Text = txtMainDesc.Text + .HSize + NewLineIf(.HSize)
        txtMainDesc.Text = txtMainDesc.Text + .HVolumeSerialNumber + NewLineIf(.HVolumeSerialNumber)
    End With

'----------------------------MEMOIRE----------------------------
Case "MEMORY"
For i = 1 To NbArrayMod
    If lstArray.Text = ArrayM(i).ATag Then
        txtMainDesc.Text = txtMainDesc.Text + "Caption: " + ArrayM(i).ACaption + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Location: " + ArrayM(i).ALocation + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Maximum Capacity: " + TransMemK(ArrayM(i).AMaxCapacity) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Installed Capacity: " + TransMemK(ArrayM(i).TotalCapInstallee) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Number of slots: " + CStr(ArrayM(i).AMemoryDevices) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Number of Installed Memories: " + CStr(ArrayM(i).NbMemoire) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "ECC type: " + ArrayM(i).AMemoryErrorCorrection + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Tag: " + ArrayM(i).ATag + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Use: " + ArrayM(i).AUse + vbNewLine
    End If
Next i

For i = 1 To AllSlotsNm
  If lstArray.Text = SlotM(i).ParentAr Then
    lstSlots.AddItem GetSlotInf(i - 1)
  End If
Next i


End Select
End Sub

Private Sub lstMain_Click()
ClearAllF
ClearInfo
Clear_Info1000
picEnCrs.Visible = True
DoEvents
txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine
If Ping(txtPC.Text, 500, True) <> 0 Then
    picEnCrs.Visible = False
    txtMainDesc.Text = "Machine : " + txtPC.Text + " Isn't on line"
    Exit Sub
End If

Select Case UCase(lstMain.Text)
    Case "MEMORY"
        Dim i As Long
        Call GetMemory
        ' System
        For i = 1 To NbSystemMod
            lstSystem.AddItem SystemM(i).SType
        Next i
        ' Array
        For i = 1 To NbArrayMod
            lstArray.AddItem ArrayM(i).ATag
        Next i
        ' Slots
        For i = 1 To AllSlotsNm
            lstSlots.AddItem GetSlotInf(i - 1)
        Next i
        txtMainDesc.Text = txtMainDesc.Text + "Number of  system modules: " + CStr(NbSystemMod) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Number of Array : " + CStr(NbArrayMod) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Total number of slots : " + CStr(AllSlotsNm) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Total number of installed memories : " + CStr(NbBarretteMod) + vbNewLine
     Case " GENERAL"
        Call SystemInfo
        With GeneralI
        txtMainDesc.Text = txtMainDesc.Text + .GDescription + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .GModel + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .GFabriquant + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .Gserial + vbNewLine
        End With
        lstSystem.AddItem "Hardware"
        lstSystem.AddItem "System"
        
     Case "PROCESSORS"
        GetProcessorInfo
     Case "NETWORK"
        GetNetwork
     Case "DISK"
        GetDisques
     Case "PORTS"
        GetPorts
     Case "ALL HARDWARES"
        GetToutMateriel
     Case "DRIVES"
        GetDrives
     Case "TEST"
        Gettest
     Case "SHARES"
        txtMainDesc.Text = ""
        ListShares txtPC.Text, txtUser, txtPass
End Select
picEnCrs.Visible = False
End Sub

Private Sub lstSlots_Click()
Dim K As String
Dim i As Long
K = lstSlots.Text
txtMainDesc.Text = ""
For i = 1 To AllSlotsNm
If InStr(K, "Slot " + Format(i - 1, "00")) > 0 Then
    If SlotM(i).DIMCorresp > -1 Then
        txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Capacity: " + TransMemK(BarretteM(SlotM(i).DIMCorresp).BCapacity) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Data Width: " + BarretteM(SlotM(i).DIMCorresp).BDataWidth + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Device Locator: " + BarretteM(SlotM(i).DIMCorresp).BDeviceLocator + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Form Factor: " + BarretteM(SlotM(i).DIMCorresp).BFormFactor + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Inter leave Position: " + BarretteM(SlotM(i).DIMCorresp).BInterleavePosition + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Memory Type: " + BarretteM(SlotM(i).DIMCorresp).BMemoryType + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Position In Row: " + BarretteM(SlotM(i).DIMCorresp).BPositionInRow + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Speed: " + BarretteM(SlotM(i).DIMCorresp).BSpeed + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Tag: " + BarretteM(SlotM(i).DIMCorresp).BTag + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Total Width: " + BarretteM(SlotM(i).DIMCorresp).BTotalWidth + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Type Detail: " + BarretteM(SlotM(i).DIMCorresp).BTypeDetail + vbNewLine
  
   
    End If
    Exit For
End If
Next i

End Sub

Private Sub lstSystem_Click()
Dim i As Long
Dim K As String
Dim LBf As String

On Error Resume Next
txtMainDesc.Text = "Machine : " + txtPC.Text + vbNewLine + vbNewLine
K = lstSystem.Text
lstArray.Clear
lstSlots.Clear
Select Case UCase(lstMain.Text)
'--------------------------MEMOIRE-----------------------
  Case "MEMORY"
    For i = 1 To NbArrayMod
      If K = ArrayM(i).AUse Then lstArray.AddItem ArrayM(i).ATag
    Next i
    lstSlots.Clear
    For i = 1 To AllSlotsNm
        If K = SlotM(i).ParentSystem Then lstSlots.AddItem GetSlotInf(i - 1)
    Next i
    If UCase(SystemM(lstSystem.ListIndex + 1).SType) = UCase(K) Then
        txtMainDesc.Text = txtMainDesc.Text + "Total number of installed memory Array in " + K + ": " + CStr(SystemM(lstSystem.ListIndex + 1).NbArray) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Total number of installed memory modules : " + CStr(SystemM(lstSystem.ListIndex + 1).NbMemoire) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Total number of Slots : " + CStr(lstSlots.ListCount) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "installed capacity: " + TransMemK(SystemM(lstSystem.ListIndex + 1).TotalCapInstallee) + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Maximum Capacity: " + TransMemK(SystemM(lstSystem.ListIndex + 1).MaxCap) + vbNewLine
    End If
    
'--------------------------GENERAL-----------------------
  Case " GENERAL"
    Select Case UCase(K)
      Case "HARDWARE"
        With HardWareI
        txtMainDesc.Text = txtMainDesc.Text + .HSystemType + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .GUUID + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .NbProcesseurs + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .HMemoire + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .Infrared + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .NetworkServerMode + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .HStatus + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .FrontPanelReset + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .AutoResetCap + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .AutoResetBoot + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .HDescrip + vbNewLine
        End With
    Case "SYSTEM"
        With SystemeI
        txtMainDesc.Text = txtMainDesc.Text + .Caption + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .OSName + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .SPVersion + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .OSSerialN + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .OwnerName + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .SystemDir + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .MemPhysic + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .MemVirtuelle + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .Domaine + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .DomaineRole + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .AdminPassStatus + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .TimeZone + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + .UserName + vbNewLine
        End With
     End Select
     
'----------------------------PROCESSEUR---------------------
  Case "PROCESSORS"
    With ProcesseurM(lstSystem.ListIndex + 1)
      txtMainDesc.Text = txtMainDesc.Text + .Caption + NewLineIf(.Caption)
      txtMainDesc.Text = txtMainDesc.Text + .AddressWidth + NewLineIf(.AddressWidth)
      txtMainDesc.Text = txtMainDesc.Text + .Architecture + NewLineIf(.Architecture)
      txtMainDesc.Text = txtMainDesc.Text + .CpuStatus + NewLineIf(.CpuStatus)
      txtMainDesc.Text = txtMainDesc.Text + .CurrentClockSpeed + NewLineIf(.CurrentClockSpeed)
      txtMainDesc.Text = txtMainDesc.Text + .MaxClockSpeed + NewLineIf(.MaxClockSpeed)
      txtMainDesc.Text = txtMainDesc.Text + .CurrentVoltage + NewLineIf(.CurrentVoltage)
      txtMainDesc.Text = txtMainDesc.Text + .DataWidth + NewLineIf(.DataWidth)
      txtMainDesc.Text = txtMainDesc.Text + .DeviceID + NewLineIf(.DeviceID)
      txtMainDesc.Text = txtMainDesc.Text + .ExtClock + NewLineIf(.ExtClock)
      txtMainDesc.Text = txtMainDesc.Text + .L2CacheSize + NewLineIf(.L2CacheSize)
      txtMainDesc.Text = txtMainDesc.Text + .L2CacheSpeed + NewLineIf(.L2CacheSpeed)
      txtMainDesc.Text = txtMainDesc.Text + .Level + NewLineIf(.Level)
      txtMainDesc.Text = txtMainDesc.Text + .Manufacturer + NewLineIf(.Manufacturer)
      txtMainDesc.Text = txtMainDesc.Text + .ProcessorId + NewLineIf(.ProcessorId)
      txtMainDesc.Text = txtMainDesc.Text + .Role + NewLineIf(.Role)
      txtMainDesc.Text = txtMainDesc.Text + .SocketDesignation + NewLineIf(.SocketDesignation)
      txtMainDesc.Text = txtMainDesc.Text + .Version + NewLineIf(.Version)

    End With
'---------------------------NETWORK----------------------
  Case "NETWORK"
   With NetAdapterM(lstSystem.ListIndex + 1)
        txtMainDesc.Text = txtMainDesc.Text + "Type : " + .AdapterType + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Type ID : " + .AdapterTypeID + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Description : " + .Description + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Adresse MAC : " + .MACAddress + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "manufacturer : " + .Manufacturer + vbNewLine
        txtMainDesc.Text = txtMainDesc.Text + "Status : " + .NetConnectionStatus + vbNewLine
   End With
   
'---------------------------DISQUES----------------------
  Case "DISK"
    With HDDM(lstSystem.ListIndex + 1)
       txtMainDesc.Text = txtMainDesc.Text + .HCaption + NewLineIf(.HCaption)
       txtMainDesc.Text = txtMainDesc.Text + .HDescription + NewLineIf(.HDescription)
       txtMainDesc.Text = txtMainDesc.Text + .HInterfaceType + NewLineIf(.HInterfaceType)
       txtMainDesc.Text = txtMainDesc.Text + .HManufacturer + NewLineIf(.HManufacturer)
       txtMainDesc.Text = txtMainDesc.Text + .HMediaLoaded + NewLineIf(.HMediaLoaded)
       txtMainDesc.Text = txtMainDesc.Text + .HMediaType + NewLineIf(.HMediaType)
       txtMainDesc.Text = txtMainDesc.Text + .HModel + NewLineIf(.HModel)
       txtMainDesc.Text = txtMainDesc.Text + .HPartitions + NewLineIf(.HPartitions)
       txtMainDesc.Text = txtMainDesc.Text + .HSCSIBus + NewLineIf(.HSCSIBus)
       txtMainDesc.Text = txtMainDesc.Text + .HSCSILogicalUnit + NewLineIf(.HSCSILogicalUnit)
       txtMainDesc.Text = txtMainDesc.Text + .HSCSIPort + NewLineIf(.HSCSIPort)
       txtMainDesc.Text = txtMainDesc.Text + .HSCSITargetId + NewLineIf(.HSCSITargetId)
       txtMainDesc.Text = txtMainDesc.Text + .HSignature + NewLineIf(.HSignature)
       txtMainDesc.Text = txtMainDesc.Text + .HSize + NewLineIf(.HSize)
       txtMainDesc.Text = txtMainDesc.Text + .HStatus + NewLineIf(.HStatus)
       txtMainDesc.Text = txtMainDesc.Text + .HSectorsPerTrack + NewLineIf(.HSectorsPerTrack)
       txtMainDesc.Text = txtMainDesc.Text + .HTotalCylinders + NewLineIf(.HTotalCylinders)
       txtMainDesc.Text = txtMainDesc.Text + .HTotalHeads + NewLineIf(.HTotalHeads)
       txtMainDesc.Text = txtMainDesc.Text + .HTotalSectors + NewLineIf(.HTotalSectors)
       txtMainDesc.Text = txtMainDesc.Text + .HTotalTracks + NewLineIf(.HTotalTracks)
       txtMainDesc.Text = txtMainDesc.Text + .HTracksPerCylinder + NewLineIf(.HTracksPerCylinder)
    End With
  
'---------------------------PORTS----------------------
  Case "PORTS"
   If Left(K, Len("Parallel Port")) = "Parallel Port" Then Call GetParallel
   If Left(K, Len("Serial Port")) = "Serial Port" Then Call GetSerie
   Call GetXPort(K)
'---------------------------TOUT-MATERIEL----------------------
  Case "TOUT-MATERIEL"
     
'----------------------------DRIVES---------------------
  Case "DRIVES"
  lstArray.Clear
  Dim Ar() As String
   For i = 1 To DrvNum
    With DrivesM(i)
        If Type_DriveType(.HDriveTypeN) = Left(lstSystem.Text, Len(Type_DriveType(.HDriveTypeN))) Then
           Ar = Split(.HCaption, ":")
           LBf = Trim$(Ar(1)) + ": "
           Ar = Split(.HVolumeName, ":")
           LBf = LBf + Trim$(Ar(1))
           Ar = Split(.HSize, ":")
           LBf = LBf + " (" + Trim$(Ar(1)) + ")"
           lstArray.AddItem LBf
        End If
    End With
  Next i
End Select


End Sub

Private Sub txtPass_LostFocus()
Dim P As Collection, K As String
Set P = New Collection
P.Add txtUser.Text
P.Add txtPass.Text
K = Encrypt(P, 50, "")
SaveSetting "Pierre Programs", "PC Info", "User", K
End Sub

Private Sub txtPC_Change()
ClearAllF

End Sub


Private Sub txtUser_LostFocus()
Dim P As Collection, K As String
Set P = New Collection
P.Add txtUser.Text
P.Add txtPass.Text
K = Encrypt(P, 50, "")
SaveSetting "Pierre Programs", "PC Info", "User", K
End Sub

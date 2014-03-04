Attribute VB_Name = "ModTmsWmiGetSysInfo"
Option Explicit
Public Function TmsWmiGetDisplayConfigurationInfoSet() As SWbemObjectSet
    Set TmsWmiGetDisplayConfigurationInfoSet = TmsWmiGetSet("Win32_DisplayConfiguration")
    'For Each SWbemObject In TmsWmiGetDisplayConfigurationInfoSet
    'With SWbemObject
    '  uint32 BitsPerPel;
    '  string Caption;
    '  string Description;
    '  string DeviceName;
    '  uint32 DisplayFlags;
    '  uint32 DisplayFrequency;
    '  uint32 DitherType;
    '  string DriverVersion;
    '  uint32 ICMIntent;
    '  uint32 ICMMethod;
    '  uint32 LogPixels;
    '  uint32 PelsHeight;
    '  uint32 PelsWidth;
    '  string SettingID;
    '  uint32 SpecificationVersion;
End Function

Public Function TmsWmiGetDesktopMonitorInfoSet() As SWbemObjectSet
    Set TmsWmiGetDesktopMonitorInfoSet = TmsWmiGetSet("Win32_DesktopMonitor")
    'For Each SWbemObject In TmsWmiGetDesktopMonitorInfoSet
    'With SWbemObject
    '  uint16   Availability;
    '  uint32   Bandwidth;
    '  string   Caption;
    '  uint32   ConfigManagerErrorCode;
    '  boolean  ConfigManagerUserConfig;
    '  string   CreationClassName;
    '  string   Description;
    '  string   DeviceID;
    '  uint16   DisplayType;
    '  boolean  ErrorCleared;
    '  string   ErrorDescription;
    '  datetime InstallDate;
    '  boolean  IsLocked;
    '  uint32   LastErrorCode;
    '  string   MonitorManufacturer;
    '  string   MonitorType;
    '  string   Name;
    '  uint32   PixelsPerXLogicalInch;
    '  uint32   PixelsPerYLogicalInch;
    '  string   PNPDeviceID;
    '  uint16   PowerManagementCapabilities[];
    '  boolean  PowerManagementSupported;
    '  uint32   ScreenHeight;
    '  uint32   ScreenWidth;
    '  string   Status;
    '  uint16   StatusInfo;
    '  string   SystemCreationClassName;
    '  string   SystemName;
End Function

Public Function TmsWmiGetMemoryDeviceInfoSet() As SWbemObjectSet
    Set TmsWmiGetMemoryDeviceInfoSet = TmsWmiGetSet("Win32_MemoryDevice")
    'For Each SWbemObject In TmsWmiGetMemoryDeviceInfoSet
    'With SWbemObject
    '  uint16   Access;
    '  uint8    AdditionalErrorData[];
    '  uint16   Availability;
    '  uint64   BlockSize;
    '  string   Caption;
    '  uint32   ConfigManagerErrorCode;
    '  boolean  ConfigManagerUserConfig;
    '  boolean  CorrectableError;
    '  string   CreationClassName;
    '  string   Description;
    '  string   DeviceID;
    '  uint64   EndingAddress;
    '  uint16   ErrorAccess;
    '  uint64   ErrorAddress;
    '  boolean  ErrorCleared;
    '  uint8    ErrorData[];
    '  uint16   ErrorDataOrder;
    '  string   ErrorDescription;
    '  uint16   ErrorGranularity;
    '  uint16   ErrorInfo;
    '  string   ErrorMethodology;
    '  uint64   ErrorResolution;
    '  datetime ErrorTime;
    '  uint32   ErrorTransferSize;
    '  datetime InstallDate;
    '  uint32   LastErrorCode;
    '  string   Name;
    '  uint64   NumberOfBlocks;
    '  string   OtherErrorDescription;
    '  string   PNPDeviceID;
    '  uint16   PowerManagementCapabilities[];
    '  boolean  PowerManagementSupported;
    '  string   Purpose;
    '  uint64   StartingAddress;
    '  string   Status;
    '  uint16   StatusInfo;
    '  string   SystemCreationClassName;
    '  boolean  SystemLevelAddress;
    '  string   SystemName;
End Function

Public Function TmsWmiGetBIOSInfoSet() As SWbemObjectSet
    Set TmsWmiGetBIOSInfoSet = TmsWmiGetSet("Win32_BIOS")
    'For Each SWbemObject In TmsWmiGetBIOSInfoSet
    'With SWbemObject
    '  uint16   BiosCharacteristics[];
    '  string   BIOSVersion[];
    '  string   BuildNumber;
    '  string   Caption;
    '  string   CodeSet;
    '  string   CurrentLanguage;
    '  string   Description;
    '  string   IdentificationCode;
    '  uint16   InstallableLanguages;
    '  datetime InstallDate;
    '  string   LanguageEdition;
    '  String   ListOfLanguages[];
    '  string   Manufacturer;
    '  string   Name;
    '  string   OtherTargetOS;
    '  boolean  PrimaryBIOS;
    '  datetime ReleaseDate;
    '  string   SerialNumber;
    '  string   SMBIOSBIOSVersion;
    '  uint16   SMBIOSMajorVersion;
    '  uint16   SMBIOSMinorVersion;
    '  boolean  SMBIOSPresent;
    '  string   SoftwareElementID;
    '  uint16   SoftwareElementState;
    '  string   Status;
    '  uint16   TargetOperatingSystem;
    '  string   Version;
End Function

Public Function TmsWmiGetBaseBoardInfoSet() As SWbemObjectSet
    Set TmsWmiGetBaseBoardInfoSet = TmsWmiGetSet("Win32_BaseBoard")
    'For Each SWbemObject In TmsWmiGetBaseBoardInfoSet
    'With SWbemObject
    '  string   Caption;
    '  string   ConfigOptions[];
    '  string   CreationClassName;
    '  real32   Depth;
    '  string   Description;
    '  real32   Height;
    '  boolean  HostingBoard;
    '  boolean  HotSwappable;
    '  datetime InstallDate;
    '  string   Manufacturer;
    '  string   Model;
    '  string   Name;
    '  string   OtherIdentifyingInfo;
    '  string   PartNumber;
    '  boolean  PoweredOn;
    '  string   Product;
    '  boolean  Removable;
    '  boolean  Replaceable;
    '  string   RequirementsDescription;
    '  boolean  RequiresDaughterBoard;
    '  string   SerialNumber;
    '  string   SKU;
    '  string   SlotLayout;
    '  boolean  SpecialRequirements;
    '  string   Status;
    '  string   Tag;
    '  string   Version;
    '  real32   Weight;
    '  real32   Width;
End Function

Public Function TmsWmiGetDiskDriveInfoSet() As SWbemObjectSet
    Set TmsWmiGetDiskDriveInfoSet = TmsWmiGetSet("Win32_DiskDrive")
    'For Each SWbemObject In TmsWmiGetDiskDriveInfoSet
    'With SWbemObject
    '  uint16   Availability;
    '  uint32   BytesPerSector;
    '  uint16   Capabilities[];
    '  string   CapabilityDescriptions[];
    '  string   Caption;
    '  string   CompressionMethod;
    '  uint32   ConfigManagerErrorCode;
    '  boolean  ConfigManagerUserConfig;
    '  string   CreationClassName;
    '  uint64   DefaultBlockSize;
    '  string   Description;
    '  string   DeviceID;
    '  boolean  ErrorCleared;
    '  string   ErrorDescription;
    '  string   ErrorMethodology;
    '  string   FirmwareRevision;
    '  uint32   Index;
    '  datetime InstallDate;
    '  string   InterfaceType;
    '  uint32   LastErrorCode;
    '  string   Manufacturer;
    '  uint64   MaxBlockSize;
    '  uint64   MaxMediaSize;
    '  boolean  MediaLoaded;
    '  string   MediaType;
    '  uint64   MinBlockSize;
    '  string   Model;
    '  string   Name;
    '  boolean  NeedsCleaning;
    '  uint32   NumberOfMediaSupported;
    '  uint32   Partitions;
    '  string   PNPDeviceID;
    '  uint16   PowerManagementCapabilities[];
    '  boolean  PowerManagementSupported;
    '  uint32   SCSIBus;
    '  uint16   SCSILogicalUnit;
    '  uint16   SCSIPort;
    '  uint16   SCSITargetId;
    '  uint32   SectorsPerTrack;
    '  string   SerialNumber;
    '  uint32   Signature;
    '  uint64   Size;
    '  string   Status;
    '  uint16   StatusInfo;
    '  string   SystemCreationClassName;
    '  string   SystemName;
    '  uint64   TotalCylinders;
    '  uint32   TotalHeads;
    '  uint64   TotalSectors;
    '  uint64   TotalTracks;
    '  uint32   TracksPerCylinder;
End Function

Public Function TmsWmiGetDiskPartitionInfoSet() As SWbemObjectSet
    Set TmsWmiGetDiskPartitionInfoSet = TmsWmiGetSet("Win32_DiskPartition")
    'For Each SWbemObject In TmsWmiGetDiskPartitionInfoSet
    'With SWbemObject
    '  uint16   Access;
    '  uint16   Availability;
    '  uint64   BlockSize;
    '  boolean  Bootable;
    '  boolean  BootPartition;
    '  string.  Caption;
    '  uint32   ConfigManagerErrorCode;
    '  boolean  ConfigManagerUserConfig;
    '  string.  CreationClassName;
    '  string   Description;
    '  string   DeviceID;
    '  uint32   DiskIndex;
    '  boolean  ErrorCleared;
    '  string   ErrorDescription;
    '  string   ErrorMethodology;
    '  uint32   HiddenSectors;
    '  uint32   Index;
    '  datetime InstallDate;
    '  uint32   LastErrorCode;
    '  string   Name;
    '  uint64   NumberOfBlocks;
    '  string   PNPDeviceID;
    '  uint16   PowerManagementCapabilities[];
    '  boolean  PowerManagementSupported;
    '  boolean  PrimaryPartition;
    '  string   Purpose;
    '  boolean  RewritePartition;
    '  uint64   Size;
    '  uint64   StartingOffset;
    '  string   Status;
    '  uint16   StatusInfo;
    '  string   SystemCreationClassName;
    '  string   SystemName;
    '  string   Type;
End Function

Public Function TmsWmiGetNetworkAdapterInfoSet() As SWbemObjectSet
    Set TmsWmiGetNetworkAdapterInfoSet = TmsWmiGetSet("Win32_NetworkAdapter")
    'For Each SWbemObject In TmsWmiGetNetworkAdapterInfoSet
    'With SWbemObject
    '  string   AdapterType;
    '  uint16   AdapterTypeID;
    '  boolean  AutoSense;
    '  uint16   Availability;
    '  string   Caption;
    '  uint32   ConfigManagerErrorCode;
    '  boolean  ConfigManagerUserConfig;
    '  string   CreationClassName;
    '  string   Description;
    '  string   DeviceID;
    '  boolean  ErrorCleared;
    '  string   ErrorDescription;
    '  string   GUID;
    '  uint32   Index;
    '  datetime InstallDate;
    '  boolean  Installed;
    '  uint32   InterfaceIndex;
    '  uint32   LastErrorCode;
    '  string   MACAddress;
    '  string   Manufacturer;
    '  uint32   MaxNumberControlled;
    '  uint64   MaxSpeed;
    '  string   Name;
    '  string   NetConnectionID;
    '  uint16   NetConnectionStatus;
    '  boolean  NetEnabled;
    '  string   NetworkAddresses[];
    '  string   PermanentAddress;
    '  boolean  PhysicalAdapter;
    '  string   PNPDeviceID;
    '  uint16   PowerManagementCapabilities[];
    '  boolean  PowerManagementSupported;
    '  string   ProductName;
    '  string   ServiceName;
    '  uint64   Speed;
    '  string   Status;
    '  uint16   StatusInfo;
    '  string   SystemCreationClassName;
    '  string   SystemName;
    '  datetime TimeOfLastReset;
End Function

Public Function TmsWmiGetProcessorInfoSet() As SWbemObjectSet
    Set TmsWmiGetProcessorInfoSet = TmsWmiGetSet("Win32_Processor")
    'For Each SWbemObject In TmsWmiGetProcessorInfoSet
    'With SWbemObject
    '  uint16   AddressWidth;
    '  uint16   Architecture;
    '  uint16   Availability;
    '  string   Caption;
    '  uint32   ConfigManagerErrorCode;
    '  boolean  ConfigManagerUserConfig;
    '  uint16   CpuStatus;
    '  string   CreationClassName;
    '  uint32   CurrentClockSpeed;
    '  uint16   CurrentVoltage;
    '  uint16   DataWidth;
    '  string   Description;
    '  string   DeviceID;
    '  boolean  ErrorCleared;
    '  string   ErrorDescription;
    '  uint32   ExtClock;
    '  uint16   Family;
    '  datetime InstallDate;
    '  uint32   L2CacheSize;
    '  uint32   L2CacheSpeed;
    '  uint32   L3CacheSize;
    '  uint32   L3CacheSpeed;
    '  uint32   LastErrorCode;
    '  uint16   Level;
    '  uint16   LoadPercentage;
    '  string   Manufacturer;
    '  uint32   MaxClockSpeed;
    '  string   Name;                           CPU名称
    '  uint32   NumberOfCores;                  物理核心数量
    '  uint32   NumberOfLogicalProcessors;      逻辑核心数量
    '  string   OtherFamilyDescription;
    '  string   PNPDeviceID;
    '  uint16   PowerManagementCapabilities[];
    '  boolean  PowerManagementSupported;
    '  string   ProcessorId;                    CPU序列号
    '  uint16   ProcessorType;
    '  uint16   Revision;
    '  string   Role;
    '  string   SocketDesignation;
    '  string   Status;
    '  uint16   StatusInfo;
    '  string   Stepping;
    '  string   SystemCreationClassName;
    '  string   SystemName;
    '  string   UniqueId;
    '  uint16   UpgradeMethod;
    '  string   Version;
    '  uint32   VoltageCaps;
End Function

Public Function TmsWmiGetSet(instances As String) As SWbemObjectSet
    Set TmsWmiGetSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(instances)
End Function

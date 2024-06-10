Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set os = wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")

For Each value In os

WScript.StdOut.WriteLine "Hello, world!"

WScript.StdOut.WriteLine "BootDevice: " & value.BootDevice
WScript.StdOut.WriteLine "BuildNumber: " & value.BuildNumber
WScript.StdOut.WriteLine "BuildType: " & value.BuildType
WScript.StdOut.WriteLine "Caption: " & value.Caption
WScript.StdOut.WriteLine "CodeSet: " & value.CodeSet
WScript.StdOut.WriteLine "CountryCode: " & value.CountryCode
WScript.StdOut.WriteLine "CreationClassName: " & value.CreationClassName
WScript.StdOut.WriteLine "CSCreationClassName: " & value.CSCreationClassName
WScript.StdOut.WriteLine "CSDVersion: " & value.CSDVersion
WScript.StdOut.WriteLine "CSName: " & value.CSName
WScript.StdOut.WriteLine "CurrentTimeZone: " & value.CurrentTimeZone
WScript.StdOut.WriteLine "DataExecutionPrevention_Available: " & value.DataExecutionPrevention_Available
WScript.StdOut.WriteLine "DataExecutionPrevention_32BitApplications: " & value.DataExecutionPrevention_32BitApplications
WScript.StdOut.WriteLine "DataExecutionPrevention_Drivers: " & value.DataExecutionPrevention_Drivers
WScript.StdOut.WriteLine "DataExecutionPrevention_SupportPolicy: " & value.DataExecutionPrevention_SupportPolicy
WScript.StdOut.WriteLine "Debug: " & value.Debug
WScript.StdOut.WriteLine "Description: " & value.Description
WScript.StdOut.WriteLine "Distributed: " & value.Distributed
WScript.StdOut.WriteLine "EncryptionLevel: " & value.EncryptionLevel
WScript.StdOut.WriteLine "ForegroundApplicationBoost: " & value.ForegroundApplicationBoost
WScript.StdOut.WriteLine "FreePhysicalMemory: " & value.FreePhysicalMemory
WScript.StdOut.WriteLine "FreeSpaceInPagingFiles: " & value.FreeSpaceInPagingFiles
WScript.StdOut.WriteLine "FreeVirtualMemory: " & value.FreeVirtualMemory
WScript.StdOut.WriteLine "InstallDate: " & value.InstallDate
WScript.StdOut.WriteLine "LargeSystemCache: " & value.LargeSystemCache
WScript.StdOut.WriteLine "LastBootUpTime: " & value.LastBootUpTime
WScript.StdOut.WriteLine "LocalDateTime: " & value.LocalDateTime
WScript.StdOut.WriteLine "Locale: " & value.Locale
WScript.StdOut.WriteLine "Manufacturer: " & value.Manufacturer
WScript.StdOut.WriteLine "MaxNumberOfProcesses: " & value.MaxNumberOfProcesses
WScript.StdOut.WriteLine "MaxProcessMemorySize: " & value.MaxProcessMemorySize
'WScript.StdOut.WriteLine "MUILanguages[]: " & value.MUILanguages[]
WScript.StdOut.WriteLine "Name: " & value.Name
WScript.StdOut.WriteLine "NumberOfLicensedUsers: " & value.NumberOfLicensedUsers
WScript.StdOut.WriteLine "NumberOfProcesses: " & value.NumberOfProcesses
WScript.StdOut.WriteLine "NumberOfUsers: " & value.NumberOfUsers
WScript.StdOut.WriteLine "OperatingSystemSKU: " & value.OperatingSystemSKU
WScript.StdOut.WriteLine "Organization: " & value.Organization
WScript.StdOut.WriteLine "OSArchitecture: " & value.OSArchitecture
WScript.StdOut.WriteLine "OSLanguage: " & value.OSLanguage
WScript.StdOut.WriteLine "OSProductSuite: " & value.OSProductSuite
WScript.StdOut.WriteLine "OSType: " & value.OSType
WScript.StdOut.WriteLine "OtherTypeDescription: " & value.OtherTypeDescription
WScript.StdOut.WriteLine "PAEEnabled: " & value.PAEEnabled
WScript.StdOut.WriteLine "PlusProductID: " & value.PlusProductID
WScript.StdOut.WriteLine "PlusVersionNumber: " & value.PlusVersionNumber
WScript.StdOut.WriteLine "PortableOperatingSystem: " & value.PortableOperatingSystem
WScript.StdOut.WriteLine "Primary: " & value.Primary
WScript.StdOut.WriteLine "ProductType: " & value.ProductType
WScript.StdOut.WriteLine "RegisteredUser: " & value.RegisteredUser
WScript.StdOut.WriteLine "SerialNumber: " & value.SerialNumber
WScript.StdOut.WriteLine "ServicePackMajorVersion: " & value.ServicePackMajorVersion
WScript.StdOut.WriteLine "ServicePackMinorVersion: " & value.ServicePackMinorVersion
WScript.StdOut.WriteLine "SizeStoredInPagingFiles: " & value.SizeStoredInPagingFiles
WScript.StdOut.WriteLine "Status: " & value.Status
WScript.StdOut.WriteLine "SuiteMask: " & value.SuiteMask
WScript.StdOut.WriteLine "SystemDevice: " & value.SystemDevice
WScript.StdOut.WriteLine "SystemDirectory: " & value.SystemDirectory
WScript.StdOut.WriteLine "SystemDrive: " & value.SystemDrive
WScript.StdOut.WriteLine "TotalSwapSpaceSize: " & value.TotalSwapSpaceSize
WScript.StdOut.WriteLine "TotalVirtualMemorySize: " & value.TotalVirtualMemorySize
WScript.StdOut.WriteLine "TotalVisibleMemorySize: " & value.TotalVisibleMemorySize
WScript.StdOut.WriteLine "Version: " & value.Version
WScript.StdOut.WriteLine "WindowsDirectory: " & value.WindowsDirectory
'WScript.StdOut.WriteLine "QuantumLength: " & value.QuantumLength
'WScript.StdOut.WriteLine "QuantumType: " & value.QuantumType

Next

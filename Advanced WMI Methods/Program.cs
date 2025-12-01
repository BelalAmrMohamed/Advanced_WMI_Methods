using System;
using System.Collections.Generic;
using System.Data;
using System.Management;
using System.Text;

namespace Advanced_WMI_Methods
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DisplayMenuAndExecuteQuery();

            _ = Console.ReadKey(true);
        }

        private static StringBuilder[] GetPartitionInfo() // This is merged with 'GetLogicalDiskInfo()'
        {
            // 'collection' is for getting the number that we will set the StringBuilder with.
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk");
            ManagementObjectCollection collection = searcher.Get();

            StringBuilder[] info = new StringBuilder[collection.Count];
            int index = 0;
            const double BytesInGB = 1024.0 * 1024.0 * 1024.0;

            foreach (ManagementObject obj in collection)
            {
                info[index] = new StringBuilder();

                info[index].AppendLine($"{"Drive:", -15} {obj["Name"]}");
                info[index].AppendLine($"{"ID:",-15} {obj["DeviceID"]}");
                info[index].AppendLine($"{"File system:", -15} {obj["FileSystem"]}");
                info[index].AppendLine($"{"Description:",-15} {obj["Description"]}");

                double size = Convert.ToDouble(obj["Size"]) / BytesInGB;
                double free = Convert.ToDouble(obj["FreeSpace"]) / BytesInGB;

                info[index].AppendLine($"{"Size:", -15} {Math.Round(size, 2)} GB");
                info[index].AppendLine($"{"Free space:", -15} {Math.Round(free, 2)} GB");   

                index++;
            }
            return info;
        }        

        private static string GetComputerSystemInfo()
        {
            ManagementObjectSearcher Searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            StringBuilder info = new StringBuilder();
            foreach (ManagementObject obj in Searcher.Get())
            {
                info.AppendLine($"{"Computer Name:",-22} {obj["Name"]}");
                info.AppendLine($"{"Domain:",-22} {obj["Domain"]}");
                info.AppendLine($"{"Model:",-22} {obj["Model"]}");
                info.AppendLine($"{"Manufacturer:",-22} {obj["Manufacturer"]}");
                info.AppendLine($"{"Total Physical Memory:",-22} {Math.Round(Convert.ToDouble(obj["TotalPhysicalMemory"]) / (1024.0 * 1024.0 * 1024.0), 2)} GB");
                info.AppendLine($"{"System Type:",-22} {obj["SystemType"]}");
                info.AppendLine($"{"Workgroup/Domain Join:",-22} {obj["Workgroup"]}");
            }
            return info.ToString();
        }
        private static string GetComputerType()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            StringBuilder type = new StringBuilder(30);
            foreach (ManagementObject obj in searcher.Get())
            {
                switch (Convert.ToInt32(obj["DomainRole"]))
                {
                    case 0:
                        type.Append("Standalone Workstation");
                        break;
                    case 1:
                        type.Append("Member Workstation"); // Represents a member of a domain
                        break;
                    case 2:
                        type.Append("Primary Domain Controller");
                        break;
                    case 3:
                        type.Append("Backup Domain Controller");
                        break;
                    // Note: Roles 4 and 5 are legacy/deprecated and rarely used in modern systems.
                    case 4:
                        type.Append("Standalone Server"); // Standalone server that's not part of a domain
                        break;
                    case 5:
                        type.Append("Member Server"); // Server that is a member of a domain
                        break;
                    default:
                        type.Append("Role Undefined (Code: " + Convert.ToInt32(obj["DomainRole"]) + ")");
                        break;
                }
            }
            return type.ToString();
        }
        private static void RenameComputer(string name)  // This will change the device name
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");

            object[] newName = { name };
            foreach (ManagementObject obj in searcher.Get())
            {
                obj.InvokeMethod("Rename", newName);
            }
        }

        private static string GetProductInfo()
        {
            ManagementObjectSearcher os = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystemProduct");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in os.Get())
            {
                info.AppendLine($"{"Manufacturer:",-20} {obj["Vendor"]}\n");
                info.AppendLine($"{"UUID:",-20} {obj["UUID"]}\n");
                info.AppendLine($"{"Name:",-20} {obj["Name"]}\n");
                info.AppendLine($"{"Identifying Number:",-20} {obj["IdentifyingNumber"]}\n");
            }
            return info.ToString();
        }

        private static string GetProcessorInfo()
        {
            ManagementObjectSearcher cpuSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in cpuSearcher.Get())
            {
                info.AppendLine($"Number of Cores: {obj["NumberOfCores"]}");
            }
            return info.ToString();
        }

        private static string Get_OS_Info()
        {
            StringBuilder info = new StringBuilder();
            ManagementObjectSearcher osSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject obj in osSearcher.Get())
            {
                info.AppendLine($"{"Name:",-20}{obj["Caption"]}");
                info.AppendLine($"{"Version:",-20}{obj["Version"]}");
                info.AppendLine($"{"Manufacturer:",-20}{obj["Manufacturer"]}");
                info.AppendLine($"{"Windows Directory:",-20}{obj["WindowsDirectory"]}");
            }
            return info.ToString();
        }

        private static string GetDesktopInfo()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Desktop WHERE Name = '.Default'");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in searcher.Get())
            {
                info.AppendLine($"{"Desktop Name:",-20} {obj["Name"]}");
                info.AppendLine($"{"Icon Title Size:",-20} {obj["IconTitleSize"]}");
                info.AppendLine($"{"Wallpaper Stretched:",-20} {obj["WallpaperStretched"]}");
                info.AppendLine($"{"Is there a screen saver:",-20} {obj["ScreenSaverActive"]}");

                try
                {
                    if (obj["ScreenSaverActive"].ToString() != "False")
                    {
                        info.AppendLine($"{"Screen Saver time out:",-20} {obj["ScreenSaverTimeout"]}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return info.ToString();
        }

        private static string GetAllDesktopInfo()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Desktop WHERE Name = '.Default'");
            StringBuilder allInfo = new StringBuilder();
            foreach (ManagementObject obj in searcher.Get())
            {
                foreach (PropertyData prop in obj.Properties)
                {
                    allInfo.AppendLine($"{prop.Name} : {prop.Value}");
                }
            }
            return allInfo.ToString();
        }

        private static string GetMemoryInformation()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PerfFormattedData_PerfOS_Memory");
            StringBuilder info = new StringBuilder();
            foreach (ManagementObject obj in searcher.Get())
            {
                info.AppendLine($"{"Available MBs:",-20} {obj["AvailableMbytes"]}");
                info.AppendLine($"{"Cache Bytes:",-20} {obj["CacheBytes"]}");
                info.AppendLine($"{"Committed Bytes:",-20} {obj["CommittedBytes"]}");
                info.AppendLine($"{"Commit Limit:",-20} {obj["CommitLimit"]}");
            }
            return info.ToString();
        }

        // Program 1: Get Logical disk info
        // query: "SELECT * FROM Win32_LogicalDisk WHERE Device ID = 'C: ' ",
        // Properties: DeviceID, Description, FreeSpace, Size
        //private static string GetLogicalDiskInfo()
        //{
        //    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk");
        //    StringBuilder info = new StringBuilder();
        //    const double BytesInGB = 1024.0 * 1024.0 * 1024.0;

        //    foreach (ManagementObject obj in searcher.Get())
        //    {
        //        double freeSpaceGB = Convert.ToDouble(obj["FreeSpace"]) / BytesInGB;
        //        double diskSizeGB = Convert.ToDouble(obj["Size"]) / BytesInGB;

        //        info.AppendLine($"{"Name:",-15} {obj["DeviceID"]}");
        //        info.AppendLine($"{"Description:",-15} {obj["Description"]}");

        //        info.AppendLine($"{"Free space:",-15} {Math.Round(freeSpaceGB, 2)} GB");
        //        info.AppendLine($"{"Disk size:",-15} {Math.Round(diskSizeGB, 2)} GB");
        //    }
        //    return info.ToString();
        //}

        // Program 2: Get CD-ROM/DVD information
        // Query: Win32 CDROMDrive
        // Properties: Description, Dri ve, MediaType, Size, TransferRate
        // Note: If you don't have DVD (Laptop users) you won't see output
        private static string GET_CD_RomInfo()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_CDROMDrive");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in searcher.Get())
            {
                info.AppendLine($"{"Description:",-15} {obj["Description"]}");
                info.AppendLine($"{"Drive:",-15} {obj["Drive"]}");
                info.AppendLine($"{"Media Type:",-15} {obj["MediaType"]}");
                info.AppendLine($"{"Size:",-15} {obj["Size"]}");
                info.AppendLine($"{"Transfer Rate:",-15} {obj["TransferRate"]}");
            }
            return info.ToString();
        }

        // Program 3: Get boot configuration
        // Query: Win32_BootConfiguration
        // Properties: BootDirectory, Description, ScratchDirectory, TempDirectory
        private static string GetBootConfiguration()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BootConfiguration");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in searcher.Get())
            {
                info.AppendLine($"{"BootDirectory:",-20} {obj["BootDirectory"]}");
                info.AppendLine($"{"Description:",-20} {obj["Description"]}");
                info.AppendLine($"{"Scratch Directory:",-20} {obj["ScratchDirectory"]}");
                info.AppendLine($"{"Temp Directory:",-20} {obj["TempDirectory"]}");
            }
            return info.ToString();
        }

        // Program 4: get List of file shares on Local machine
        // Query: Win32_Share
        // Properties: Name, Path, Description
        private static string GetListOfFileShares()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Share");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in searcher.Get())
            {
                info.AppendLine($"{"Name:",-15} {obj["Name"]}");
                info.AppendLine($"{"Path:",-15} {obj["Path"]}");
                info.AppendLine($"{"Description:",-15} {obj["Description"]}");
            }
            return info.ToString();
        }

        private static string GetServices(string state = "")
        {
            string whereClause;

            if (!string.IsNullOrWhiteSpace(state)) 
                whereClause = $" WHERE State='{state}'";           
            
            else whereClause = string.Empty;
            
            ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Service{state}");
            StringBuilder info = new StringBuilder();

            foreach (ManagementObject obj in searcher.Get())
            {
                info.AppendLine($"{"Service Name:",-15} {obj["DisplayName"]}");
                info.AppendLine($"{"Start Mode:",-15} {obj["StartMode"]}");
                info.AppendLine($"{"Description:",-15} {obj["Description"]}");
            }
            return info.ToString();
        }

        private static string GetBatteryInfo()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Battery");
            StringBuilder info = new StringBuilder();

            // Check if any battery objects were found (will be empty on desktops)
            ManagementObjectCollection batteries = searcher.Get();

            if (batteries.Count == 0)
            {
                info.AppendLine("No Win32_Battery instances found (likely a desktop PC).");
            }
            else
            {
                foreach (ManagementObject obj in batteries)
                {
                    info.AppendLine($"{"Device ID:",-25} {obj["DeviceID"]}");
                    info.AppendLine($"{"Design Capacity:",-25} {obj["DesignCapacity"]} mWh");
                    info.AppendLine($"{"Full Charge Capacity:",-25} {obj["FullChargeCapacity"]} mWh");
                    info.AppendLine($"{"Estimated Run Time:",-25} {obj["EstimatedRunTime"]} minutes");
                    //info.AppendLine($"{"Remaining Capacity:",-25} {obj["RemainingCapacity"]} mWh");
                    info.AppendLine($"{"Battery Status Code:",-25} {obj["BatteryStatus"]}");
                }
            }
            return info.ToString();
        }

        //private static string GetUserAccount() // Make this method return an array of StringBuilder[], each index has the info of a user
        //{
        //    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_UserAccount");
        //    StringBuilder info = new StringBuilder();

        //    foreach (ManagementObject obj in searcher.Get())
        //    {
        //        info.AppendLine($"{"User name:",-15} {obj["Name"]}");
        //        info.AppendLine($"{"Domain:",-15} {obj["Domain"]}");
        //        info.AppendLine($"{"Status:",-15} {obj["Status"]}");
        //        info.AppendLine($"{"Disabled:",-15} {obj["Disabled"]}");
        //        info.AppendLine($"{"Local account:",-15} {obj["LocalAccount"]}");
        //    }
        //    return info.ToString();
        //}

        private static StringBuilder[] GetUserAccount() // Similar to "GetPartitionInfo()"
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_UserAccount");
            ManagementObjectCollection collection = searcher.Get();
            StringBuilder[] userAccounts = new StringBuilder[collection.Count];
            int index = 0;

            foreach (ManagementObject obj in collection)
            {
                userAccounts[index] = new StringBuilder();

                userAccounts[index].AppendLine($"{"User name:",-15} {obj["Name"]}");
                userAccounts[index].AppendLine($"{"Domain:",-15} {obj["Domain"]}");
                userAccounts[index].AppendLine($"{"Status:",-15} {obj["Status"]}");
                userAccounts[index].AppendLine($"{"Disabled:",-15} {obj["Disabled"]}");
                userAccounts[index].AppendLine($"{"Local account:",-15} {obj["LocalAccount"]}");

                index++;
            }
            return userAccounts;
        }

        private static string GetPartitionInfoString()
        {
            // Call the original method which returns StringBuilder[]
            StringBuilder[] partitionArray = GetPartitionInfo();

            // Concatenate all StringBuilder objects into one string
            StringBuilder finalOutput = new StringBuilder();
            foreach (StringBuilder sb in partitionArray)
            {
                finalOutput.AppendLine(sb.ToString());
                finalOutput.AppendLine(new string('-', 40));
            }
            return finalOutput.ToString();            
        }

        private static string GetUserAccountString()
        {
            // Call the original method which returns StringBuilder[]
            StringBuilder[] partitionArray = GetUserAccount();

            // Concatenate all StringBuilder objects into one string
            StringBuilder finalOutput = new StringBuilder();
            foreach (StringBuilder sb in partitionArray)
            {
                finalOutput.AppendLine(sb.ToString());
                finalOutput.AppendLine(new string('-', 40));
            }
            return finalOutput.ToString();

        }

        private static void DisplayMenuAndExecuteQuery()
        {
            // 1. Map Menu Options to Methods using a Dictionary of Delegates (Func<string>)
            // The key is the menu number, and the value is the function to call.
            var menuOptions = new Dictionary<int, Func<string>>
    {
        { 1, GetPartitionInfoString },
        { 2, GetComputerSystemInfo },
        { 3, GetProcessorInfo },
        { 4, Get_OS_Info },
        { 5, GetDesktopInfo },
        { 6, GetAllDesktopInfo },
        { 7, GetMemoryInformation },
        //{ 8, GetLogicalDiskInfo },
        { 8, GetProductInfo },
        { 9, GET_CD_RomInfo },
        { 10, GetBootConfiguration },
        { 11, GetListOfFileShares },
        { 12, GetUserAccountString },
        { 13, GetBatteryInfo },
        { 14, GetComputerType },
    };

            // 2. Display the Menu
            Console.WriteLine("╔═════════════════════════════════════════════╗");
            Console.WriteLine("║        WMI System Information Queries       ║");
            Console.WriteLine("╚═════════════════════════════════════════════╝");
            Console.WriteLine("Select a number to run the corresponding query:");
            Console.WriteLine("---------------------------------------------");

            // Iterate through the dictionary to display the menu options dynamically
            foreach (var kvp in menuOptions)
            {
                // Use reflection (or a known convention) to get the name of the method
                // For simplicity, we'll use a hardcoded name based on the key
                string methodName = kvp.Value.Method.Name;
                Console.WriteLine($"[{kvp.Key,2}] {methodName}");
            }

            // Add the special case method
            Console.WriteLine("[15] GetServices (Requires state: Running/Stopped)");
            Console.WriteLine("[ 0] Exit");
            Console.WriteLine("---------------------------------------------");

            // 3. Get and Process User Choice
            Console.Write("Enter your choice (0-15): ");
            string input = Console.ReadLine();

            if (int.TryParse(input, out int choice))
            {
                if (choice == 0)
                {
                    Console.WriteLine("Exiting WMI Query Tool.");
                    return;
                }
                else if (menuOptions.ContainsKey(choice))
                {
                    // Execute the chosen method from the dictionary
                    Console.WriteLine($"\n--- Running Query: {menuOptions[choice].Method.Name} ---");
                    string result = menuOptions[choice].Invoke();
                    Console.WriteLine(result);
                }
                else if (choice == 15)
                {
                    // Handle the method that requires an argument (GetServices)
                    Console.Write("Enter service state ('running', 'stopped', or leave empty): ");
                    string state = Console.ReadLine().ToLowerInvariant() ?? "";
                    if (state == "running" || state == "stopped" || state == "")
                    {
                        Console.WriteLine("\n--- Running Query: GetServices ---");
                        string result = GetServices(state);
                        Console.WriteLine(result);
                    }
                    else
                    {
                        Console.WriteLine("Invalid service state entered. Please try again.");
                    }
                }
                else
                {
                    Console.WriteLine("Invalid choice. Please enter a number from the menu.");
                }
            }
            else
            {
                Console.WriteLine("Invalid input. Please enter a number.");
            }
        }

    }
}
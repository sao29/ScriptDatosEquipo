using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        string ramBytesStr = GetWMI("Win32_ComputerSystem", "TotalPhysicalMemory");

        var info = new Dictionary<string, object>
        {
            ["nombre_equipo"] = Environment.MachineName,
            ["modelo"] = GetWMI("Win32_ComputerSystem", "Model"),
            ["serial"] = GetWMI("Win32_BIOS", "SerialNumber"),
            ["sistema_operativo"] = Environment.OSVersion.ToString(),
            ["cpu"] = GetWMI("Win32_Processor", "Name"),
            ["cores"] = GetWMI("Win32_Processor", "NumberOfCores"),
            ["threads"] = GetWMI("Win32_Processor", "NumberOfLogicalProcessors"),
            ["ram"] = ramBytesStr, // <-- Enviar el valor en bytes, sin cálculos
            ["ip"] = GetLocalIPv4(),
            ["mac"] = GetMacAddress(),
            ["tarjeta_madre"] = GetWMI("Win32_BaseBoard", "Product"),
            ["serial_motherboard"] = GetWMI("Win32_BaseBoard", "SerialNumber"),
            ["uuid"] = GetWMI("Win32_ComputerSystemProduct", "UUID"),
            ["office_licencia_ultimos5"] = GetOfficeLicenseLast5()
        };

        string json = JsonConvert.SerializeObject(info, Formatting.Indented);

        Console.WriteLine($"RAM a guardar/enviar: {ramBytesStr} bytes");
        Console.WriteLine("JSON a enviar:");
        Console.WriteLine(json);

        using (var client = new HttpClient())
        {
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await client.PostAsync("http://cmdb.oorale.com/api/equipos/guardarEquipo", content);

            if (response.IsSuccessStatusCode)
                Console.WriteLine("Inventario enviado exitosamente.");
            else
            {
                Console.WriteLine($"Error al enviar inventario: {response.StatusCode}");
                var responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Respuesta del servidor: {responseBody}");
            }
        }

        Console.WriteLine("Presiona Enter para salir...");
        Console.ReadLine();
    }

    static string GetWMI(string className, string property)
    {
        try
        {
            using (var searcher = new ManagementObjectSearcher($"SELECT {property} FROM {className}"))
            {
                foreach (var obj in searcher.Get())
                {
                    return obj[property]?.ToString();
                }
            }
        }
        catch { }
        return "";
    }

    static string GetLocalIPv4()
    {
        foreach (var ni in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (ni.OperationalStatus == OperationalStatus.Up && ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
            {
                foreach (var ip in ni.GetIPProperties().UnicastAddresses)
                {
                    if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        return ip.Address.ToString();
                    }
                }
            }
        }
        return "";
    }

    static string GetMacAddress()
    {
        foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
            {
                var mac = nic.GetPhysicalAddress().ToString();
                if (!string.IsNullOrEmpty(mac))
                    return mac;
            }
        }
        return "";
    }


    static string GetOfficeLicenseLast5()
    {
        string[] possiblePaths = {
        @"C:\Program Files\Microsoft Office\Office14\ospp.vbs",
        @"C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs",
        @"C:\Program Files\Microsoft Office\Office15\ospp.vbs",
        @"C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs",
        @"C:\Program Files\Microsoft Office\Office16\ospp.vbs",
        @"C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs",
        @"C:\Program Files\Microsoft Office\root\Office16\ospp.vbs",
        @"C:\Program Files (x86)\Microsoft Office\root\Office16\ospp.vbs"
    };

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                var psi = new System.Diagnostics.ProcessStartInfo("cscript")
                {
                    Arguments = $"//NoLogo \"{path}\" /dstatus",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = System.Diagnostics.Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    foreach (var line in output.Split('\n'))
                    {
                        if (line.Contains("Last 5 characters of installed product key"))
                        {
                            var parts = line.Split(':');
                            if (parts.Length > 1)
                                return parts[1].Trim();
                        }
                    }
                }
            }
        }
        return "No encontrado";
    }
}
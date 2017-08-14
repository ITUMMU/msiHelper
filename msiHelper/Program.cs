using CommandLine;
using CommandLine.Text;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace msiHelper
{
    class Program
    {
        private static Boolean Verbose = false;

        static void writeLog(String Message)
        {
            if (Verbose)
            {
                Console.Error.WriteLine(Message);
            }
        }

        static void Main(string[] args)
        {
            Environment.ExitCode = 0;
            Options options = new Options();
            //Parser parser = new Parser(new ParserSettings { MutuallyExclusive = true, CaseSensitive = true, HelpWriter = Console.Error });
            Parser parser = new Parser(s =>
            {
                s.MutuallyExclusive = true;
                s.CaseSensitive = true;
                s.HelpWriter = Console.Error;
            });

            try
            {
                if (parser.ParseArguments(args, options))
                {
                    Verbose = options.Verbose;

                    if (!String.IsNullOrEmpty(options.LogFile))
                    {
                        Console.SetError(new StreamWriter(options.LogFile, false, Encoding.Unicode));
                    }

                    writeLog(String.Format(@"Searching for package ""{0}""", options.Package));
                    SoftwareSearch swSearch = new SoftwareSearch(options.Package);

                    if (swSearch.Installed == true)
                    {
                        if (options.Uninstall)
                        {
                            writeLog(String.Format(@"Uninstalling package: {0}", swSearch.GUID));
                            ProcessStartInfo cmdsi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, @"MsiExec.exe"));
                            cmdsi.Arguments = String.Format(@"/uninstall {0}", swSearch.GUID);
                            if (options.Quiet) cmdsi.Arguments += @" /quiet";
                            if (options.NoRestart) cmdsi.Arguments += @" /norestart";
                            if (!String.IsNullOrEmpty(options.CmdLine)) cmdsi.Arguments += String.Format(@" {0}", options.CmdLine);
                            if (!String.IsNullOrEmpty(options.LogFile))
                            {
                                cmdsi.Arguments += String.Format(@" /l*+ ""{0}""", options.LogFile);
                                Console.Error.Flush();
                                Console.Error.Close();
                            }

                            Process cmd = Process.Start(cmdsi);
                            cmd.WaitForExit();

                            if (!String.IsNullOrEmpty(options.LogFile))
                            {
                                Console.SetError(new StreamWriter(options.LogFile, true, Encoding.Unicode));
                            }

                            writeLog(String.Format(@"Result code is {0}", cmd.ExitCode));
                            Environment.ExitCode = cmd.ExitCode;
                        }
                        else if (!String.IsNullOrEmpty(options.Install))
                        {
                            writeLog(String.Format(@"Reinstalling package: {0}", options.Install));
                            ProcessStartInfo cmdsi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, @"MsiExec.exe"));
                            cmdsi.Arguments = String.Format(@"/fvomus ""{0}""", options.Install);
                            if (options.Quiet) cmdsi.Arguments += @" /quiet";
                            if (options.NoRestart) cmdsi.Arguments += @" /norestart";
                            if (!String.IsNullOrEmpty(options.LogFile))
                            {
                                cmdsi.Arguments += String.Format(@" /l*+ ""{0}""", options.LogFile);
                                Console.Error.Flush();
                                Console.Error.Close();
                            }
                            if (!String.IsNullOrEmpty(options.CmdLine)) cmdsi.Arguments += String.Format(@" {0}", options.CmdLine);

                            Process cmd = Process.Start(cmdsi);
                            cmd.WaitForExit();

                            if (!String.IsNullOrEmpty(options.LogFile))
                            {
                                Console.SetError(new StreamWriter(options.LogFile, true, Encoding.Unicode));
                            }

                            writeLog(String.Format(@"Result code is {0}", cmd.ExitCode));
                            Environment.ExitCode = cmd.ExitCode;
                        }
                        else
                        {
                            writeLog(String.Format(@"Package GUID: {0}", swSearch.GUID));
                            Console.WriteLine(swSearch.GUID);
                        }
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(options.Install))
                        {
                            writeLog(String.Format(@"Installing package {0}", options.Package));
                            ProcessStartInfo cmdsi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, @"MsiExec.exe"));
                            cmdsi.Arguments = String.Format(@"/package ""{0}""", options.Install);
                            if (options.Quiet) cmdsi.Arguments += @" /quiet";
                            if (options.NoRestart) cmdsi.Arguments += @" /norestart";
                            if (!String.IsNullOrEmpty(options.LogFile))
                            {
                                cmdsi.Arguments += String.Format(@" /l*+ ""{0}""", options.LogFile);
                                Console.Error.Flush();
                                Console.Error.Close();
                            }
                            if (!String.IsNullOrEmpty(options.CmdLine)) cmdsi.Arguments += String.Format(@" {0}", options.CmdLine);

                            Process cmd = Process.Start(cmdsi);
                            cmd.WaitForExit();

                            if (!String.IsNullOrEmpty(options.LogFile))
                            {
                                Console.SetError(new StreamWriter(options.LogFile, true, Encoding.Unicode));
                            }

                            writeLog(String.Format(@"Result code is {0}", cmd.ExitCode));
                            Environment.ExitCode = cmd.ExitCode;
                        }
                        else
                        {
                            if (options.Verbose)
                            {
                                writeLog(String.Format(@"Package {0} not found", options.Package));
                            }
                        }
                    }
                }
                Console.Error.Flush();
                Console.Error.Close();
            }
            catch (Exception e)
            {
                Environment.ExitCode = 1;
                if (options.Verbose)
                {
                    Console.Error.WriteLine(e.ToString());
                }
            }
        }
    }

    // Define a class to receive parsed values
    class Options
    {
        [Option('p', "package", Required = true, HelpText = "Name of msi package.")]
        public string Package { get; set; }

        [Option('v', "verbose", DefaultValue = false, HelpText = "Prints debug messages to standard error.")]
        public bool Verbose { get; set; }

        [Option('x', "uninstall", DefaultValue = false, HelpText = "Uninstall software if found.", MutuallyExclusiveSet = "install")]
        public bool Uninstall { get; set; }

        [Option('i', "install", HelpText = "Install or repair software if found.", MutuallyExclusiveSet = "install")]
        public String Install { get; set; }

        [Option('q', "quiet", DefaultValue = false, HelpText = "Force MsiExec to run in silent mode.")]
        public bool Quiet { get; set; }

        [Option('l', "logfile", Required = false, HelpText = "Log messages to file instead of stderr.")]
        public String LogFile { get; set; }

        [Option('n', "norestart", DefaultValue = false, HelpText = "Do not restart at end of uninstall.")]
        public bool NoRestart { get; set; }

        [Option('c', "cmdline", Required = false, HelpText = "Optional parameters to pass to MsiExec")]
        public string CmdLine { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }

    class SoftwareSearch
    {
        public Boolean? Installed = false;
        public String RegisteryKey = null;
        public String GUID = null;
        public RegistryView RegView = RegistryView.Default;
        public Exception Exception = null;

        static bool is64BitProcess = (IntPtr.Size == 8);
        static bool is64BitOperatingSystem = is64BitProcess || InternalCheckIsWow64();

        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWow64Process([In] IntPtr hProcess, [Out] out bool wow64Process);

        public static bool InternalCheckIsWow64()
        {
            if ((Environment.OSVersion.Version.Major == 5 && Environment.OSVersion.Version.Minor >= 1) ||
                Environment.OSVersion.Version.Major >= 6)
            {
                using (Process p = Process.GetCurrentProcess())
                {
                    bool retVal;
                    if (!IsWow64Process(p.Handle, out retVal))
                    {
                        return false;
                    }
                    return retVal;
                }
            }
            else
            {
                return false;
            }
        }

        public SoftwareSearch(String Software)
        {
            try
            {
                String[] registeryLocation = { @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" };
                List<RegistryView> RegView = new List<RegistryView>();

                if (InternalCheckIsWow64())
                {
                    RegView.Add(RegistryView.Registry64);
                }
                RegView.Add(RegistryView.Registry32);

                foreach (RegistryView regView in RegView)
                {
                    RegistryKey registryKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, regView);

                    foreach (String location in registeryLocation)
                    {
                        using (RegistryKey key = registryKey.OpenSubKey(location))
                        {
                            foreach (string subkey_name in key.GetSubKeyNames())
                            {
                                using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                                {
                                    if(subkey != null)
                                    {
                                        if (subkey.GetValue("DisplayName") != null)
                                        {
                                            if (subkey.GetValue("DisplayName").ToString().CompareTo(Software) == 0)
                                            {
                                                this.Installed = true;
                                                this.RegisteryKey = subkey.Name;
                                                this.GUID = subkey_name;
                                                this.RegView = regView;
                                                return;
                                            }
                                        }
                                    }
                                    
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {
                this.Installed = null;
                this.RegisteryKey = null;
                this.GUID = null;
                this.Exception = e;
            }
        }
    }
}

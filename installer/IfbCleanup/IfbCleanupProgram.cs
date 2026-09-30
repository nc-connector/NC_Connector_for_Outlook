// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using NcTalkOutlookAddIn.Services;

namespace NcConnectorOutlookInstaller
{
    // MSI-only cleanup with a protected, short-lived rollback journal.
    internal static class IfbCleanupProgram
    {
        private const string JournalRoot = @"Software\NC4OL\Installer\IfbCleanup";
        private const string ProfileListPath =
            @"Software\Microsoft\Windows NT\CurrentVersion\ProfileList";
        private const string OfficePath = @"Software\Microsoft\Office";
        private const string CalendarSuffix = @"\Outlook\Options\Calendar";
        private const string CalendarValue = "FreeBusySearchPath";
        private const string InternetValue = "Read URL";
        private static readonly IntPtr UsersHandle = new IntPtr(unchecked((int)0x80000003));

        private static int Main(string[] args)
        {
            try
            {
                if (args.Length == 3 && args[0] == "check-outlook")
                {
                    string messageKey = !string.IsNullOrEmpty(args[2]) ? "rollback" :
                        IsOutlookRunning() ? "outlook" : null;
                    if (messageKey != null)
                    {
                        int uiLevel;
                        if (int.TryParse(args[1], out uiLevel) && uiLevel >= 3)
                        {
                            MessageBox(IntPtr.Zero,
                                SetupMessage(CultureInfo.CurrentUICulture, messageKey),
                                "NC Connector for Outlook", 0x10);
                        }
                        return 1603;
                    }
                    return 0;
                }

                Guid productCode;
                if (args.Length < 2 || !Guid.TryParse(args[1], out productCode))
                {
                    throw new InvalidOperationException("Invalid cleanup arguments.");
                }
                string transactionPath = JournalRoot + "\\" + productCode.ToString("N");
                if (args[0] == "clean" && args.Length == 3)
                {
                    if (!string.IsNullOrEmpty(args[2]))
                    {
                        throw new InvalidOperationException("MSI rollback must be enabled.");
                    }
                    RequireOutlookClosed();
                    EnableHivePrivileges();
                    CleanProfiles(transactionPath);
                }
                else if (args[0] == "rollback" && args.Length == 2)
                {
                    EnableHivePrivileges();
                    RestoreProfiles(transactionPath);
                }
                else if (args[0] == "commit" && args.Length == 2)
                {
                    DeleteJournal(transactionPath);
                }
                else
                {
                    throw new InvalidOperationException("Invalid cleanup mode.");
                }
                return 0;
            }
            catch (Exception ex)
            {
                // Registry values contain process secrets; never print values or exception messages.
                Console.Error.WriteLine("NC Connector IFB setup action failed: {0} (0x{1:X8}).",
                    ex.GetType().Name, ex.HResult);
                return 1603;
            }
        }

        internal static bool IsUserSid(string value)
        {
            return value != null && Regex.IsMatch(value,
                @"\AS-1-(?:5-21-[0-9]+-[0-9]+-[0-9]+-[0-9]+|12-1-[0-9]+-[0-9]+-[0-9]+-[0-9]+)\z",
                RegexOptions.CultureInvariant);
        }

        internal static string SetupMessage(CultureInfo culture, string key)
        {
            var messages = new XmlDocument { XmlResolver = null };
            using (Stream stream = typeof(IfbCleanupProgram).Assembly.GetManifestResourceStream(
                "NcConnectorIfbCleanup.SetupMessages.xml"))
            {
                messages.Load(stream);
            }
            string locale = culture.Name.Replace('-', '_');
            XmlElement language = messages.DocumentElement.SelectSingleNode(
                "language[@id='" + locale + "']") as XmlElement;
            if (language == null)
            {
                language = messages.DocumentElement.SelectSingleNode(
                    "language[@id='" + culture.TwoLetterISOLanguageName + "']") as XmlElement;
            }
            if (language == null)
            {
                language = (XmlElement)messages.DocumentElement.SelectSingleNode("language[@id='en']");
            }
            return language[key].InnerText;
        }

        internal static bool IsOfficeVersion(string value)
        {
            return value != null && Regex.IsMatch(value,
                @"\A[0-9]{1,3}\.[0-9]{1,3}\z", RegexOptions.CultureInvariant);
        }

        internal static bool IsOutlookRunning()
        {
            Process[] processes = Process.GetProcessesByName("OUTLOOK");
            try
            {
                return processes.Length != 0;
            }
            finally
            {
                foreach (Process process in processes)
                {
                    process.Dispose();
                }
            }
        }

        private static void RequireOutlookClosed()
        {
            if (IsOutlookRunning())
            {
                throw new InvalidOperationException("Close Outlook in every Windows session.");
            }
        }

        private static RegistrySecurity CreateJournalSecurity()
        {
            var security = new RegistrySecurity();
            security.SetAccessRuleProtection(true, false);
            foreach (WellKnownSidType type in new[] {
                WellKnownSidType.LocalSystemSid, WellKnownSidType.BuiltinAdministratorsSid })
            {
                security.AddAccessRule(new RegistryAccessRule(
                    new SecurityIdentifier(type, null), RegistryRights.FullControl,
                    InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow));
            }
            return security;
        }

        private static void CleanProfiles(string transactionPath)
        {
            // Recover an interrupted maintenance run before starting a new journal.
            RestoreProfiles(transactionPath);
            using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                using (RegistryKey journal = machine.CreateSubKey(transactionPath,
                    RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None, CreateJournalSecurity()))
                {
                    foreach (KeyValuePair<string, string> profile in GetProfiles())
                    {
                        RequireOutlookClosed();
                        WithUserHive(profile.Key, profile.Value, false,
                            delegate(RegistryKey hive) { CleanUserHive(hive, profile.Key, journal); });
                    }
                    RequireOutlookClosed();
                }
            }
        }

        internal static void CleanUserHive(RegistryKey userHive, string sid, RegistryKey journal)
        {
            if (!IsUserSid(sid))
            {
                throw new InvalidOperationException("Invalid user SID.");
            }
            using (RegistryKey office = userHive.OpenSubKey(OfficePath, false))
            {
                if (office == null)
                {
                    return;
                }
                foreach (string version in office.GetSubKeyNames())
                {
                    if (!IsOfficeVersion(version))
                    {
                        continue;
                    }
                    CleanValue(userHive, sid, version, false, journal);
                    CleanValue(userHive, sid, version, true, journal);
                }
            }
        }

        private static string TargetPath(string version, bool internet)
        {
            return OfficePath + "\\" + version + CalendarSuffix
                + (internet ? @"\Internet Free/Busy" : string.Empty);
        }

        private static void CleanValue(RegistryKey userHive, string sid, string version,
            bool internet, RegistryKey journal)
        {
            using (RegistryKey key = userHive.OpenSubKey(TargetPath(version, internet), true))
            {
                if (key == null)
                {
                    return;
                }
                string name = internet ? InternetValue : CalendarValue;
                string value = key.GetValue(name, null,
                    RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                if (!IfbRegistryEndpoint.IsSearchPath(value))
                {
                    return;
                }
                RegistryValueKind kind = key.GetValueKind(name);
                if (kind != RegistryValueKind.String && kind != RegistryValueKind.ExpandString)
                {
                    return;
                }
                var removed = new RemovedValue(sid, version, internet, kind, value);
                string entryName = Guid.NewGuid().ToString("N");
                journal.SetValue(entryName, removed.Encode(), RegistryValueKind.Binary);
                journal.Flush();

                // Recheck after persisting the snapshot; another writer may have changed the value.
                if (string.Equals(key.GetValue(name, null,
                        RegistryValueOptions.DoNotExpandEnvironmentNames) as string,
                        value, StringComparison.Ordinal) && key.GetValueKind(name) == kind)
                {
                    key.DeleteValue(name, false);
                    key.Flush();
                }
                else
                {
                    journal.DeleteValue(entryName, false);
                }
            }
        }

        private static void RestoreProfiles(string transactionPath)
        {
            using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (RegistryKey journal = machine.OpenSubKey(transactionPath, true))
            {
                if (journal == null)
                {
                    return;
                }
                var sids = new HashSet<string>(StringComparer.Ordinal);
                foreach (string name in journal.GetValueNames())
                {
                    sids.Add(ReadEntry(journal, name).Sid);
                }
                Dictionary<string, string> profiles = GetProfiles();
                Exception failure = null;
                foreach (string sid in sids)
                {
                    try
                    {
                        string path;
                        profiles.TryGetValue(sid, out path);
                        WithUserHive(sid, path, true,
                            delegate(RegistryKey hive) { RestoreUserHive(hive, journal, sid); });
                    }
                    catch (Exception ex)
                    {
                        failure = ex;
                    }
                }
                if (failure != null)
                {
                    throw new InvalidOperationException("IFB rollback is incomplete; its journal was retained.", failure);
                }
            }
            DeleteJournal(transactionPath);
        }

        internal static void RestoreUserHive(RegistryKey userHive, RegistryKey journal, string sid)
        {
            foreach (string entryName in journal.GetValueNames())
            {
                RemovedValue entry = ReadEntry(journal, entryName);
                if (!string.Equals(entry.Sid, sid, StringComparison.Ordinal))
                {
                    continue;
                }
                string name = entry.Internet ? InternetValue : CalendarValue;
                using (RegistryKey key = userHive.OpenSubKey(TargetPath(entry.Version, entry.Internet), true))
                {
                    // Cleanup removes values, not keys. A removed key is an external change too.
                    if (key != null && key.GetValue(name, null,
                        RegistryValueOptions.DoNotExpandEnvironmentNames) == null)
                    {
                        key.SetValue(name, entry.Value, entry.Kind);
                        key.Flush();
                    }
                }
                journal.DeleteValue(entryName, false);
                journal.Flush();
            }
        }

        private static RemovedValue ReadEntry(RegistryKey journal, string name)
        {
            byte[] data = journal.GetValue(name) as byte[];
            if (data == null || data.Length > 1048576)
            {
                throw new InvalidOperationException("Invalid IFB rollback entry.");
            }
            using (var stream = new MemoryStream(data, false))
            using (var reader = new BinaryReader(stream, Encoding.UTF8))
            {
                var entry = new RemovedValue(reader.ReadString(), reader.ReadString(),
                    reader.ReadBoolean(), (RegistryValueKind)reader.ReadInt32(), reader.ReadString());
                if (stream.Position != stream.Length || !IsUserSid(entry.Sid)
                    || !IsOfficeVersion(entry.Version) || !IfbRegistryEndpoint.IsSearchPath(entry.Value)
                    || (entry.Kind != RegistryValueKind.String && entry.Kind != RegistryValueKind.ExpandString))
                {
                    throw new InvalidOperationException("Invalid IFB rollback entry.");
                }
                return entry;
            }
        }

        private static void DeleteJournal(string transactionPath)
        {
            using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                machine.DeleteSubKeyTree(transactionPath, false);
            }
        }

        private static Dictionary<string, string> GetProfiles()
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            using (RegistryKey users = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64))
            {
                foreach (string sid in users.GetSubKeyNames())
                {
                    if (IsUserSid(sid))
                    {
                        result[sid] = null;
                    }
                }
            }
            using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (RegistryKey profiles = machine.OpenSubKey(ProfileListPath, false))
            {
                if (profiles == null)
                {
                    throw new InvalidOperationException("Windows ProfileList is unavailable.");
                }
                foreach (string sid in profiles.GetSubKeyNames())
                {
                    if (!IsUserSid(sid))
                    {
                        continue;
                    }
                    using (RegistryKey profile = profiles.OpenSubKey(sid, false))
                    {
                        string path = profile == null ? null : profile.GetValue("ProfileImagePath") as string;
                        if (string.IsNullOrWhiteSpace(path))
                        {
                            throw new InvalidOperationException("A Windows profile path is unavailable.");
                        }
                        path = Environment.ExpandEnvironmentVariables(path);
                        if (!Path.IsPathRooted(path) || path.StartsWith(@"\\", StringComparison.Ordinal))
                        {
                            throw new InvalidOperationException("A Windows profile path is not local.");
                        }
                        result[sid] = Path.Combine(Path.GetFullPath(path), "NTUSER.DAT");
                    }
                }
            }
            return result;
        }

        private static void WithUserHive(string sid, string hivePath, bool required, Action<RegistryKey> action)
        {
            using (RegistryKey users = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64))
            {
                using (RegistryKey loaded = users.OpenSubKey(sid, true))
                {
                    if (loaded != null)
                    {
                        action(loaded);
                        return;
                    }
                }
                if (string.IsNullOrEmpty(hivePath))
                {
                    throw new InvalidOperationException("The user hive is no longer loaded.");
                }
                try
                {
                    // RegLoadKey creates an empty hive if the file is absent.
                    File.GetAttributes(hivePath);
                }
                catch (FileNotFoundException)
                {
                    if (required) { throw; }
                    return;
                }
                catch (DirectoryNotFoundException)
                {
                    if (required) { throw; }
                    return;
                }
                string mount = "NC4OL_IFB_" + Guid.NewGuid().ToString("N");
                int result = RegLoadKey(UsersHandle, mount, hivePath);
                if (result != 0)
                {
                    // A logon may have loaded the profile between the initial check and RegLoadKey.
                    using (RegistryKey loaded = users.OpenSubKey(sid, true))
                    {
                        if (loaded != null)
                        {
                            action(loaded);
                            return;
                        }
                    }
                    throw new Win32Exception(result);
                }
                try
                {
                    using (RegistryKey hive = users.OpenSubKey(mount, true))
                    {
                        if (hive == null)
                        {
                            throw new InvalidOperationException("The mounted user hive is unavailable.");
                        }
                        action(hive);
                    }
                }
                finally
                {
                    int unloadResult = RegUnLoadKey(UsersHandle, mount);
                    if (unloadResult != 0)
                    {
                        throw new Win32Exception(unloadResult);
                    }
                }
            }
        }

        private static void EnableHivePrivileges()
        {
            SafeAccessTokenHandle token;
            if (!OpenProcessToken(Process.GetCurrentProcess().Handle, 0x28, out token))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            using (token)
            {
                foreach (string name in new[] { "SeBackupPrivilege", "SeRestorePrivilege" })
                {
                    Luid luid;
                    if (!LookupPrivilegeValue(null, name, out luid))
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }
                    var privileges = new TokenPrivileges { Count = 1, Luid = luid, Attributes = 2 };
                    if (!AdjustTokenPrivileges(token, false, ref privileges, 0, IntPtr.Zero, IntPtr.Zero))
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }
                    int error = Marshal.GetLastWin32Error();
                    if (error != 0)
                    {
                        throw new Win32Exception(error);
                    }
                }
            }
        }

        private sealed class RemovedValue
        {
            internal readonly string Sid;
            internal readonly string Version;
            internal readonly bool Internet;
            internal readonly RegistryValueKind Kind;
            internal readonly string Value;

            internal RemovedValue(string sid, string version, bool internet, RegistryValueKind kind, string value)
            {
                Sid = sid;
                Version = version;
                Internet = internet;
                Kind = kind;
                Value = value;
            }

            internal byte[] Encode()
            {
                using (var stream = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(stream, Encoding.UTF8, true))
                    {
                        writer.Write(Sid);
                        writer.Write(Version);
                        writer.Write(Internet);
                        writer.Write((int)Kind);
                        writer.Write(Value);
                    }
                    return stream.ToArray();
                }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Luid { internal uint Low; internal int High; }
        [StructLayout(LayoutKind.Sequential)]
        private struct TokenPrivileges { internal uint Count; internal Luid Luid; internal uint Attributes; }

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegLoadKeyW")]
        private static extern int RegLoadKey(IntPtr root, string subKey, string file);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegUnLoadKeyW")]
        private static extern int RegUnLoadKey(IntPtr root, string subKey);
        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool OpenProcessToken(IntPtr process, uint access, out SafeAccessTokenHandle token);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool LookupPrivilegeValue(string system, string name, out Luid luid);
        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AdjustTokenPrivileges(SafeAccessTokenHandle token,
            [MarshalAs(UnmanagedType.Bool)] bool disableAll, ref TokenPrivileges privileges,
            uint bufferLength, IntPtr previous, IntPtr returnLength);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr window, string text, string caption, uint type);
    }
}

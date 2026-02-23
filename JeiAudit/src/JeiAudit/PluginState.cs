using System;
using System.IO;
using System.Linq;
using System.Text;

namespace JeiAudit
{
    internal static class PluginState
    {
        private const string StateFolderName = "JeiAudit";
        private const string LastXmlPathFileName = "last_model_checker_xml.txt";
        private const string LastOutputFolderFileName = "last_output_folder.txt";

        internal static string GetStateFolder()
        {
            string stateFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                StateFolderName);

            Directory.CreateDirectory(stateFolder);
            return stateFolder;
        }

        internal static void SaveLastXmlPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            string stateFile = Path.Combine(GetStateFolder(), LastXmlPathFileName);
            File.WriteAllText(stateFile, path, Encoding.UTF8);
        }

        internal static string LoadLastXmlPath()
        {
            string stateFile = Path.Combine(GetStateFolder(), LastXmlPathFileName);
            if (!File.Exists(stateFile))
            {
                return string.Empty;
            }

            return File.ReadAllText(stateFile, Encoding.UTF8).Trim();
        }

        internal static string ResolveDefaultXmlPath(string lastPath, string documentDir)
        {
            if (!string.IsNullOrWhiteSpace(lastPath) && File.Exists(lastPath))
            {
                return lastPath;
            }

            if (!string.IsNullOrWhiteSpace(documentDir) && Directory.Exists(documentDir))
            {
                string first = Directory
                    .EnumerateFiles(documentDir, "*MCSettings*.xml", SearchOption.TopDirectoryOnly)
                    .FirstOrDefault();

                if (!string.IsNullOrWhiteSpace(first))
                {
                    return first;
                }
            }

            return string.Empty;
        }

        internal static void SaveLastOutputFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return;
            }

            string stateFile = Path.Combine(GetStateFolder(), LastOutputFolderFileName);
            File.WriteAllText(stateFile, folderPath, Encoding.UTF8);
        }

        internal static string LoadLastOutputFolder()
        {
            string stateFile = Path.Combine(GetStateFolder(), LastOutputFolderFileName);
            if (!File.Exists(stateFile))
            {
                return string.Empty;
            }

            return File.ReadAllText(stateFile, Encoding.UTF8).Trim();
        }

        internal static string FindLatestOutputFolder()
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (!Directory.Exists(desktop))
            {
                return string.Empty;
            }

            string latest = Directory
                .EnumerateDirectories(desktop, "JeiAudit_*", SearchOption.TopDirectoryOnly)
                .Select(path => new DirectoryInfo(path))
                .OrderByDescending(info => info.LastWriteTimeUtc)
                .Select(info => info.FullName)
                .FirstOrDefault();

            return latest ?? string.Empty;
        }
    }
}

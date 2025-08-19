using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ProgramKontrol //buranın özet şu önce denetim masası program kaldırdan bakıyor sonra programfiles sonra start up
{                         //mesal duckduckgo denetim masaında yok ama programfilesda var
    public class ProgramKontrol
    {
        public static bool IsProgramInstalled(string programAdi)
        {
            if (string.IsNullOrWhiteSpace(programAdi))
                return false;

            string normalAranan = programAdi.ToLower().Replace(" ", "");

            // 1. Denetim Masasından bakıyo
            string[] registryPaths =
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (string path in registryPaths)
            {
                using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(path))
                {
                    if (key == null) continue;

                    foreach (string subKeyName in key.GetSubKeyNames())
                    {
                        using (RegistryKey? subKey = key.OpenSubKey(subKeyName))
                        {
                            string? displayName = subKey?.GetValue("DisplayName") as string;
                            if (!string.IsNullOrEmpty(displayName))
                            {
                                string normalized = displayName.ToLower().Replace(" ", "");
                                if (normalized.Contains(normalAranan))
                                    return true;
                            }
                        }
                    }
                }
            }

            // 2. Program Files ve Program Files (x86) klasörüne bakıyor
            string[] programDirs =
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            };

            foreach (string baseDir in programDirs)
            {
                try
                {
                    if (!Directory.Exists(baseDir)) continue;

                    foreach (string dir in Directory.GetDirectories(baseDir))
                    {
                        string klasorAdi = Path.GetFileName(dir).ToLower().Replace(" ", "");
                        if (klasorAdi.Contains(normalAranan))
                            return true;
                    }
                }
                catch
                {
                    // Hataları yoksay amaaan gereksizdir
                }
            }

            // 3. Başlat Menüsü^nde var mı yok mu diye bakıyo (tek tırnak nasıl atıyoduk ?)
            try
            {
                string startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.Programs);

                var entries = Directory.EnumerateFileSystemEntries(startMenuPath, "*", SearchOption.AllDirectories);

                foreach (string entry in entries)
                {
                    string isim = Path.GetFileNameWithoutExtension(entry).ToLower().Replace(" ", "");
                    if (isim.Contains(normalAranan))
                        return true;
                }
            }
            catch
            {
                // Hataları yoksay gene
            }

            return false;
        }

        public static Dictionary<string, bool> ProgramListesiKontrol(List<string> programlar)
        {
            var sonuc = new Dictionary<string, bool>();

            foreach (string program in programlar)
            {
                sonuc[program] = IsProgramInstalled(program);
            }

            return sonuc;
        }
    }
}

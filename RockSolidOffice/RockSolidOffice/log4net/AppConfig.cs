using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;

namespace RockSolidOffice
{
    public class AppConfig
    {
        public static FileInfo GetFile()
        {
            string path = ConvertFromFileProtocol(Assembly.GetExecutingAssembly().CodeBase);
            path = path + ".config";
            return new FileInfo(path);
        }

        public static string ConvertFromFileProtocol(string path)
        {
            path = path.ToLower();
            path = path.Replace("file:///", "");
            return path.Replace("/", "\\");
        }
    }
}

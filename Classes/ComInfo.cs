using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ComActivator.Classes
{
    class ComInfo
    {
        private ComInfo()
        { }

        public string CodeBaseUri { get; private set; }

        public string CodeBaseLocal { get; private set; }

        public string ClassName { get; private set; }

        public string ClsId { get; private set; }

        public string AppDomain { get; private set; }

        public static ComInfo GetComInfoFromClsid(Guid clsId)
        {
            ComInfo result = new ComInfo();

            using (RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default))
            using (RegistryKey codeBaseKey = key.OpenSubKey(@"CLSID\" + clsId.ToString("B") + @"\InProcServer32"))
            {
                result.CodeBaseUri = codeBaseKey.GetValue("CodeBase")?.ToString();
                result.ClassName = codeBaseKey.GetValue("Class")?.ToString();
                result.CodeBaseLocal = new Uri(result.CodeBaseUri).LocalPath;
                result.AppDomain = codeBaseKey.GetValue("AppDomain")?.ToString();
                return result;
            }
        }
    }
}

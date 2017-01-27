using ComActivator.Classes;
using ComActivator.Helpers;
using ComActivator.Interfaces;
using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ComActivator
{
    public static class Exports
    {
        static Dictionary<string, AppDomain> appDomains = new Dictionary<string, AppDomain>();

        private static string CurrentFolder()
        {
            return Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;
        }

        internal static AppDomain NewAppDomain(Guid clsId)
        {
            string clsIdName = clsId.ToString("B");
            if (appDomains.ContainsKey(clsIdName))
                return appDomains[clsIdName];

            //http://stackoverflow.com/a/13355702/864414
            AppDomainSetup domaininfo = new AppDomainSetup();
            domaininfo.ApplicationBase = CurrentFolder();
            var curDomEvidence = AppDomain.CurrentDomain.Evidence;

            string appDomainName = ComInfo.GetComInfoFromClsid(clsId).AppDomain ??
                    clsId.ToString("B");

            //The AppDomain name should be the DLL's name.
            AppDomain newDomain = AppDomain.CreateDomain(
                appDomainName, 
                curDomEvidence, domaininfo);

            appDomains.Add(clsIdName, newDomain);

            return newDomain;
        }
        
        [DllExport]
        public static uint DllGetClassObject(Guid rclsid, Guid riid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;

            try
            {
                if (riid == new Guid(ComHelper.IID_IClassFactory))
                {
                    Type type = typeof(Classes.ComClassFactory);

                    object _FactoryInstance = NewAppDomain(rclsid).CreateInstanceAndUnwrap(
                           type.Assembly.FullName,
                           type.FullName, 
                           true,
                           BindingFlags.Default,
                           null,
                           new object[] { rclsid },
                           null,
                           null);

                    //Call to DllClassObject is requesting IClassFactory.
                    ppv = Marshal.GetComInterfaceForObject(_FactoryInstance, typeof(IClassFactory));
                    return 0; //S_OK
                }
                else
                    return ComHelper.E_NOINTERFACE; //CLASS_E_CLASSNOTAVAILABLE
            }
            catch
            {
                return ComHelper.E_NOINTERFACE; //CLASS_E_CLASSNOTAVAILABLE
            }
        }

        //[DllExport]
        public static int DllCanUnloadNow()
        {
            appDomains.Clear();
            return 0; //S_OK
        }
    }
}

using ComActivator.Helpers;
using ComActivator.Interfaces;
using RGiesecke.DllExport;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ComActivator
{
    public static class Exports
    {
        static object _FactoryInstance = null;
        static string _ComClassId = String.Empty;

        private static string CurrentFolder()
        {
            return Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;
        }

        internal static Lazy<AppDomain> _AppDomain = new Lazy<AppDomain>(() => {
            //http://stackoverflow.com/a/13355702/864414
            AppDomainSetup domaininfo = new AppDomainSetup();
            domaininfo.ApplicationBase = CurrentFolder();
            var curDomEvidence = AppDomain.CurrentDomain.Evidence;
            AppDomain newDomain = AppDomain.CreateDomain(
                _ComClassId, 
                curDomEvidence, domaininfo);
             
            return newDomain;
        });
        
        [DllExport]
        public static uint DllGetClassObject(Guid rclsid, Guid riid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            _ComClassId = rclsid.ToString("B");

            try
            {
                if (riid == new Guid(ComHelper.IID_IClassFactory))
                {
                    Type type = typeof(Classes.ComClassFactory);

                    _FactoryInstance = _AppDomain.Value.CreateInstanceAndUnwrap(
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
            return 0; //S_OK
        }
    }
}

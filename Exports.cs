using ComActivator.Classes;
using ComActivator.Helpers;
using ComActivator.Interfaces;
using Microsoft.Win32.SafeHandles;
//using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
namespace ComActivator
{
    public static class Exports
    {
        static Dictionary<string, AppDomain> appDomains = new Dictionary<string, AppDomain>();
        static object _ClassFactoryInstance;

        static Exports()
        {
            /*
            using (Stream assmStr = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ComActivator.DllExport.dll"))
            using (BinaryReader binRdr = new BinaryReader(assmStr))
            {
                byte[] rawAssm = new byte[assmStr.Length];
                int nRead = 0;

                nRead = assmStr.Read(rawAssm, 0, (int)assmStr.Length);

                var assm = Assembly.Load(rawAssm);
            }
            */
        }

        private static string CurrentFolder()
        {
            return Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;
        }

        internal static AppDomain NewAppDomain(ComInfo comInfo)
        {
            string clsIdName = comInfo.ClsId.ToString("B");

            //Have we already created an AppDomain for this class id in the loaded module?
            //If so, return it.
            if (appDomains.ContainsKey(clsIdName))
                return appDomains[clsIdName];

            //http://stackoverflow.com/a/13355702/864414
            AppDomainSetup domaininfo = new AppDomainSetup();
            domaininfo.ApplicationBase = Path.GetDirectoryName(comInfo.CodeBaseLocal); // CurrentFolder();
            var curDomEvidence = AppDomain.CurrentDomain.Evidence;

            string appDomainName = comInfo.AppDomain ?? clsIdName;

            //The AppDomain name should be the DLL's name.
            AppDomain newDomain = AppDomain.CreateDomain(
                appDomainName, 
                curDomEvidence, domaininfo);

            appDomains.Add(clsIdName, newDomain);

            return newDomain;
        }
        
        
        /// <summary>
        /// This is a "native" DLL entry point for creating a COM object. Note that the first two
        /// parameters are Guids, but for some reason they are not marshalled properly for x86.  Thus
        /// we are manually marshalling them, which seems to work.
        /// </summary>
        /// <param name="prclsid">Guid of the CLSID to instantiate</param>
        /// <param name="priid">Guid of the interface to return</param>
        /// <param name="ppv">Out pointer to the IClassFactory object which creates instances of the COM object.</param>
        /// <returns></returns>
        [DllExport(CallingConvention.StdCall, ExportName = "DllGetClassObject")]
        //[LoaderOptimization(LoaderOptimization.MultiDomain)]
        public static IntPtr DllGetClassObject(IntPtr prclsid, IntPtr priid, out IntPtr ppv)
        {
            Guid rclsid = (Guid)Marshal.PtrToStructure(prclsid, typeof(Guid));
            Guid riid = (Guid)Marshal.PtrToStructure(priid, typeof(Guid));

            ppv = IntPtr.Zero;

            try
            {

                ComInfo comInfo = ComInfo.GetComInfoFromClsid(rclsid);

                if (ComHelper.GuidIsIn(riid,
                        comInfo.ClsId,
                        ComHelper.IID_IClassFactory,
                        ComHelper.IID_IUnknown,
                        ComHelper.IID_IDispatch,
                        ComHelper.IID_IOleObject))
                {
                    Type type = typeof(Classes.ComClassFactory);

                    //AddressMarshaler objAddress = new AddressMarshaler();

                    //Pass a memory address to write to so we can marshal the COM interface.
                    IntPtr hComObjAddress = Marshal.AllocHGlobal(IntPtr.Size);

                    _ClassFactoryInstance = NewAppDomain(comInfo).CreateInstanceFromAndUnwrap(
                            type.Assembly.Location,  //Changed from type.Assembly.FullName - specify path instead of class name
                            type.FullName,
                            true,
                            BindingFlags.Default,
                            null,
                            new object[] { rclsid, hComObjAddress },
                            null,
                            null);


                    /*
                    //http://stackoverflow.com/a/2132939/864414
                    object _FactoryInstance = NewAppDomain(comInfo).
                        CreateInstanceFromAndUnwrap(
                            Assembly.GetExecutingAssembly().Location,
                            typeof(ComClassFactory).FullName
                        );
                   */


                    ////Get the requested Guid
                    //Guid requestedGuid = new Guid();
                    //Marshal.PtrToStructure(priid, requestedGuid);

                    ////Class factory is coming out
                    //Guid newGuid = new Guid(ComHelper.IID_IClassFactory);
                    //Marshal.PtrToStructure(priid, newGuid);

                    IntPtr comMarshalDirect = Marshal.GetComInterfaceForObject(_ClassFactoryInstance, typeof(IClassFactory));
                    IntPtr comObjectAddr = Marshal.ReadIntPtr(hComObjAddress);

                    //Call to DllClassObject is requesting IClassFactory.
                    ppv = comObjectAddr;

                    //ppv = Marshal.UnsafeAddrOfPinnedArrayElement(new[]{ _FactoryInstance }, 0);
                    //ppv = Marshal.GetIUnknownForObject(_FactoryInstance);
                    //int refCount = Marshal.Release(ppv);

                    Marshal.FreeHGlobal(hComObjAddress);

                    return IntPtr.Zero; //S_OK
                }
                else
                {
                    return new IntPtr(ComHelper.E_NOINTERFACE); //CLASS_E_CLASSNOTAVAILABLE
                }
            }
            catch
            {
                return new IntPtr(ComHelper.E_NOINTERFACE); //CLASS_E_CLASSNOTAVAILABLE
            }
        }

        [DllExport(CallingConvention.StdCall, ExportName = "DllCanUnloadNow")]
        public static int DllCanUnloadNow()
        {
            GC.KeepAlive(_ClassFactoryInstance);
            foreach (var appDomain in appDomains)
            {
                AppDomain.Unload(appDomain.Value);
            }
            appDomains.Clear();
            return 0; //S_OK
        }
    }
}

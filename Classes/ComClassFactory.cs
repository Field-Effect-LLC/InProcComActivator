using ComActivator.Helpers;
using ComActivator.Interfaces;
using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ComActivator.Classes
{
    [ComVisible(true)]
    [Serializable]
    public class ComClassFactory : MarshalByRefObject, IClassFactory
    {
        //private static object _ComObject = null;
        private string _BasePath = String.Empty;
        Type _ComClass;

        public ComClassFactory(Guid rclsid, IntPtr hComObjAddress)
        {
            Assembly loadedAssm = null;

            //We know the type being requested, because we have rclsid.
            //We have to look in the registry to figure out what the codebase
            //is (you *did* register your COM with the /codebase option, right? :)

            ComInfo comInfo = ComInfo.GetComInfoFromClsid(rclsid);
            loadedAssm = Assembly.LoadFile(comInfo.CodeBaseLocal);
            _BasePath = AppDomain.CurrentDomain.BaseDirectory;
            _ComClass = loadedAssm.GetType(comInfo.ClassName);

            IntPtr pComAddr = Marshal.GetComInterfaceForObject(this, typeof(IClassFactory));

            Marshal.WriteIntPtr(hComObjAddress, pComAddr);
        }

        public IClassFactory GetSelf()
        {
            return (IClassFactory)this;
        }

        private static string CurrentFolder(Assembly assm)
        {
            return Directory.GetParent(assm.Location).FullName;
        }

        public uint CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
        {
           
            ppvObject = IntPtr.Zero;

            if (pUnkOuter != IntPtr.Zero)
                return ComHelper.CLASS_E_NOAGGREGATION;

            //Get Guid attribute
            //AARGH! This bug was hard to track down.  We can't use the GetCustomAttribute *extension*
            //because there's a version mismatch.  A Windows 7 installation doesn't appear to have
            //this implemented in mscorlib.dll, even though it appears to work on my dev machine.
            //When deploying, it crashes.

            Guid guidAttr = _ComClass.GUID;

            if (guidAttr == null)
                return ComHelper.E_NOINTERFACE;

            if (!ComHelper.GuidIsIn(riid,
                  guidAttr,
                  ComHelper.IID_IUnknown,
                  ComHelper.IID_IDispatch,
                  ComHelper.IID_IOleObject))
                    return ComHelper.E_NOINTERFACE;

            //Instantiate the object with its default constructor
            ConstructorInfo defaultConstr = _ComClass
                .GetConstructor(Type.EmptyTypes);

            object _ComObject = defaultConstr.Invoke(null);

            if (_ComObject == null)
                return ComHelper.E_NOINTERFACE;

            ppvObject = Marshal.GetIUnknownForObject(_ComObject);

            //Get the correct interface. This is especially important when IOleObject is 
            //being requested, because I think that .NET implements this interface implicitly
            var result = Marshal.QueryInterface(ppvObject, ref riid, out ppvObject);
            return 0; //S_OK
        }

        public uint LockServer(bool fLock)
        {
            return 0; //S_OK
        }
    }
}

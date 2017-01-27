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
        private static object _ComObject = null;
        private Guid _ClassId;
        private string _CodeBase = String.Empty;
        private string _ClassName = String.Empty;
        private string _BasePath = String.Empty;
        Type _ComClass;

        static ComClassFactory()
        {
            
        }

        public ComClassFactory(Guid rclsid)
        {
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolver;
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += AssemblyResolver;

            this._ClassId = rclsid;
            Assembly loadedAssm = null;

            //We know the type being requested, because we have rclsid.
            //We have to look in the registry to figure out what the codebase
            //is (you *did* register your COM with the /codebase option, right? :)

            ComInfo comInfo = ComInfo.GetComInfoFromClsid(rclsid);
            loadedAssm = Assembly.LoadFile(comInfo.CodeBaseLocal);
            _BasePath = Directory.GetParent(comInfo.CodeBaseLocal).FullName;
            //LoadReferencedAssemblies(AppDomain.CurrentDomain, loadedAssm);
            _ComClass = loadedAssm.GetType(comInfo.ClassName);
        }

        private static string CurrentFolder(Assembly assm)
        {
            return Directory.GetParent(assm.Location).FullName;
        }

        private Assembly AssemblyResolver(object sender, ResolveEventArgs args)
        {
            Assembly foundAssembly = null;

            //Is the assembly already loaded?
            foundAssembly = AppDomain.CurrentDomain.GetAssemblies()
                .Where<Assembly>((curAssm) => curAssm.FullName == args.Name)
                .FirstOrDefault<Assembly>();

            if (foundAssembly != null)
                return foundAssembly;

            AssemblyName assmName = new AssemblyName(args.Name);

            string potentialPath = Path.Combine(_BasePath, assmName.Name) + ".dll";

            if (File.Exists(potentialPath))
            {
                foundAssembly = Assembly.LoadFrom(potentialPath);
            }

            //Either return the found assembly or null if we couldn't find.
            return foundAssembly;
        }

        public uint CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
        {
            ppvObject = IntPtr.Zero;

            //Get Guid attribute
            GuidAttribute guidAttr = (GuidAttribute)_ComClass.GetCustomAttribute(typeof(GuidAttribute), true);

            if (guidAttr == null)
                return ComHelper.E_NOINTERFACE;

            if (riid != new Guid(guidAttr.Value) &&
                riid != new Guid(ComHelper.IID_IOleObject) &&
                riid != new Guid(ComHelper.IID_IDispatch) &&
                riid != new Guid(ComHelper.IID_IUnknown))
                return ComHelper.E_NOINTERFACE;

            //Instantiate the object with its default constructor

            ConstructorInfo defaultConstr = _ComClass
                .GetConstructor(Type.EmptyTypes);

            _ComObject = defaultConstr.Invoke(null);

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

        private static void LoadAssembly(AppDomain domain, AssemblyName reqAssm, string path)
        {
            bool isLoaded = domain.GetAssemblies()
                .Any<Assembly>(i => i.FullName == reqAssm.FullName);

            if (!isLoaded)
            {
                Assembly curAssm = null;
                try
                {
                    //Load from the GAC
                    curAssm = domain.Load(reqAssm);
                }
                catch
                {
                    //Not found in the GAC
                    string assmName = String.Empty;
                    string assmExt = Path.GetExtension(reqAssm.Name.ToLower());

                    if (!new[] { ".dll", ".resources" }
                        .Any<string>((ext) => assmExt.Equals(ext, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        assmName = reqAssm.Name + ".dll";
                    }
                    var curPath = path;
                    var rawAssembly = File.ReadAllBytes(Path.Combine(curPath, assmName));
                    curAssm = domain.Load(rawAssembly);
                }
            }
        }

        private static void LoadReferencedAssemblies(AppDomain domain, Assembly assm)
        {
            foreach (var reqAssm in assm.GetReferencedAssemblies())
            {
                LoadAssembly(domain, reqAssm, Directory.GetParent(assm.Location).FullName);
            }
        }
    }
}

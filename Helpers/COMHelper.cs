using System;
using System.Linq;

namespace ComActivator.Helpers
{
    class ComHelper
    {
        /// <summary>
        /// Interface Id of IID_IClassFactory
        /// </summary>
        public const string IID_IClassFactory =
        "00000001-0000-0000-C000-000000000046";

        /// <summary>
        /// Interface Id of IUnknown
        /// </summary>
        public const string IID_IUnknown =
            "00000000-0000-0000-C000-000000000046";

        /// <summary>
        /// Interface Id of IDispatch
        /// </summary>
        public const string IID_IDispatch =
            "00020400-0000-0000-C000-000000000046";

        /// <summary>
        /// Interface Id of IOleObject
        /// </summary>
        public const string IID_IOleObject =
            "00000112-0000-0000-c000-000000000046";

        /// <summary>
        /// Interface ID for IOleClientSite
        /// </summary>
        public const string IID_IOleClientSite =
            "00000118-0000-0000-C000-000000000046";

        /// <summary>
        /// Class does not support aggregation (or class object is remote)
        /// </summary>
        public const uint CLASS_E_NOAGGREGATION = unchecked((uint)0x80040110);

        /// <summary>
        /// No such interface supported
        /// </summary>
        public const uint E_NOINTERFACE = unchecked((uint)0x80004002);

        /// <summary>
        /// Given an object, either string or Guid, return a Guid.
        /// </summary>
        /// <param name="guidObj">The object to convert</param>
        /// <returns></returns>
        private static Guid GuidFromObject(object guidObj)
        {
            Guid resultantGuid = Guid.Empty;
            if (guidObj is String)
            {
                resultantGuid = new Guid((String)guidObj);
            }
            else if (guidObj is Guid)
            {
                resultantGuid = (Guid)guidObj;
            }
            return resultantGuid;
        }

        /// <summary>
        /// Convenience method to compare a Guid against several other Guids.
        /// </summary>
        /// <param name="interfaceToTest">The Guid to compare (either string or Guid object)</param>
        /// <param name="interfacesToCompare">The Guids to compare against (either strings or Guid objects)</param>
        /// <returns></returns>
        public static bool GuidIsIn(object interfaceToTest, params object[] interfacesToCompare)
        {
            Guid testGuidObj = GuidFromObject(interfaceToTest);

            return interfacesToCompare
                .Any<object>((interfaceGuid)=>
                    testGuidObj == GuidFromObject(interfaceGuid));
        }
    }
}

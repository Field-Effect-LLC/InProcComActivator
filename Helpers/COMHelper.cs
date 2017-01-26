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

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComActivator.Classes
{
    class SafeAddrHandle : SafeHandle
    {
        public SafeAddrHandle(IntPtr handle)
        {

        }

        public override bool IsInvalid
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        protected override bool ReleaseHandle()
        {
            throw new NotImplementedException();
        }
    }
}

// Not used by consoleapp - just as a generator for call.il ildasm
// run build.cmd
using System;

namespace call
{
    public static class IUnknown
    {
        public static uint AddRefRelease(IntPtr pThis)
        {
            return (uint)(pThis);
        }
    }
}

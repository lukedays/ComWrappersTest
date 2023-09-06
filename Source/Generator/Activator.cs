using System.Runtime.InteropServices;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source.Generator;

public partial class Activator
{
    [LibraryImport("ole32.dll")]
    public static partial int CoCreateInstance(
        ref Guid rclsid,
        nint pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out IWshShell ppv
    );

    public static IWshShell ActivateClass(Guid clsid, CLSCTX server)
    {
        var guid = typeof(IWshShell).GUID;

        int hr = CoCreateInstance(ref clsid, nint.Zero, (uint)server, ref guid, out IWshShell obj);
        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }
        return obj;
    }
}

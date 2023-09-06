using System.Runtime.InteropServices;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source.Manual;

public partial class Activator
{
    [LibraryImport("ole32.dll")]
    public static partial int CoCreateInstance(
        ref Guid rclsid,
        nint pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out nint ppv
    );

    public static I ActivateClass<I>(Guid clsid, CLSCTX server)
    {
        var guid = typeof(I).GUID;

        int hr = CoCreateInstance(ref clsid, nint.Zero, (uint)server, ref guid, out nint obj);
        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }
        var cw = new MyComWrappers();
        return (I)cw.GetOrCreateObjectForComInstance(obj, CreateObjectFlags.UniqueInstance);
    }
}

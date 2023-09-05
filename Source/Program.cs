using System.Runtime.InteropServices;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source;

public class Program
{
    public static void Main()
    {
        // This is an adaptation of https://github.com/dotnet/samples/tree/main/core/interop/comwrappers/Tutorial
        // Intro to COM https://learn.microsoft.com/en-us/windows/win32/com/com-technical-overview

        // The CLSID for WScript.Shell (COMView.exe -> CLSID table)
        var clsid = new Guid("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}");

        // The IID for IWshShell (COMView.exe -> TypeLib -> IWshShell interface)
        var iid = IWshShell.IID;

        // COMView.exe -> CLSID table -> Type column -> InProcServer32
        var server = CLSCTX.CLSCTX_INPROC_SERVER;

        // Create pointer to COM object with wrapper. Technically, a RCW
        var shell = ActivateClass<IWshShell>(clsid, iid, server);

        // Call COM functions. Since it's very tedious, we can use source generators to have more functions.
        shell.Run("calc.exe", 1, true);

        Console.WriteLine(shell.ExpandEnvironmentStrings("%USERPROFILE%"));
    }

    [DllImport("ole32.dll")]
    static extern int CoCreateInstance(
        ref Guid rclsid,
        nint pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out nint ppv
    );

    public static I ActivateClass<I>(Guid clsid, Guid iid, CLSCTX server)
    {
        int hr = CoCreateInstance(ref clsid, nint.Zero, (uint)server, ref iid, out nint obj);
        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }
        var cw = new MyComWrappers();
        return (I)cw.GetOrCreateObjectForComInstance(obj, CreateObjectFlags.UniqueInstance);
    }
}

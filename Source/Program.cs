using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source;

public class Program
{
    public static void Main()
    {
        // The CLSID for WScript.Shell (COMView.exe -> CLSID table)
        var clsid = new Guid("{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}");

        // COMView.exe -> CLSID table -> Type column -> InProcServer32
        var server = CLSCTX.CLSCTX_INPROC_SERVER;

        // Using manual ComWrappers
        // Create pointer to COM object with wrapper. Technically, a RCW
        var shell1 = Manual.Activator.ActivateClass<Manual.IWshShell>(clsid, server);

        shell1.Run("calc.exe", 1, true);

        Console.WriteLine(
            "Hello from manual: {0}",
            shell1.ExpandEnvironmentStrings("%USERPROFILE%")
        );

        // Using source generators
        // I wasn't able to use a generic parameter for the LibraryImport call
        var shell2 = Generator.Activator.ActivateClass(clsid, server);

        shell2.Run("calc.exe"); // easier to use optional parameters

        Console.WriteLine(
            "Hello from generator: {0}",
            shell2.ExpandEnvironmentStrings("%USERPROFILE%")
        );
    }
}

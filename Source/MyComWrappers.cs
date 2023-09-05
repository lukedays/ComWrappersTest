using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ComWrappersTest.Source;

unsafe class MyComWrappers : ComWrappers
{
    protected override unsafe ComInterfaceEntry* ComputeVtables(
        object obj,
        CreateComInterfaceFlags flags,
        out int count
    )
    {
        // This is only used for CCWs
        throw new NotImplementedException();
    }

    protected override object CreateObject(nint externalComObject, CreateObjectFlags flags)
    {
        Debug.Assert(flags.HasFlag(CreateObjectFlags.UniqueInstance));

        return new NativeStaticWrapper(externalComObject);
    }

    protected override void ReleaseObjects(IEnumerable objects)
    {
        throw new NotImplementedException();
    }
}

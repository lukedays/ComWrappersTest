using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source.Generator;

[CustomMarshaller(typeof(bool), MarshalMode.Default, typeof(VariantMarshaller))]
[CustomMarshaller(typeof(int), MarshalMode.Default, typeof(VariantMarshaller))]
internal static unsafe class VariantMarshaller
{
    public static nint ConvertToUnmanaged(bool managed)
    {
        var variant = new VARIANT
        {
            vt = VARENUM.VT_BOOL,
            boolVal = managed ? VARIANT_BOOL.VARIANT_TRUE : VARIANT_BOOL.VARIANT_FALSE
        };

        return VariantToPtr(variant);
    }

    public static nint ConvertToUnmanaged(int managed)
    {
        var variant = new VARIANT { vt = VARENUM.VT_I4, lVal = managed };

        return VariantToPtr(variant);
    }

    public static nint ConvertToManaged(nint unmanaged)
    {
        throw new NotImplementedException();
    }

    static nint VariantToPtr(VARIANT str)
    {
        var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(str));
        Marshal.StructureToPtr(str, ptr, false);
        return ptr;
    }
}

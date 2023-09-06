using System.Runtime.InteropServices.Marshalling;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source.Generator;

[CustomMarshaller(typeof(bool), MarshalMode.Default, typeof(BoolToVariant))]
public static class BoolToVariant
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

    public static bool ConvertToManaged(nint unmanaged)
    {
        throw new NotImplementedException();
    }
}

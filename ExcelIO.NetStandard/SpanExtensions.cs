using System.Runtime.InteropServices;
using System.Text;

namespace ExcelIO;

internal static class SpanExtensions
{
    public static unsafe string GetString(this Encoding encoding, ReadOnlySpan<byte> bytes)
    {
        if (bytes.IsEmpty) return string.Empty;
        fixed (byte* ptr = bytes)
            return encoding.GetString(ptr, bytes.Length);
    }

    public static unsafe void Write(this MemoryStream ms, ReadOnlySpan<byte> span)
    {
        if (span.IsEmpty) return;
        var pos = (int)ms.Position;
        var newLen = pos + span.Length;
        if (newLen > ms.Capacity) ms.Capacity = newLen;
        if (newLen > ms.Length) ms.SetLength(newLen);
        var buf = ms.GetBuffer();
        fixed (byte* dst = &buf[pos])
        fixed (byte* src = span)
        {
            UnsafeCopy(dst, src, (uint)span.Length);
        }
        ms.Position = newLen;
    }

    private static unsafe void UnsafeCopy(void* dst, void* src, uint length)
    {
#if NETSTANDARD2_0
        for (uint i = 0; i < length; i++)
            ((byte*)dst)[i] = ((byte*)src)[i];
#else
        System.Buffer.MemoryCopy(src, dst, length, length);
#endif
    }
}

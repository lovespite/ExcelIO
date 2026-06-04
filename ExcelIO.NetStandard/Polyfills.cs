// Polyfill for record init-only setters on netstandard2.0
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}

// Polyfill for Index and Range types used by range operators on netstandard2.0
namespace System
{
    internal readonly struct Index : IEquatable<Index>
    {
        private readonly int _value;
        private readonly bool _fromEnd;
        public Index(int value, bool fromEnd = false) { _value = value; _fromEnd = fromEnd; }
        public int Value => _value;
        public bool IsFromEnd => _fromEnd;
        public static Index Start => new(0);
        public static Index End => new(~0);
        public static Index FromStart(int value) => new(value);
        public static Index FromEnd(int value) => new(~value);
        public int GetOffset(int length) => _fromEnd ? length - _value : _value;
        public override bool Equals(object? obj) => obj is Index other && Equals(other);
        public bool Equals(Index other) => _value == other._value && _fromEnd == other._fromEnd;
        public override int GetHashCode() => HashCode.Combine(_value, _fromEnd);
        public static implicit operator Index(int value) => FromStart(value);
    }

    internal readonly struct Range : IEquatable<Range>
    {
        public Index Start { get; }
        public Index End { get; }
        public Range(Index start, Index end) { Start = start; End = end; }
        public static Range All => new(Index.Start, Index.End);
        public static Range StartAt(Index start) => new(start, Index.End);
        public static Range EndAt(Index end) => new(Index.Start, end);
        public override bool Equals(object? obj) => obj is Range other && Equals(other);
        public bool Equals(Range other) => Start.Equals(other.Start) && End.Equals(other.End);
        public override int GetHashCode() => HashCode.Combine(Start, End);
    }
}

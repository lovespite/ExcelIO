using System.Buffers.Binary;
using System.Text;

namespace ExcelIO.Formula;

/// <summary>
/// Decompiles BIFF8 RPN formula bytecode into A1-notation formula strings.
/// Uses stack-based postfix-to-infix conversion.
/// </summary>
public static class Biff8FormulaReader
{
    // ── Token IDs ──
    private const byte tAdd     = 0x03;
    private const byte tSub     = 0x04;
    private const byte tMul     = 0x05;
    private const byte tDiv     = 0x06;
    private const byte tPower   = 0x07;
    private const byte tConcat  = 0x08;
    private const byte tLt      = 0x09;
    private const byte tLe      = 0x0A;
    private const byte tEq      = 0x0B;
    private const byte tGe      = 0x0C;
    private const byte tGt      = 0x0D;
    private const byte tNe      = 0x0E;
    private const byte tRange   = 0x11;
    private const byte tUplus   = 0x12;
    private const byte tUminus  = 0x13;
    private const byte tPercent = 0x14;
    private const byte tParen   = 0x15;
    private const byte tMissArg = 0x16;
    private const byte tStr     = 0x17;
    private const byte tAttr    = 0x19;
    private const byte tErr     = 0x1C;
    private const byte tBool    = 0x1D;
    private const byte tInt     = 0x1E;
    private const byte tNum     = 0x1F;
    private const byte tArray   = 0x20;
    private const byte tFunc    = 0x22;
    private const byte tFuncVar = 0x23;
    // References (base IDs; may appear with 0x40 and 0x20 bits set)
    private const byte tRef     = 0x24;
    private const byte tArea    = 0x25;
    private const byte tRef3d   = 0x3A;
    private const byte tName    = 0x39;

    /// <summary>
    /// Decompile BIFF8 RPN formula bytes into an A1-notation formula string (with leading =).
    /// Returns empty string if the formula data is empty or unparseable.
    /// </summary>
    public static string Decompile(byte[] formulaBytes)
    {
        if (formulaBytes.Length == 0) return "";

        try
        {
            var stack = new Stack<string>();
            int pos = 0;
            var parenDepth = new Stack<int>();

            while (pos < formulaBytes.Length)
            {
                byte id = formulaBytes[pos];
                int consumed = ParseToken(formulaBytes, pos, id, stack, parenDepth);
                if (consumed == 0) consumed = 1; // skip unknown 1-byte tokens
                pos += consumed;
            }

            if (stack.Count == 0) return "";
            var result = stack.Pop();
            // Clean up remaining paren markers
            result = result.Replace("((", "(").Replace("))", ")");
            return "=" + result;
        }
        catch
        {
            return "";
        }
    }

    private static int ParseToken(byte[] data, int pos, byte id, Stack<string> stack, Stack<int> parenDepth)
    {
        // ── Operators ──
        switch (id)
        {
            case tAdd: return PopBinary(stack, "+");
            case tSub: return PopBinary(stack, "-");
            case tMul: return PopBinary(stack, "*");
            case tDiv: return PopBinary(stack, "/");
            case tPower: return PopBinary(stack, "^");
            case tConcat: return PopBinary(stack, "&");
            case tLt: return PopBinary(stack, "<");
            case tLe: return PopBinary(stack, "<=");
            case tEq: return PopBinary(stack, "=");
            case tGe: return PopBinary(stack, ">=");
            case tGt: return PopBinary(stack, ">");
            case tNe: return PopBinary(stack, "<>");
            case tRange: return PopBinary(stack, ":");
            case tUplus: return PopUnary(stack, "+", true);
            case tUminus: return PopUnary(stack, "-", true);
            case tPercent: return PopUnary(stack, "%", false);

            case tParen:
                parenDepth.Push(stack.Count);
                return 1;

            case tMissArg:
                stack.Push("");
                return 1;

            case tAttr:
                // Attribute byte — skip (volatile flag, etc.)
                if (pos + 1 < data.Length)
                {
                    byte flags = data[pos + 1];
                    int attrLen = 2;
                    if ((flags & 0x10) != 0 && pos + attrLen < data.Length)
                    {
                        // Has additional data
                        attrLen += 2; // skip extra
                    }
                    return attrLen;
                }
                return 1;
        }

        // ── Values ──
        switch (id)
        {
            case tInt:
                if (pos + 3 <= data.Length)
                {
                    short val = BinaryPrimitives.ReadInt16LittleEndian(data.AsSpan(pos + 1, 2));
                    stack.Push(val.ToString());
                }
                return 3;

            case tNum:
                if (pos + 9 <= data.Length)
                {
                    double val = BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(data.AsSpan(pos + 1, 8)));
                    stack.Push(val.ToString("G15"));
                }
                return 9;

            case tStr:
                if (pos + 1 < data.Length)
                {
                    byte len = data[pos + 1];
                    if (pos + 2 + len <= data.Length)
                    {
                        var str = Encoding.GetEncoding(28591).GetString(data, pos + 2, len);
                        stack.Push("\"" + str.Replace("\"", "\"\"") + "\"");
                        return 2 + len;
                    }
                }
                return 1;

            case tBool:
                if (pos + 2 <= data.Length)
                {
                    stack.Push(data[pos + 1] != 0 ? "TRUE" : "FALSE");
                }
                return 2;

            case tErr:
                if (pos + 2 <= data.Length)
                {
                    byte errorCode = data[pos + 1];
                    stack.Push(ErrorName(errorCode));
                }
                return 2;

            case tArray:
                return ParseArray(data, pos, stack);
        }

        // ── Functions (0x22/0x23, or with 0x40 bit: 0x42/0x43/0x62/0x63) ──
        byte baseId = (byte)(id & 0x1F); // mask off upper bits
        if (baseId == (tFunc & 0x1F) || baseId == (tFuncVar & 0x1F))
        {
            if (pos + 2 <= data.Length)
            {
                ushort funcIdx = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + 1, 2));
                string name = FunctionName(funcIdx);
                bool isFixed = baseId == (tFunc & 0x1F);
                byte argCount = (isFixed && pos + 3 <= data.Length) ? data[pos + 3] : (byte)0;
                int consumed = isFixed ? Math.Min(4, data.Length - pos) : Math.Min(3, data.Length - pos);

                if (isFixed && argCount > 0)
                {
                    // Fixed args: pop argCount values from stack
                    var funcArgs = new List<string>();
                    for (int i = 0; i < argCount; i++)
                    {
                        if (stack.Count > 0)
                            funcArgs.Insert(0, stack.Pop());
                    }
                    stack.Push(name + "(" + string.Join(",", funcArgs) + ")");
                }
                else
                {
                    // tFuncVar: pop args since last paren marker, or all remaining stack items
                    var funcArgs = new List<string>();
                    int argsSince;
                    if (parenDepth.Count > 0)
                    {
                        argsSince = stack.Count - parenDepth.Peek();
                    }
                    else
                    {
                        // No paren marker — all remaining items on stack are arguments
                        argsSince = stack.Count;
                    }

                    for (int i = 0; i < argsSince; i++)
                    {
                        if (stack.Count > 0)
                            funcArgs.Insert(0, stack.Pop());
                    }

                    if (parenDepth.Count > 0 && stack.Count >= parenDepth.Peek())
                    {
                        // Remove the opening paren marker
                        stack.Push(name + "(" + string.Join(",", funcArgs) + ")");
                        parenDepth.Pop();
                    }
                    else
                    {
                        stack.Push(name + "(" + string.Join(",", funcArgs) + ")");
                    }
                }

                return consumed;
            }
            return 1;
        }

        // ── References: tRef (0x24/0x44/0x64), tArea (0x25/0x45/0x65) ──
        if ((id & 0x44) == 0x44 || (id & 0x24) == 0x24 || (id & 0x45) == 0x45 || (id & 0x25) == 0x25)
        {
            return ParseCellAreaRef(data, pos, id, stack);
        }

        // Unknown token — skip it
        return 1;
    }

    private static int ParseCellAreaRef(byte[] data, int pos, byte id, Stack<string> stack)
    {
        bool isArea = (id & 0x01) != 0;  // tArea: 0x25, tRef: 0x24

        // Row and column are 2-byte values at pos+1, pos+3
        if (pos + 5 > data.Length) return 1;

        ushort row1 = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + 1, 2));
        ushort col1 = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + 3, 2));

        // Check for relative/absolute flags in the upper bits of id
        bool col1Abs = (id & 0x40) != 0 && (id & 0x08) != 0;
        bool row1Abs = (id & 0x40) != 0 && (id & 0x80) != 0;

        // Actually, the reference encoding is:
        // id bits: 7-6 determine row/col size, bit 5 (0x20) = relative if set
        // For 0x44: row/col are 2 bytes each, both relative
        // For 0x64: row/col are 2 bytes each, both absolute
        // For 0x24: row/col are 2 bytes each with individual abs/rel flags in a following byte

        bool rowAbs, colAbs;
        int consumed;

        if ((id & 0x40) != 0)
        {
            // 0x44 (bit5=0): relative. 0x64 (bit5=1): absolute.
            rowAbs = (id & 0x20) != 0;
            colAbs = rowAbs;
            consumed = 5; // id + 2 bytes row + 2 bytes col = 5
        }
        else
        {
            // 0x24 variant: individual flags in byte at pos+5
            colAbs = true;
            rowAbs = true;
            consumed = 5;
            if (pos + 5 <= data.Length)
            {
                byte flags = data[pos + 5];
                colAbs = (flags & 0x80) == 0;
                rowAbs = (flags & 0x40) == 0;
                consumed = 6; // id + 2 bytes row + 2 bytes col + 1 flag = 6
            }
        }

        string colName = XlFormulaUtil.GetColumnName(col1);
        string rowName = (row1 + 1).ToString();

        string startRef = (colAbs ? "$" : "") + colName + (rowAbs ? "$" : "") + rowName;

        if (!isArea)
        {
            stack.Push(startRef);
            return consumed;
        }

        // tArea: read second cell reference
        if (pos + consumed + 4 > data.Length)
        {
            stack.Push(startRef);
            return consumed;
        }

        ushort row2 = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + consumed, 2));
        ushort col2 = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + consumed + 2, 2));
        consumed += 4;

        // Second cell uses same abs/rel as first
        string col2Name = XlFormulaUtil.GetColumnName(col2);
        string row2Name = (row2 + 1).ToString();
        string endRef = (colAbs ? "$" : "") + col2Name + (rowAbs ? "$" : "") + row2Name;

        stack.Push(startRef + ":" + endRef);
        return consumed;
    }

    private static int ParseArray(byte[] data, int pos, Stack<string> stack)
    {
        // Array constant: next byte is element count, then 8 bytes per element
        if (pos + 1 >= data.Length) return 1;
        byte count = data[pos + 1];
        var parts = new List<string>();
        int consumed = 2;
        for (int i = 0; i < count && consumed + 8 <= data.Length; i++)
        {
            double val = BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(data.AsSpan(pos + consumed, 8)));
            parts.Add(val.ToString("G15"));
            consumed += 8;
        }
        stack.Push("{" + string.Join(",", parts) + "}");
        return consumed;
    }

    private static int PopBinary(Stack<string> stack, string op)
    {
        if (stack.Count < 2) return 1;
        var right = stack.Pop();
        var left = stack.Pop();
        stack.Push(left + op + right);
        return 1;
    }

    private static int PopUnary(Stack<string> stack, string op, bool prefix)
    {
        if (stack.Count < 1) return 1;
        var val = stack.Pop();
        stack.Push(prefix ? op + val : val + op);
        return 1;
    }

    private static string ErrorName(byte code) => code switch
    {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        _ => "#ERR!",
    };

    // ── Function index → name mapping (commonly used functions) ──

    // BIFF8 function indices. Many indices are reserved or undocumented.
    // This table covers the most common ones found in real-world files.
    private static readonly string[] FunctionIndexTable = InitFunctionTable();

    private static string[] InitFunctionTable()
    {
        // Create a large enough table (BIFF8 supports up to ~400+ functions)
        var table = new string[512];
        for (int i = 0; i < table.Length; i++)
            table[i] = $"FUNC_{i}";

        table[0] = "COUNT";
        table[1] = "IF";
        table[2] = "ISNA";
        table[3] = "ISERROR";
        table[4] = "SUM";
        table[5] = "AVERAGE";
        table[6] = "MIN";
        table[7] = "MAX";
        table[8] = "ROW";
        table[9] = "COLUMN";
        table[10] = "NA";
        table[11] = "NPV";
        table[12] = "STDEV";
        table[13] = "DOLLAR";
        table[14] = "FIXED";
        table[15] = "SIN";
        table[16] = "COS";
        table[17] = "TAN";
        table[18] = "ATAN";
        table[19] = "PI";
        table[20] = "SQRT";
        table[21] = "EXP";
        table[22] = "LN";
        table[23] = "LOG10";
        table[24] = "ABS";
        table[25] = "INT";
        table[26] = "SIGN";
        table[27] = "ROUND";
        table[28] = "LOOKUP";
        table[29] = "INDEX";
        table[30] = "REPT";
        table[31] = "MID";
        table[32] = "LEN";
        table[33] = "VALUE";
        table[34] = "TRUE";
        table[35] = "FALSE";
        table[36] = "AND";
        table[37] = "OR";
        table[38] = "NOT";
        table[39] = "MOD";
        table[40] = "DCOUNT";
        table[41] = "DSUM";
        table[42] = "DAVERAGE";
        table[43] = "DMIN";
        table[44] = "DMAX";
        table[45] = "DSTDEV";
        table[46] = "VAR";
        table[47] = "DVAR";
        table[48] = "TEXT";
        table[56] = "HLOOKUP";
        table[57] = "VLOOKUP";
        table[60] = "MATCH";
        table[61] = "DATE";
        table[63] = "FIND";
        table[65] = "ISBLANK";
        table[74] = "UPPER";
        table[75] = "LOWER";
        table[76] = "PROPER";
        table[77] = "LEFT";
        table[78] = "RIGHT";
        table[82] = "TRIM";
        table[85] = "SUBSTITUTE";
        table[92] = "SLN";
        table[97] = "CHOOSE";
        table[117] = "RANK";
        table[148] = "ROMAN";
        table[157] = "ISNUMBER";
        table[158] = "ISTEXT";
        table[171] = "NOW";
        table[173] = "TODAY";
        table[183] = "PRODUCT";
        table[190] = "ISLOGICAL";
        table[194] = "CELL";
        table[201] = "CONCATENATE";
        table[213] = "SUMIF";
        table[217] = "ADDRESS";
        table[219] = "MAXA";
        table[220] = "MINA";
        table[221] = "COUNTA";
        table[227] = "ROUNDUP";
        table[228] = "ROUNDDOWN";
        table[229] = "RAND";
        table[235] = "AVERAGEA";
        table[247] = "HYPERLINK";
        table[269] = "AVERAGEIF";
        table[275] = "RANDBETWEEN";
        table[318] = "XLOOKUP";
        table[344] = "IFERROR";

        return table;
    }

    private static string FunctionName(ushort index)
    {
        if (index < FunctionIndexTable.Length)
            return FunctionIndexTable[index];
        return $"FUNC_{index}";
    }
}

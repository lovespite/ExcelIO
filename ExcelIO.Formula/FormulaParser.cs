namespace ExcelIO.Formula;

public sealed class FormulaParser
{
    public static Expr Parse(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return new ErrorExpr("#VALUE!");

        var input = formula.StartsWith('=') ? formula.Substring(1) : formula;
        var tokenizer = new Tokenizer(input);
        try
        {
            var expr = ParseExpression(tokenizer);
            if (tokenizer.Current.Type != TokenType.Eof)
                return new ErrorExpr("#VALUE!");
            return expr;
        }
        catch (ParseException)
        {
            return new ErrorExpr("#VALUE!");
        }
    }

    private static Expr ParseExpression(Tokenizer t, int minPrec = 0)
    {
        var left = ParsePrimary(t);

        while (true)
        {
            var op = t.Current;
            int prec = Precedence(op);
            if (prec <= minPrec) break;

            if (op.Type == TokenType.Comma) break;
            if (op.Type == TokenType.RParen) break;
            if (op.Type == TokenType.Eof) break;

            t.Advance();

            if (IsBinaryOp(op))
            {
                var right = ParseExpression(t, prec);
                left = MakeBinary(op, left, right);
            }
            else if (IsUnaryOp(op))
            {
                // Postfix unary: %
                left = new UnaryExpr(MapUnaryOp(op), left);
            }
            else
            {
                throw new ParseException($"Unexpected operator: {op.Type}");
            }
        }

        return left;
    }

    private static Expr ParsePrimary(Tokenizer t)
    {
        var tok = t.Current;
        t.Advance();

        switch (tok.Type)
        {
            case TokenType.Number:
                return new NumberExpr(double.Parse(tok.Text));

            case TokenType.String:
                return new TextExpr(tok.Text);

            case TokenType.Identifier:
                // Could be a function call or a named constant (TRUE/FALSE)
                if (string.Equals(tok.Text, "TRUE", StringComparison.OrdinalIgnoreCase))
                    return new BoolExpr(true);
                if (string.Equals(tok.Text, "FALSE", StringComparison.OrdinalIgnoreCase))
                    return new BoolExpr(false);

                if (t.Current.Type == TokenType.LParen)
                {
                    t.Advance(); // skip '('
                    return ParseFunctionCall(tok.Text, t);
                }

                // Standalone cell or range reference
                return ParseCellOrRange(tok.Text, t);

            case TokenType.LParen:
                var expr = ParseExpression(t);
                if (t.Current.Type != TokenType.RParen)
                    return new ErrorExpr("#VALUE!");
                t.Advance(); // skip ')'
                return expr;

            case TokenType.Minus:
                return new UnaryExpr(UnaryOp.Neg, ParseExpression(t, 50)); // high prec for unary minus

            case TokenType.Plus:
                return new UnaryExpr(UnaryOp.Plus, ParseExpression(t, 50));

            case TokenType.Percent:
                return new UnaryExpr(UnaryOp.Percent, null!); // placeholder, handled as postfix

            case TokenType.Dollar:
            {
                // Cell or range reference starting with $
                var colPart = "";
                var rowPart = "";
                bool colAbs = true, rowAbs = false;
                if (t.Current.Type == TokenType.Identifier)
                {
                    colPart = t.Current.Text;
                    t.Advance();
                }
                if (t.Current.Type == TokenType.Dollar)
                {
                    rowAbs = true;
                    t.Advance();
                }
                if (t.Current.Type == TokenType.Number)
                {
                    rowPart = t.Current.Text;
                    t.Advance();
                }
                return ParseRangeTail(null, ParseColumn(colPart), int.Parse(rowPart) - 1, colAbs, rowAbs, t);
            }

            default:
                return new ErrorExpr("#VALUE!");
        }
    }

    private static Expr ParseFunctionCall(string name, Tokenizer t)
    {
        var args = new List<Expr>();
        if (t.Current.Type != TokenType.RParen)
        {
            args.Add(ParseExpression(t));
            while (t.Current.Type == TokenType.Comma)
            {
                t.Advance();
                args.Add(ParseExpression(t));
            }
        }
        if (t.Current.Type != TokenType.RParen)
            return new ErrorExpr("#VALUE!");
        t.Advance(); // skip ')'
        return new FunctionExpr(name.ToUpperInvariant(), args);
    }

    private static Expr ParseCellOrRange(string firstPart, Tokenizer t)
    {
        // First part could be: column letters like "A", "AB", or a sheet name like "Sheet1"
        bool colAbs = false, rowAbs = false;
        int col = ParseColumn(firstPart);
        int row = 0;

        // sheet!ref
        string? sheet = null;
        if (t.Current.Type == TokenType.Bang)
        {
            // firstPart was a sheet name
            sheet = firstPart;
            t.Advance(); // skip !
            return ParseCellOrRange("", t);
        }

        if (t.Current.Type == TokenType.Dollar)
        {
            // after column: $row
            rowAbs = true;
            t.Advance();
        }

        if (t.Current.Type == TokenType.Number)
        {
            row = int.Parse(t.Current.Text) - 1;
            t.Advance();
        }
        else
        {
            // Just a column reference like "A" (no row number)
            return new RangeRefExpr(sheet,
                int.MinValue, col, int.MaxValue, col,
                false, false, false, false);
        }

        return ParseRangeTail(sheet, col, row, colAbs, rowAbs, t);
    }

    private static Expr ParseRangeTail(string? sheet, int startCol, int startRow,
                                       bool colAbsStart, bool rowAbsStart, Tokenizer t)
    {
        if (t.Current.Type != TokenType.Colon)
            return new CellRefExpr(sheet, startRow, startCol, colAbsStart, rowAbsStart);

        t.Advance(); // skip ':'

        bool colAbsEnd = false, rowAbsEnd = false;
        int endCol, endRow;

        if (t.Current.Type == TokenType.Dollar)
        {
            colAbsEnd = true;
            t.Advance();
        }

        string? colPartEnd = null;
        if (t.Current.Type == TokenType.Identifier)
        {
            colPartEnd = t.Current.Text;
            t.Advance();
        }

        if (t.Current.Type == TokenType.Dollar)
        {
            rowAbsEnd = true;
            t.Advance();
        }

        int? rowPartEnd = null;
        if (t.Current.Type == TokenType.Number)
        {
            rowPartEnd = int.Parse(t.Current.Text);
            t.Advance();
        }

        endCol = colPartEnd != null ? ParseColumn(colPartEnd) : startCol;
        endRow = rowPartEnd.HasValue ? rowPartEnd.Value - 1 : startRow;

        return new RangeRefExpr(sheet,
            startRow, startCol, endRow, endCol,
            colAbsStart, rowAbsStart, colAbsEnd, rowAbsEnd);
    }

    // ── Helpers ──

    private static int ParseColumn(string s)
    {
        int col = 0;
        foreach (char c in s.ToUpperInvariant())
            col = col * 26 + (c - 'A' + 1);
        return col - 1; // zero-based
    }

    private static int Precedence(Token tok) => tok.Type switch
    {
        TokenType.Ampersand => 10,
        TokenType.Equal or TokenType.NotEqual or TokenType.Less or TokenType.Greater
            or TokenType.LessEqual or TokenType.GreaterEqual => 20,
        TokenType.Plus or TokenType.Minus => 30,
        TokenType.Star or TokenType.Slash => 40,
        TokenType.Caret => 50,
        TokenType.Percent => 60,
        _ => 0,
    };

    private static bool IsBinaryOp(Token tok) => tok.Type is
        TokenType.Plus or TokenType.Minus or TokenType.Star or TokenType.Slash
        or TokenType.Caret or TokenType.Ampersand
        or TokenType.Equal or TokenType.NotEqual or TokenType.Less
        or TokenType.Greater or TokenType.LessEqual or TokenType.GreaterEqual;

    private static bool IsUnaryOp(Token tok) => tok.Type is TokenType.Percent;

    private static Expr MakeBinary(Token op, Expr left, Expr right)
    {
        var binOp = MapBinaryOp(op);
        return new BinaryExpr(binOp, left, right);
    }

    private static BinaryOp MapBinaryOp(Token op) => op.Type switch
    {
        TokenType.Plus => BinaryOp.Add,
        TokenType.Minus => BinaryOp.Sub,
        TokenType.Star => BinaryOp.Mul,
        TokenType.Slash => BinaryOp.Div,
        TokenType.Caret => BinaryOp.Pow,
        TokenType.Ampersand => BinaryOp.Concat,
        TokenType.Equal => BinaryOp.Eq,
        TokenType.NotEqual => BinaryOp.Ne,
        TokenType.Less => BinaryOp.Lt,
        TokenType.Greater => BinaryOp.Gt,
        TokenType.LessEqual => BinaryOp.Le,
        TokenType.GreaterEqual => BinaryOp.Ge,
        _ => throw new ParseException($"Unknown binary operator: {op.Type}"),
    };

    private static UnaryOp MapUnaryOp(Token op) => op.Type switch
    {
        TokenType.Minus => UnaryOp.Neg,
        TokenType.Plus => UnaryOp.Plus,
        TokenType.Percent => UnaryOp.Percent,
        _ => throw new ParseException($"Unknown unary operator: {op.Type}"),
    };

    // ── Tokenizer ──

    private enum TokenType
    {
        Number, String, Identifier,
        LParen, RParen, Comma, Colon, Bang, Dollar, Percent,
        Plus, Minus, Star, Slash, Caret, Ampersand,
        Equal, NotEqual, Less, Greater, LessEqual, GreaterEqual,
        Error, Eof,
    }

    private readonly struct Token
    {
        public TokenType Type { get; }
        public string Text { get; }

        public Token(TokenType type, string text = "")
        {
            Type = type;
            Text = text;
        }
    }

    private sealed class Tokenizer
    {
        private readonly string _input;
        private int _pos;
        private Token _current;

        public Token Current => _current;

        public Tokenizer(string input)
        {
            _input = input;
            _pos = 0;
            _current = Next();
        }

        public void Advance()
        {
            _current = Next();
        }

        private Token Next()
        {
            SkipWhitespace();
            if (_pos >= _input.Length)
                return new Token(TokenType.Eof);

            char ch = _input[_pos];

            // Numbers
            if (char.IsDigit(ch) || (ch == '.' && _pos + 1 < _input.Length && char.IsDigit(_input[_pos + 1])))
                return ReadNumber();

            // Strings
            if (ch == '"')
                return ReadString();

            // Identifiers (letters only — start with letter)
            if (char.IsLetter(ch))
                return ReadIdentifier();

            // Operators
            switch (ch)
            {
                case '(': _pos++; return new Token(TokenType.LParen, "(");
                case ')': _pos++; return new Token(TokenType.RParen, ")");
                case ',': _pos++; return new Token(TokenType.Comma, ",");
                case ':': _pos++; return new Token(TokenType.Colon, ":");
                case '!': _pos++; return new Token(TokenType.Bang, "!");
                case '$': _pos++; return new Token(TokenType.Dollar, "$");
                case '%': _pos++; return new Token(TokenType.Percent, "%");
                case '+': _pos++; return new Token(TokenType.Plus, "+");
                case '-': _pos++; return new Token(TokenType.Minus, "-");
                case '*': _pos++; return new Token(TokenType.Star, "*");
                case '/': _pos++; return new Token(TokenType.Slash, "/");
                case '^': _pos++; return new Token(TokenType.Caret, "^");
                case '&': _pos++; return new Token(TokenType.Ampersand, "&");
                case '=':
                    _pos++;
                    return new Token(TokenType.Equal, "=");
                case '<':
                    _pos++;
                    if (_pos < _input.Length && _input[_pos] == '>') { _pos++; return new Token(TokenType.NotEqual, "<>"); }
                    if (_pos < _input.Length && _input[_pos] == '=') { _pos++; return new Token(TokenType.LessEqual, "<="); }
                    return new Token(TokenType.Less, "<");
                case '>':
                    _pos++;
                    if (_pos < _input.Length && _input[_pos] == '=') { _pos++; return new Token(TokenType.GreaterEqual, ">="); }
                    return new Token(TokenType.Greater, ">");
                default:
                    return new Token(TokenType.Error, ch.ToString());
            }
        }

        private void SkipWhitespace()
        {
            while (_pos < _input.Length && char.IsWhiteSpace(_input[_pos]))
                _pos++;
        }

        private Token ReadNumber()
        {
            int start = _pos;
            while (_pos < _input.Length && (char.IsDigit(_input[_pos]) || _input[_pos] == '.'))
                _pos++;
            return new Token(TokenType.Number, _input.Substring(start, _pos - start));
        }

        private Token ReadString()
        {
            _pos++; // skip opening "
            int start = _pos;
            while (_pos < _input.Length && _input[_pos] != '"')
                _pos++;
            var text = _input.Substring(start, _pos - start);
            if (_pos < _input.Length) _pos++; // skip closing "
            return new Token(TokenType.String, text);
        }

        private Token ReadIdentifier()
        {
            int start = _pos;
            while (_pos < _input.Length && char.IsLetter(_input[_pos]))
                _pos++;
            return new Token(TokenType.Identifier, _input.Substring(start, _pos - start));
        }
    }

    private sealed class ParseException : Exception
    {
        public ParseException(string message) : base(message) { }
    }
}

namespace PowerTables.Core.Formulas;

public abstract record Node;
public sealed record NumNode(double V) : Node;
public sealed record StrNode(string V) : Node;
public sealed record BoolNode(bool V) : Node;
public sealed record ErrNode(CellError V) : Node;
public sealed record MissingNode : Node;
public sealed record RefNode(RefToken R) : Node;
public sealed record NameNode(string Name) : Node;
public sealed record UnaryNode(char Op, Node X) : Node;
public sealed record PercentNode(Node X) : Node;
public sealed record BinNode(string Op, Node L, Node R) : Node;
public sealed record CallNode(string Name, Node[] Args) : Node;
public sealed record ArrayNode(Node[][] Rows) : Node;

/// <summary>Parser a discesa ricorsiva con le precedenze di Excel (il meno unario lega più della potenza).</summary>
public sealed class Parser
{
    readonly List<Token> _t;
    int _p;
    Parser(List<Token> t) { _t = t; }

    /// <summary>Analizza "formula" (senza '=' iniziale). Lancia FormatException se non valida.</summary>
    public static Node Parse(string formula)
    {
        var p = new Parser(Lexer.Tokenize(formula));
        var n = p.Comparison();
        if (p._t[p._p].Kind != T.End) throw new FormatException("token inatteso: " + p._t[p._p].Text);
        return n;
    }

    Token Cur => _t[_p];
    bool IsOp(string op) => Cur.Kind == T.Op && Cur.Text == op;

    Node Comparison()
    {
        var l = Concat();
        while (Cur.Kind == T.Op && Cur.Text is "=" or "<>" or "<" or ">" or "<=" or ">=")
        {
            if (Cur.Kind != T.Op) break;
            string op = Cur.Text; _p++;
            l = new BinNode(op, l, Concat());
        }
        return l;
    }

    Node Concat()
    {
        var l = Add();
        while (IsOp("&")) { _p++; l = new BinNode("&", l, Add()); }
        return l;
    }

    Node Add()
    {
        var l = Mul();
        while (IsOp("+") || IsOp("-")) { string op = Cur.Text; _p++; l = new BinNode(op, l, Mul()); }
        return l;
    }

    Node Mul()
    {
        var l = Pow();
        while (IsOp("*") || IsOp("/")) { string op = Cur.Text; _p++; l = new BinNode(op, l, Pow()); }
        return l;
    }

    Node Pow()
    {
        var l = Unary();
        while (IsOp("^")) { _p++; l = new BinNode("^", l, Unary()); }
        return l;
    }

    Node Unary()
    {
        if (IsOp("-") || IsOp("+")) { char op = Cur.Text[0]; _p++; return new UnaryNode(op, Unary()); }
        var n = Primary();
        while (IsOp("%")) { _p++; n = new PercentNode(n); }
        return n;
    }

    Node Primary()
    {
        var t = Cur;
        switch (t.Kind)
        {
            case T.Num: _p++; return new NumNode(t.Num);
            case T.Str: _p++; return new StrNode(t.Text);
            case T.Bool: _p++; return new BoolNode(t.Num != 0);
            case T.Err: _p++; return new ErrNode(t.Err);
            case T.Ref:
                _p++;
                if (Cur.Kind == T.Colon && !t.Ref!.IsRange) throw new FormatException("operatore ':' tra riferimenti non supportato");
                return new RefNode(t.Ref!);
            case T.Name: _p++; return new NameNode(t.Text);
            case T.LP:
                {
                    _p++; var n = Comparison();
                    if (Cur.Kind != T.RP) throw new FormatException("atteso ')'");
                    _p++; return n;
                }
            case T.LB: return ArrayLiteral();
            case T.Func:
                {
                    _p += 2; // nome + '('
                    var args = new List<Node>();
                    if (Cur.Kind == T.RP) { _p++; return new CallNode(Normalize(t.Text), []); }
                    while (true)
                    {
                        if (Cur.Kind is T.Comma or T.RP) args.Add(new MissingNode()); else args.Add(Comparison());
                        if (Cur.Kind == T.Comma) { _p++; continue; }
                        if (Cur.Kind == T.RP) { _p++; break; }
                        throw new FormatException("atteso ',' o ')'");
                    }
                    return new CallNode(Normalize(t.Text), [.. args]);
                }
            default: throw new FormatException("espressione attesa vicino a '" + t.Text + "'");
        }
    }

    Node ArrayLiteral()
    {
        _p++; // {
        var rows = new List<Node[]>(); var cur = new List<Node>();
        while (true)
        {
            bool neg = false;
            if (IsOp("-")) { neg = true; _p++; }
            var t = Cur; Node v;
            switch (t.Kind)
            {
                case T.Num: v = new NumNode(neg ? -t.Num : t.Num); break;
                case T.Str: v = new StrNode(t.Text); break;
                case T.Bool: v = new BoolNode(t.Num != 0); break;
                case T.Err: v = new ErrNode(t.Err); break;
                default: throw new FormatException("costante di matrice attesa");
            }
            _p++; cur.Add(v);
            if (Cur.Kind == T.Comma) { _p++; continue; }
            if (Cur.Kind == T.Semi) { _p++; rows.Add([.. cur]); cur.Clear(); continue; }
            if (Cur.Kind == T.RB) { _p++; rows.Add([.. cur]); break; }
            throw new FormatException("matrice non valida");
        }
        int w = rows[0].Length;
        if (rows.Any(r => r.Length != w)) throw new FormatException("righe di matrice di lunghezza diversa");
        return new ArrayNode([.. rows]);
    }

    /// <summary>Nome funzione in forma canonica (maiuscolo, senza prefissi _xlfn.).</summary>
    public static string Normalize(string name)
    {
        string n = name.ToUpperInvariant();
        if (n.StartsWith("_XLFN.")) n = n[6..];
        if (n.StartsWith("_XLWS.")) n = n[6..];
        return n;
    }
}

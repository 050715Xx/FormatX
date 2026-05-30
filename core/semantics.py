"""Shared formula semantic constants used across parse/convert/style."""

from __future__ import annotations

FUNCTION_NAMES = frozenset(
    {
        "sin", "cos", "tan", "sec", "csc", "cot",
        "sinh", "cosh", "tanh",
        "log", "ln", "exp",
        "max", "min", "det", "lim",
        "arcsin", "arccos", "arctan",
        "gcd", "inf", "sup", "liminf", "limsup",
    }
)

BIG_OPERATOR_NAMES = frozenset({
    "sum", "int", "iint", "iiint", "oint", "prod", "lim",
    "coprod", "bigcap", "bigcup", "bigsqcup", "bigvee", "bigwedge",
    "bigodot", "bigoplus", "bigotimes", "oiint", "oiiint",
})

GREEK_COMMAND_NAMES = frozenset(
    {
        "alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta",
        "theta", "iota", "kappa", "lambda", "mu", "nu", "xi",
        "pi", "rho", "sigma", "tau", "upsilon", "phi", "chi", "psi", "omega",
        "varepsilon", "vartheta", "varpi", "varrho", "varsigma", "varphi",
    }
)

BOLD_ITALIC_MARKER_COMMANDS = frozenset({"vec", "overrightarrow", "boldsymbol"})
BOLD_ROMAN_MARKER_COMMANDS = frozenset({"mathbf"})
ROMAN_MARKER_COMMANDS = frozenset({"mathrm", "operatorname"})
ITALIC_MARKER_COMMANDS = frozenset({"mathit"})
UPRIGHT_FAMILY_MARKER_COMMANDS = frozenset(
    {
        "mathbb",
        "mathcal",
        "mathfrak",
        "mathscr",
        "mathsf",
        "mathtt",
    }
)

RELATION_AND_OPERATOR_SYMBOLS = frozenset({
    "pm", "mp", "times", "div", "cdot", "colon", "circ", "bullet",
    "ast", "star", "dagger", "ddagger", "amalg", "wr", "diamond",
    "bigcirc", "setminus", "cap", "cup", "sqcap", "sqcup", "wedge",
    "vee", "oplus", "ominus", "otimes", "oslash", "odot",
    "leq", "geq", "ll", "gg", "equiv", "sim", "simeq", "approx",
    "neq", "doteq", "propto", "parallel", "perp", "mid", "bowtie",
    "subset", "supset", "subseteq", "supseteq", "in", "ni", "notin",
    "sqsubseteq", "sqsupseteq", "models", "prec", "succ", "preceq",
    "succeq", "asymp", "smile", "frown",
})

ARROW_SYMBOLS = frozenset({
    "leftarrow", "rightarrow", "uparrow", "downarrow",
    "Leftarrow", "Rightarrow", "Uparrow", "Downarrow",
    "leftrightarrow", "Leftrightarrow", "mapsto", "longmapsto",
    "longleftarrow", "longrightarrow", "longleftrightarrow",
    "Longleftarrow", "Longrightarrow", "Longleftrightarrow",
    "hookleftarrow", "hookrightarrow", "leftharpoonup",
    "leftharpoondown", "rightharpoonup", "rightharpoondown",
    "rightleftharpoons", "nearrow", "searrow", "swarrow", "nwarrow",
    "iff", "to", "gets", "implies",
})

DOT_SYMBOLS = frozenset({"cdots", "ddots", "vdots", "ldots"})

EXTRA_ACCENT_COMMANDS = frozenset({
    "acute", "breve", "check", "grave", "mathring",
})

PHANTOM_AND_STYLE_COMMANDS = frozenset({
    "displaystyle", "textstyle", "hphantom", "phantom", "vphantom",
})

STACKED_AND_EXTENSIBLE_COMMANDS = frozenset({
    "overbrace", "overbracket", "overleftarrow",
    "overleftrightarrow", "overset", "stackrel",
    "underbrace", "underbracket", "underleftarrow",
    "underleftrightarrow", "underrightarrow", "underset",
    "xleftarrow", "xrightarrow",
})

MISC_MATH_SYMBOLS = frozenset({
    "infty", "partial", "nabla", "ell", "Re", "Im", "emptyset",
    "varnothing", "aleph", "hbar", "imath", "jmath", "wp",
    "angle", "triangle", "triangledown", "triangleleft",
    "triangleright", "square", "Box", "Diamond", "mho",
    "forall", "exists", "neg", "flat", "sharp", "natural",
    "surd", "prime", "backslash", "dag", "ddag",
})

FORMATX_LEGACY_COMPAT_COMMANDS = frozenset({
    "and", "or", "not", "ne", "le", "ge",
})

STRUCTURE_DELIMITERS = frozenset({"left", "right"})

STRUCTURAL_COMMANDS = frozenset({
    "sqrt", "frac", "left", "right", "begin", "end", "text",
    "over", "choose", "atop", "above", "brace", "brack",
    "hat", "bar", "tilde", "dot", "ddot", "vec",
    "underline", "overline", "widehat", "widetilde",
    "limits", "nolimits",
})

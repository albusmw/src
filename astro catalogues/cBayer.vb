Option Explicit On
Option Strict On

Public Class cBayer

  Public Shared Function ToGreek(ByVal Text As String) As String

    Text = Text.Replace("alf", "α")
    Text = Text.Replace("bet", "β")
    Text = Text.Replace("gam", "γ")
    Text = Text.Replace("kap", "κ")
    Text = Text.Replace("eps", "ε")
    Text = Text.Replace("the", "θ")
    Text = Text.Replace("chi", "χ")
    Text = Text.Replace("sig", "σ")
    Text = Text.Replace("iot", "ι")
    Text = Text.Replace("zet", "ζ")
    Text = Text.Replace("rho", "ρ")
    Text = Text.Replace("lam", "λ")
    Text = Text.Replace("omi", "ο")
    Text = Text.Replace("eta", "η")
    Text = Text.Replace("phi", "φ")
    Text = Text.Replace("ome", "ω")
    Text = Text.Replace("tau", "τ")
    Text = Text.Replace("ups", "ϒ")
    Text = Text.Replace("psi", "ψ")
    Text = Text.Replace("pi", "π")
    Text = Text.Replace("del", "δ")

    Text = Text.Replace("01", "₁")
    Text = Text.Replace("02", "₂")
    Text = Text.Replace("03", "₃")
    Text = Text.Replace("04", "₄")
    Text = Text.Replace("05", "₅")
    Text = Text.Replace("06", "₆")
    Text = Text.Replace("07", "₇")
    Text = Text.Replace("08", "₈")
    Text = Text.Replace("09", "₉")

    Return Text

  End Function

End Class

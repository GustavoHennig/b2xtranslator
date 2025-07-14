using System;
using System.Collections.Generic;
using System.Globalization;

namespace b2xtranslator.Tools
{
    /// <summary>
    /// Utility class for converting symbol font characters to Unicode equivalents
    /// </summary>
    public static class SymbolMapping
    {
        /// <summary>
        /// Symbol font character mappings (Microsoft's Symbol font)
        /// </summary>
        private static readonly Dictionary<byte, string> SymbolFontMap = new Dictionary<byte, string>
        {
            // Greek letters (lowercase)
            { 0x61, "α" },  // alpha
            { 0x62, "β" },  // beta
            { 0x63, "χ" },  // chi
            { 0x64, "δ" },  // delta
            { 0x65, "ε" },  // epsilon
            { 0x66, "φ" },  // phi
            { 0x67, "γ" },  // gamma
            { 0x68, "η" },  // eta
            { 0x69, "ι" },  // iota
            { 0x6A, "ϕ" },  // phi (variant)
            { 0x6B, "κ" },  // kappa
            { 0x6C, "λ" },  // lambda
            { 0x6D, "μ" },  // mu
            { 0x6E, "ν" },  // nu
            { 0x6F, "ο" },  // omicron
            { 0x70, "π" },  // pi
            { 0x71, "θ" },  // theta
            { 0x72, "ρ" },  // rho
            { 0x73, "σ" },  // sigma
            { 0x74, "τ" },  // tau
            { 0x75, "υ" },  // upsilon
            { 0x76, "ϖ" },  // pi (variant)
            { 0x77, "ω" },  // omega
            { 0x78, "ξ" },  // xi
            { 0x79, "ψ" },  // psi
            { 0x7A, "ζ" },  // zeta
            { 0x7B, "{" },  // left curly bracket
            { 0x7C, "|" },  // vertical bar
            { 0x7D, "}" },  // right curly bracket
            { 0x7E, "~" },  // tilde
            { 0x7F, "" },  // del character (non-printable)
            
            // Greek letters (uppercase)
            { 0x41, "Α" },  // Alpha
            { 0x42, "Β" },  // Beta
            { 0x43, "Χ" },  // Chi
            { 0x44, "Δ" },  // Delta
            { 0x45, "Ε" },  // Epsilon
            { 0x46, "Φ" },  // Phi
            { 0x47, "Γ" },  // Gamma
            { 0x48, "Η" },  // Eta
            { 0x49, "Ι" },  // Iota
            { 0x4A, "ϑ" },  // theta (variant)
            { 0x4B, "Κ" },  // Kappa
            { 0x4C, "Λ" },  // Lambda
            { 0x4D, "Μ" },  // Mu
            { 0x4E, "Ν" },  // Nu
            { 0x4F, "Ο" },  // Omicron
            { 0x50, "Π" },  // Pi
            { 0x51, "Θ" },  // Theta
            { 0x52, "Ρ" },  // Rho
            { 0x53, "Σ" },  // Sigma
            { 0x54, "Τ" },  // Tau
            { 0x55, "Υ" },  // Upsilon
            { 0x56, "ς" },  // sigma (final)
            { 0x57, "Ω" },  // Omega
            { 0x58, "Ξ" },  // Xi
            { 0x59, "Ψ" },  // Psi
            { 0x5A, "Ζ" },  // Zeta
            
            // Mathematical symbols
            { 0x2B, "±" },  // plus-minus
            { 0x2D, "−" },  // minus
            { 0x2A, "∗" },  // asterisk operator
            { 0x2F, "/" },  // solidus
            { 0x3D, "=" },  // equals
            { 0x3C, "<" },  // less-than
            { 0x3E, ">" },  // greater-than
            { 0xB1, "±" },  // plus-minus
            { 0xB7, "·" },  // middle dot
            { 0xD7, "×" },  // multiplication
            { 0xF7, "÷" },  // division
            { 0xB0, "°" },  // degree
            { 0xB5, "μ" },  // micro
            { 0xA3, "≤" },  // less-than or equal
            { 0xB3, "≥" },  // greater-than or equal
            { 0xB9, "¹" },  // superscript one
            { 0xB2, "²" },  // superscript two
            { 0xBD, "½" },  // one half
            { 0xBC, "¼" },  // one quarter
            { 0xBE, "¾" },  // three quarters
            
            // Arrows
            { 0xAB, "←" },  // left arrow
            { 0xBB, "→" },  // right arrow
            { 0xAC, "↑" },  // up arrow
            { 0xDB, "↔" },  // left-right arrow
            { 0xDD, "↕" },  // up-down arrow
            
            // Other symbols
            { 0xA5, "∞" },  // infinity
            { 0xB6, "∂" },  // partial differential
            { 0xD1, "∑" },  // summation
            { 0xD5, "∏" },  // product
            { 0xD6, "√" },  // square root
            { 0xD8, "∝" },  // proportional to
            { 0xDC, "∠" },  // angle
            { 0xE0, "◊" },  // diamond
            { 0xE5, "∅" },  // empty set
            { 0xE6, "∈" },  // element of
            { 0xE7, "∉" },  // not element of
            { 0xE8, "∋" },  // contains
            { 0xE9, "∌" },  // not contains
            { 0xEA, "∩" },  // intersection
            { 0xEB, "∪" },  // union
            { 0xEC, "⊂" },  // subset
            { 0xED, "⊃" },  // superset
            { 0xEE, "⊆" },  // subset or equal
            { 0xEF, "⊇" },  // superset or equal
            { 0xF0, "⊥" },  // perpendicular
            { 0xF1, "∴" },  // therefore
            { 0xF2, "∵" },  // because
        };

        /// <summary>
        /// Wingdings font character mappings
        /// </summary>
        private static readonly Dictionary<byte, string> WingdingsFontMap = new Dictionary<byte, string>
        {
            // Common symbols
            { 0x4A, "☺" },  // smiling face
            { 0x4B, "☻" },  // black smiling face
            { 0x4C, "♥" },  // heart
            { 0x4D, "♦" },  // diamond
            { 0x4E, "♣" },  // club
            { 0x4F, "♠" },  // spade
            { 0x6C, "✓" },  // check mark
            { 0x6D, "✗" },  // ballot x
            { 0x6E, "✪" },  // star
            { 0x6F, "✫" },  // star
            { 0x70, "✬" },  // star
            { 0x71, "✭" },  // star
            { 0x72, "✮" },  // star
            { 0x73, "✯" },  // star
            { 0x74, "✰" },  // star
            { 0x75, "✱" },  // star
            { 0x76, "✲" },  // star
            { 0x77, "✳" },  // star
            { 0x78, "✴" },  // star
            { 0x79, "✵" },  // star
            { 0x7A, "✶" },  // star
            { 0x81, "✁" },  // scissors
            { 0x82, "✂" },  // scissors
            { 0x83, "✃" },  // scissors
            { 0x84, "✄" },  // scissors
            { 0x85, "☎" },  // telephone
            { 0x86, "✆" },  // telephone
            { 0x87, "✇" },  // tape drive
            { 0x88, "✈" },  // airplane
            { 0x89, "✉" },  // envelope
            { 0x8A, "✊" },  // fist
            { 0x8B, "✋" },  // hand
            { 0x8C, "✌" },  // victory hand
            { 0x8D, "✍" },  // writing hand
            { 0x8E, "✎" },  // pencil
            { 0x8F, "✏" },  // pencil
            { 0x90, "✐" },  // pencil
            { 0x91, "✑" },  // pencil
            { 0x92, "✒" },  // pen
            { 0x93, "✓" },  // check mark
            { 0x94, "✔" },  // check mark
            { 0x95, "✕" },  // multiplication x
            { 0x96, "✖" },  // multiplication x
            { 0x97, "✗" },  // ballot x
            { 0x98, "✘" },  // ballot x
            { 0x99, "✙" },  // cross
            { 0x9A, "✚" },  // cross
            { 0x9B, "✛" },  // cross
            { 0x9C, "✜" },  // cross
            { 0x9D, "✝" },  // cross
            { 0x9E, "✞" },  // cross
            { 0x9F, "✟" },  // cross
            { 0xA0, "✠" },  // cross
            { 0xA1, "✡" },  // star of David
            { 0xA2, "✢" },  // star
            { 0xA3, "✣" },  // star
            { 0xA4, "✤" },  // star
            { 0xA5, "✥" },  // star
            { 0xA6, "✦" },  // star
            { 0xA7, "✧" },  // star
            { 0xA8, "✨" },  // sparkles
            { 0xA9, "✩" },  // star
            { 0xAA, "✪" },  // star
            { 0xAB, "✫" },  // star
            { 0xAC, "✬" },  // star
            { 0xAD, "✭" },  // star
            { 0xAE, "✮" },  // star
            { 0xAF, "✯" },  // star
        };

        /// <summary>
        /// Converts a symbol character to its Unicode equivalent
        /// </summary>
        /// <param name="charCode">The character code from the symbol font</param>
        /// <param name="fontName">The name of the symbol font</param>
        /// <returns>The Unicode string representation of the symbol</returns>
        public static string ConvertSymbolCharacter(byte charCode, string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return ConvertToUnicode(charCode, fontName);

            string normalizedFontName = fontName.ToLower();
            
            switch (normalizedFontName)
            {
                case "symbol":
                    if (SymbolFontMap.TryGetValue(charCode, out string symbolResult))
                        return symbolResult;
                    break;
                    
                case "wingdings":
                    if (WingdingsFontMap.TryGetValue(charCode, out string wingdingResult))
                        return wingdingResult;
                    break;
                    
                case "webdings":
                    // Add Webdings support in the future
                    break;
                    
                case "mt extra":
                    // Add MT Extra support in the future
                    break;
            }

            return ConvertToUnicode(charCode, fontName);
        }

        /// <summary>
        /// Converts a hex string to its Unicode equivalent
        /// </summary>
        /// <param name="hexValue">The hex string (e.g., "03b1" for alpha)</param>
        /// <param name="fontName">The name of the symbol font</param>
        /// <returns>The Unicode string representation of the symbol</returns>
        public static string ConvertSymbolHex(string hexValue, string fontName)
        {
            if (string.IsNullOrEmpty(hexValue))
                return "?";

            try
            {
                if (ushort.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort code))
                {
                    return ConvertSymbolCharacter((byte)code, fontName);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error converting hex symbol: {hexValue}, Font: {fontName}, Error: {ex.Message}");
            }

            return "?";
        }

        /// <summary>
        /// Fallback Unicode conversion for unknown symbols
        /// </summary>
        /// <param name="charCode">The character code</param>
        /// <param name="fontName">The font name</param>
        /// <returns>The converted character or a placeholder</returns>
        private static string ConvertToUnicode(byte charCode, string fontName)
        {
            // For characters in the ASCII range, try to use them directly
            if (charCode > 32 && charCode < 127)
            {
                return ((char)charCode).ToString();
            }

            // For high-range characters, return placeholder
            if (charCode > 127)
            {
                return "?";
            }

            // Log unknown symbol for debugging
            System.Diagnostics.Debug.WriteLine($"Unknown symbol: Code={charCode:X2}, Font={fontName}");
            
            // Return placeholder for unknown symbols
            return "?";
        }

        /// <summary>
        /// Checks if a font name is a known symbol font
        /// </summary>
        /// <param name="fontName">The font name to check</param>
        /// <returns>True if the font is a known symbol font</returns>
        public static bool IsSymbolFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            string normalizedFontName = fontName.ToLower();
            string[] symbolFonts = { "symbol", "wingdings", "webdings", "mt extra" };
            
            foreach (string symbolFont in symbolFonts)
            {
                if (normalizedFontName.Contains(symbolFont))
                    return true;
            }
            
            return false;
        }

        /// <summary>
        /// Provides a contextual fallback for unknown symbols
        /// </summary>
        /// <param name="charCode">The character code</param>
        /// <param name="fontName">The font name</param>
        /// <returns>A contextual fallback string</returns>
        public static string GetContextualFallback(byte charCode, string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return "?";

            string normalizedFontName = fontName.ToLower();
            
            switch (normalizedFontName)
            {
                case "symbol":
                    return GetMathematicalFallback(charCode);
                    
                case "wingdings":
                    return GetDecorativeFallback(charCode);
                    
                default:
                    return "?";
            }
        }

        /// <summary>
        /// Provides fallback for mathematical symbols
        /// </summary>
        /// <param name="charCode">The character code</param>
        /// <returns>A mathematical fallback string</returns>
        private static string GetMathematicalFallback(byte charCode)
        {
            switch (charCode)
            {
                case >= 0x61 and <= 0x7A: // Greek letters (lowercase)
                    return "[greek]";
                    
                case >= 0x41 and <= 0x5A: // Greek letters (uppercase)
                    return "[GREEK]";
                    
                case >= 0xB1 and <= 0xF7: // Math operators
                    return "[math]";
                    
                default:
                    return "?";
            }
        }

        /// <summary>
        /// Provides fallback for decorative symbols
        /// </summary>
        /// <param name="charCode">The character code</param>
        /// <returns>A decorative fallback string</returns>
        private static string GetDecorativeFallback(byte charCode)
        {
            switch (charCode)
            {
                case >= 0x4A and <= 0x4F: // Faces and card suits
                    return "[symbol]";
                    
                case >= 0x6C and <= 0x7A: // Stars and checks
                    return "[star]";
                    
                case >= 0x81 and <= 0x9F: // Tools and objects
                    return "[object]";
                    
                default:
                    return "?";
            }
        }
    }
}
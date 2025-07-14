using System;
using NUnit.Framework;
using b2xtranslator.Tools;

namespace UnitTests
{
    [TestFixture]
    public class SymbolHandlingTests
    {
        [Test]
        public void ShouldConvertSymbolFontCharacters()
        {
            // Test known Symbol font mappings
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x61, "Symbol"), Is.EqualTo("α")); // alpha
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x62, "Symbol"), Is.EqualTo("β")); // beta
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x67, "Symbol"), Is.EqualTo("γ")); // gamma
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x64, "Symbol"), Is.EqualTo("δ")); // delta
        }

        [Test]
        public void ShouldConvertSymbolFontCharactersCaseInsensitive()
        {
            // Test case insensitive font name
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x61, "SYMBOL"), Is.EqualTo("α")); // alpha
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x62, "symbol"), Is.EqualTo("β")); // beta
        }

        [Test]
        public void ShouldHandleWingdingsFont()
        {
            // Test known Wingdings font mappings
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x4A, "Wingdings"), Is.EqualTo("☺")); // smiling face
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x6C, "Wingdings"), Is.EqualTo("✓")); // check mark
        }

        [Test]
        public void ShouldConvertHexValues()
        {
            // Test hex string conversion
            Assert.That(SymbolMapping.ConvertSymbolHex("61", "Symbol"), Is.EqualTo("α")); // alpha
            Assert.That(SymbolMapping.ConvertSymbolHex("62", "Symbol"), Is.EqualTo("β")); // beta
        }

        [Test]
        public void ShouldHandleUnknownSymbols()
        {
            // Test unknown symbol characters return fallback
            string result = SymbolMapping.ConvertSymbolCharacter(0xFF, "Symbol");
            Assert.That(result, Is.EqualTo("?")); // Should return placeholder
        }

        [Test]
        public void ShouldHandleUnknownFonts()
        {
            // Test unknown font names
            string result = SymbolMapping.ConvertSymbolCharacter(0x61, "UnknownFont");
            Assert.That(result, Is.Not.Null);
        }

        [Test]
        public void ShouldIdentifySymbolFonts()
        {
            // Test symbol font identification
            Assert.That(SymbolMapping.IsSymbolFont("Symbol"), Is.True);
            Assert.That(SymbolMapping.IsSymbolFont("Wingdings"), Is.True);
            Assert.That(SymbolMapping.IsSymbolFont("Arial"), Is.False);
            Assert.That(SymbolMapping.IsSymbolFont("Times New Roman"), Is.False);
        }

        [Test]
        public void ShouldHandleNullOrEmptyInputs()
        {
            // Test null and empty inputs
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x61, null), Is.EqualTo("a")); // ASCII fallback
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x61, ""), Is.EqualTo("a")); // ASCII fallback
            Assert.That(SymbolMapping.ConvertSymbolHex("", "Symbol"), Is.EqualTo("?"));
            Assert.That(SymbolMapping.ConvertSymbolHex(null, "Symbol"), Is.EqualTo("?"));
        }

        [Test]
        public void ShouldHandleInvalidHexValues()
        {
            // Test invalid hex strings
            Assert.That(SymbolMapping.ConvertSymbolHex("GG", "Symbol"), Is.EqualTo("?"));
            Assert.That(SymbolMapping.ConvertSymbolHex("xyz", "Symbol"), Is.EqualTo("?"));
        }

        [Test]
        public void ShouldProvideContextualFallbacks()
        {
            // Test contextual fallbacks for unknown symbols
            string symbolFallback = SymbolMapping.GetContextualFallback(0x61, "Symbol");
            Assert.That(symbolFallback, Is.EqualTo("[greek]")); // Greek letter fallback
            
            string wingdingsFallback = SymbolMapping.GetContextualFallback(0x4A, "Wingdings");
            Assert.That(wingdingsFallback, Is.EqualTo("[symbol]")); // Symbol fallback
        }

        [Test]
        public void ShouldHandleMathematicalSymbols()
        {
            // Test mathematical symbols from Symbol font
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0xB1, "Symbol"), Is.EqualTo("±")); // plus-minus
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0xD7, "Symbol"), Is.EqualTo("×")); // multiplication
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0xF7, "Symbol"), Is.EqualTo("÷")); // division
        }

        [Test]
        public void ShouldHandleGreekLetters()
        {
            // Test Greek letters (both uppercase and lowercase)
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x41, "Symbol"), Is.EqualTo("Α")); // Alpha
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x42, "Symbol"), Is.EqualTo("Β")); // Beta
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x47, "Symbol"), Is.EqualTo("Γ")); // Gamma
            Assert.That(SymbolMapping.ConvertSymbolCharacter(0x44, "Symbol"), Is.EqualTo("Δ")); // Delta
        }
    }
}
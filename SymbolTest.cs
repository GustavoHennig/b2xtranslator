using System;
using b2xtranslator.Tools;

class SymbolTest
{
    static void Main()
    {
        Console.WriteLine("Testing Symbol Mapping...");
        
        // Test some basic symbol conversions
        var result1 = SymbolMapping.ConvertSymbolCharacter(0x61, "Symbol"); // Should be α
        Console.WriteLine($"Symbol font 0x61: {result1}");
        
        var result2 = SymbolMapping.ConvertSymbolCharacter(0x62, "Symbol"); // Should be β  
        Console.WriteLine($"Symbol font 0x62: {result2}");
        
        var result3 = SymbolMapping.ConvertSymbolHex("0061", "Symbol"); // Should be α
        Console.WriteLine($"Symbol font hex 0061: {result3}");
        
        var result4 = SymbolMapping.ConvertSymbolHex("0062", "Symbol"); // Should be β
        Console.WriteLine($"Symbol font hex 0062: {result4}");
        
        // Test character codes that should be in the symbol document
        var result5 = SymbolMapping.ConvertSymbolCharacter(0x21, "Symbol"); // Should be !
        Console.WriteLine($"Symbol font 0x21: {result5}");
        
        var result6 = SymbolMapping.ConvertSymbolCharacter(0x2A, "Symbol"); // Should be ∗
        Console.WriteLine($"Symbol font 0x2A: {result6}");
        
        Console.WriteLine("Test completed.");
    }
}
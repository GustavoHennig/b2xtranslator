using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class PieceTable
    {
        /// <summary>
        /// A list of PieceDescriptor standing for each piece of text.
        /// </summary>
        public List<PieceDescriptor> Pieces;

        /// <summary>
        /// A dictionary with character positions as keys and the matching FCs as values
        /// </summary>
        public Dictionary<int, int> FileCharacterPositions;

        /// <summary>
        /// A dictionary with file character positions as keys and the matching CPs as values
        /// </summary>
        public Dictionary<int, int> CharacterPositions;

        /// <summary>
        /// Creates an empty piece table for Word 95 files (single piece fallback)
        /// </summary>
        /// <param name="fib">The FIB</param>
        public PieceTable(FileInformationBlock fib, Encoding singleByteEncoding = null)
        {
            this.Pieces = new List<PieceDescriptor>();
            this.FileCharacterPositions = new Dictionary<int, int>();
            this.CharacterPositions = new Dictionary<int, int>();

            // Create a single piece descriptor covering the entire document text
            // Word 95 files typically store text as a single piece from fcMin to fcMac
            Encoding encoding1252 = GetSingleByteEncoding(singleByteEncoding);

            var singlePiece = new PieceDescriptor()
            {
                cpStart = 0,
                cpEnd = fib.ccpText,
                fc = (uint)fib.fcMin,
                encoding = encoding1252 // Word 95 typically uses Windows-1252
            };

            this.Pieces.Add(singlePiece);

            // Build the position mappings for the single piece
            int f = fib.fcMin;
            for (int c = 0; c < fib.ccpText; c++)
            {
                if (!this.FileCharacterPositions.ContainsKey(c))
                    this.FileCharacterPositions.Add(c, f);
                if (!this.CharacterPositions.ContainsKey(f))
                    this.CharacterPositions.Add(f, c);
                f++;
            }
            
            // Add the final position
            this.FileCharacterPositions.Add(fib.ccpText, fib.fcMac);
            this.CharacterPositions.Add(fib.fcMac, fib.ccpText);
        }

        /// <summary>
        /// Parses the pice table and creates a list of PieceDescriptors.
        /// </summary>
        /// <param name="fib">The FIB</param>
        /// <param name="tableStream">The 0Table or 1Table stream</param>
        public PieceTable(FileInformationBlock fib, VirtualStream tableStream, Encoding singleByteEncoding = null)
        {
            //Read the bytes of complex file information
            var bytes = new byte[fib.lcbClx];
            tableStream.Read(bytes, 0, (int)fib.lcbClx, (int)fib.fcClx);

            this.Pieces = new List<PieceDescriptor>();
            this.FileCharacterPositions = new Dictionary<int, int>();
            this.CharacterPositions = new Dictionary<int, int>();
            Encoding defaultSingleByteEncoding = GetSingleByteEncoding(singleByteEncoding);

            int pos = 0;
            bool goon = true;
            while (goon)
            {
                try
                {
                    if (pos >= bytes.Length)
                    {
                        goon = false;
                        break;
                    }
                    byte type = bytes[pos];

                    //check if the type of the entry is a piece table
                    if (type == 2)
                    {
                        int lcb = System.BitConverter.ToInt32(bytes, pos + 1);
  
                        //read the piece table
                        var piecetable = new byte[lcb];
                        Array.Copy(bytes, pos + 5, piecetable, 0, piecetable.Length);

                        //count of PCD _entries
                        int n = (lcb - 4) / 12;

                        //and n piece descriptors
                        for (int i = 0; i < n; i++)
                        {
                            //read the CP 
                            int indexCp = i * 4;
                            int cp = System.BitConverter.ToInt32(piecetable, indexCp);

                            //read the next CP
                            int indexCpNext = (i+1) * 4;
                            int cpNext = System.BitConverter.ToInt32(piecetable, indexCpNext);

                            //read the PCD
                            int indexPcd = ((n + 1) * 4) + (i * 8);
                            var pcdBytes = new byte[8];
                            Array.Copy(piecetable, indexPcd, pcdBytes, 0, 8);
                            var pcd = new PieceDescriptor(pcdBytes, defaultSingleByteEncoding)
                            {
                                cpStart = cp,
                                cpEnd = cpNext
                            };

                            //add pcd
                            this.Pieces.Add(pcd);

                            //add positions
                            int f = (int)pcd.fc;
                            int multi = 1;
                            if (pcd.encoding == Encoding.Unicode)
                            {
                                multi = 2;
                            }
                            for (int c = pcd.cpStart; c < pcd.cpEnd; c++)
                            {
                                if (!this.FileCharacterPositions.ContainsKey(c))
                                    this.FileCharacterPositions.Add(c, f);
                                if (!this.CharacterPositions.ContainsKey(f))
                                    this.CharacterPositions.Add(f, c);

                                f += multi;
                            }
                        }
                        int maxCp = this.FileCharacterPositions.Count;
                        this.FileCharacterPositions.Add(maxCp, fib.fcMac);
                        this.CharacterPositions.Add(fib.fcMac, maxCp);

                        //piecetable was found
                        goon = false;
                    }
                    //entry is no piecetable so goon
                    else if (type == 1)
                    {
                        short cb = System.BitConverter.ToInt16(bytes, pos + 1);
                        pos = pos + 1 + 2 + cb;
                    }
                    else
                    {
                        goon = false;
                    }
                }
                catch(Exception)
                {
                    goon = false;
                    
                }
            }
        }

        public List<char> GetAllChars(VirtualStream wordStream)
        {
            var chars = new List<char>();
            foreach (var pcd in this.Pieces)
            {
                //get the FC end of this piece
                int pcdFcEnd = pcd.cpEnd - pcd.cpStart;
                if (pcd.encoding == Encoding.Unicode)
                    pcdFcEnd *= 2;
                pcdFcEnd += (int)pcd.fc;

                int cb = pcdFcEnd - (int)pcd.fc;
                var bytes = new byte[cb];

                //read all bytes 
                wordStream.Read(bytes, 0, cb, (int)pcd.fc);

                //get the chars
                var plainChars = DecodeChars(pcd, bytes);

                //add to list
                foreach (char c in plainChars)
                {
                    chars.Add(c);
                }
            }
            return chars;
        }

        public List<char> GetChars(int fcStart, int fcEnd, VirtualStream wordStream)
        {
            var chars = new List<char>();
            for (int i = 0; i < this.Pieces.Count; i++)
            {
                var pcd = this.Pieces[i];

                //get the FC end of this piece
                int pcdFcEnd = pcd.cpEnd - pcd.cpStart;
                if (pcd.encoding == Encoding.Unicode)
                    pcdFcEnd *= 2;
                pcdFcEnd += (int)pcd.fc;

                if (pcdFcEnd < fcStart)
                {
                    //this piece is before the requested range
                    continue;
                }
                else if (fcStart >= pcd.fc && fcEnd > pcdFcEnd)
                {
                    //requested char range starts at this piece
                    //read from fcStart to pcdFcEnd

                    //get count of bytes
                    int cb = pcdFcEnd - fcStart;
                    var bytes = new byte[cb];

                    //read all bytes
                    wordStream.Read(bytes, 0, cb, (int)fcStart);

                    //get the chars
                    var plainChars = DecodeChars(pcd, bytes);

                    //add to list
                    foreach (char c in plainChars)
                    {
                        chars.Add(c);
                    }
                }
                else if (fcStart <= pcd.fc && fcEnd >= pcdFcEnd)
                {
                    //the full piece is part of the requested range
                    //read from pc.fc to pcdFcEnd

                    //get count of bytes
                    int cb = pcdFcEnd - (int)pcd.fc;
                    var bytes = new byte[cb];

                    //read all bytes 
                    wordStream.Read(bytes, 0, cb, (int)pcd.fc);

                    //get the chars
                    var plainChars = DecodeChars(pcd, bytes);

                    //add to list
                    foreach (char c in plainChars)
                    {
                        chars.Add(c);
                    }
                }
                else if (fcStart < pcd.fc && fcEnd >= pcd.fc && fcEnd <= pcdFcEnd)
                {
                    //requested char range ends at this piece
                    //read from pcd.fc to fcEnd

                    //get count of bytes
                    int cb = fcEnd - (int)pcd.fc;
                    var bytes = new byte[cb];

                    //read all bytes 
                    wordStream.Read(bytes, 0, cb, (int)pcd.fc);

                    //get the chars
                    var plainChars = DecodeChars(pcd, bytes);

                    //add to list
                    foreach (char c in plainChars)
                    {
                        chars.Add(c);
                    }

                    break;
                }
                else if (fcStart >= pcd.fc && fcEnd <= pcdFcEnd)
                {
                    //requested chars are completly in this piece
                    //read from fcStart to fcEnd

                    //get count of bytes
                    int cb = fcEnd - fcStart;

                    if (cb <= 0)
                    {
                        return new List<char>();
                    }
                    var bytes = new byte[cb];

                    //read all bytes 
                    wordStream.Read(bytes, 0, cb, (int)fcStart);

                    //get the chars
                    var plainChars = DecodeChars(pcd, bytes);

                    //set the list
                    chars = new List<char>(plainChars);

                    break;
                }
                else if (fcEnd < pcd.fc)
                {
                    //this piece is beyond the requested range
                    break;
                }
            }
            return chars;
        }

        public static Encoding ResolveSingleByteEncoding(FileInformationBlock fib, DocumentProperties documentProperties)
        {
            Encoding encoding = TryGetSingleByteEncoding(documentProperties?.cpgText ?? 0);
            if (encoding != null)
            {
                return encoding;
            }

            encoding = TryGetEncodingFromLcid(fib?.lid ?? 0);
            if (encoding != null)
            {
                return encoding;
            }

            encoding = TryGetEncodingFromLcid((ushort)(fib?.lidFE ?? 0));
            if (encoding != null)
            {
                return encoding;
            }

            return GetSingleByteEncoding(null);
        }

        private static Encoding GetSingleByteEncoding(Encoding singleByteEncoding)
        {
            return singleByteEncoding != null && singleByteEncoding.IsSingleByte
                ? singleByteEncoding
                : TryGetSingleByteEncoding(1252) ?? Encoding.GetEncoding("ISO-8859-1");
        }

        private static Encoding TryGetEncodingFromLcid(ushort lcid)
        {
            if (lcid == 0 || lcid == 0x0400)
            {
                return null;
            }

            try
            {
                int ansiCodePage = CultureInfo.GetCultureInfo(lcid).TextInfo.ANSICodePage;
                return TryGetSingleByteEncoding(ansiCodePage);
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static Encoding TryGetSingleByteEncoding(int codePage)
        {
            if (codePage <= 0)
            {
                return null;
            }

            try
            {
#if !NET462
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
                Encoding encoding = Encoding.GetEncoding(codePage);
                return encoding.IsSingleByte ? encoding : null;
            }
            catch
            {
                return null;
            }
        }

        private static char[] DecodeChars(PieceDescriptor pcd, byte[] bytes)
        {
            string text = pcd.encoding.GetString(bytes);

            if (pcd.encoding.CodePage == 10000)
            {
                string windows1252Text = DecodeWithSingleByteEncoding(bytes, 1252);
                if (LooksLikeMacintoshSmartPunctuationMismatch(text, windows1252Text))
                {
                    text = windows1252Text;
                }
            }

            // Some documents omit Central European metadata and default to 1252,
            // but a few pieces still contain 1250 bytes. Keep this fallback very
            // narrow to avoid changing normal Western European text.
            if (pcd.encoding.CodePage == 1252 || pcd.encoding.CodePage == 28591)
            {
                string cyrillicText = DecodeWithSingleByteEncoding(bytes, 1251);
                if (LooksLikeWesternMojibakeForCyrillic(text, cyrillicText))
                {
                    text = cyrillicText;
                }
                else
                {
                    string utf8Text = DecodeWithUtf8(bytes);
                    if (LooksLikeUtf8Mojibake(text, utf8Text))
                    {
                        text = utf8Text;
                    }
                }

                if (HasEmbeddedCharacter(text, 'ø', 'Ø'))
                {
                    Encoding centralEuropeanEncoding = TryGetSingleByteEncoding(1250);
                    if (centralEuropeanEncoding != null)
                    {
                        string centralEuropeanText = centralEuropeanEncoding.GetString(bytes);
                        if (HasEmbeddedCharacter(centralEuropeanText, 'ř', 'Ř'))
                        {
                            text = centralEuropeanText;
                        }
                    }
                }
            }

            return text.ToCharArray();
        }

        private static string DecodeWithSingleByteEncoding(byte[] bytes, int codePage)
        {
            Encoding encoding = TryGetSingleByteEncoding(codePage);
            return encoding != null ? encoding.GetString(bytes) : string.Empty;
        }

        private static string DecodeWithUtf8(byte[] bytes)
        {
            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                return string.Empty;
            }
        }

        private static bool HasEmbeddedCharacter(string text, char lower, char upper)
        {
            for (int i = 0; i < text.Length; i++)
            {
                char current = text[i];
                if (current != lower && current != upper)
                {
                    continue;
                }

                bool hasLetterBefore = i > 0 && char.IsLetter(text[i - 1]);
                bool hasLetterAfter = i + 1 < text.Length && char.IsLetter(text[i + 1]);
                if (hasLetterBefore && hasLetterAfter)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool LooksLikeMacintoshSmartPunctuationMismatch(string macintoshText, string windows1252Text)
        {
            int improvements = 0;
            int count = Math.Min(macintoshText.Length, windows1252Text.Length);

            for (int i = 0; i < count; i++)
            {
                if (!IsSmartPunctuation(windows1252Text[i]))
                {
                    continue;
                }

                if (IsSuspiciousMacintoshPunctuation(macintoshText[i]))
                {
                    improvements++;
                }
            }

            return improvements > 0;
        }

        private static bool LooksLikeWesternMojibakeForCyrillic(string westernText, string cyrillicText)
        {
            int westernLetters = CountMatchingChars(westernText, char.IsLetter);
            if (westernLetters < 6)
            {
                return false;
            }

            int asciiLetters = CountMatchingChars(westernText, c => c <= 0x7F && char.IsLetter(c));
            int supplementLetters = CountMatchingChars(westernText, c => c >= 0x00C0 && c <= 0x00FF && char.IsLetter(c));
            int cyrillicLetters = CountMatchingChars(cyrillicText, IsCyrillicLetter);

            return supplementLetters >= 6 &&
                   asciiLetters <= Math.Max(1, westernLetters / 5) &&
                   cyrillicLetters * 10 >= westernLetters * 7;
        }

        private static bool LooksLikeUtf8Mojibake(string singleByteText, string utf8Text)
        {
            if (string.IsNullOrEmpty(utf8Text) || string.Equals(singleByteText, utf8Text, StringComparison.Ordinal))
            {
                return false;
            }

            int suspiciousBefore = CountUtf8MojibakeSequences(singleByteText);
            if (suspiciousBefore < 2)
            {
                return false;
            }

            int suspiciousAfter = CountUtf8MojibakeSequences(utf8Text);
            if (suspiciousAfter >= suspiciousBefore)
            {
                return false;
            }

            int latinSupplementLetters = CountMatchingChars(utf8Text, c => c >= '\u00C0' && c <= '\u024F' && char.IsLetter(c));
            return latinSupplementLetters >= Math.Min(2, suspiciousBefore);
        }

        private static int CountUtf8MojibakeSequences(string text)
        {
            int count = 0;
            for (int i = 0; i + 1 < text.Length; i++)
            {
                char current = text[i];
                if (current != 'Ã' && current != 'Â' && current != 'â')
                {
                    continue;
                }

                char next = text[i + 1];
                if ((next >= '\u0080' && next <= '\u00BF') || next == '€' || next == '™')
                {
                    count++;
                }
            }

            return count;
        }

        private static int CountMatchingChars(string text, Func<char, bool> predicate)
        {
            int count = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (predicate(text[i]))
                {
                    count++;
                }
            }

            return count;
        }

        private static bool IsCyrillicLetter(char c)
        {
            return (c >= '\u0400' && c <= '\u04FF') && char.IsLetter(c);
        }

        private static bool IsSmartPunctuation(char c)
        {
            switch (c)
            {
                case '\u2018':
                case '\u2019':
                case '\u201C':
                case '\u201D':
                case '\u2013':
                case '\u2014':
                case '\u2026':
                    return true;

                default:
                    return false;
            }
        }

        private static bool IsSuspiciousMacintoshPunctuation(char c)
        {
            switch (c)
            {
                case 'ë':
                case 'í':
                case 'ì':
                case 'î':
                case 'ñ':
                case 'ó':
                    return true;

                default:
                    return false;
            }
        }
    }
}

using System.Collections.Generic;
using System.Text.RegularExpressions;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.txt.TextMapping
{
    public class Field : IVisitable
    {
        public string FieldCode;

        public string FieldExpansion;

        private Regex classicFieldFormat = new Regex(@"^(" + TextMark.FieldBeginMark + ")(.*)(" + TextMark.FieldSeperator + ")(.*)(" + TextMark.FieldEndMark + ")");
        
        private Regex shortFieldFormat = new Regex(@"^(" + TextMark.FieldBeginMark + ")(.*)(" + TextMark.FieldEndMark + ")");

        public Field(char[] fieldChars)
        {
            parse(new string(fieldChars));
        }

        public Field(List<char> fieldChars)
        {
            parse(new string(fieldChars.ToArray()));
        }

        public Field(string fieldString)
        {
            parse(fieldString);
        }

        private void parse(string field)
        {
            if (classicFieldFormat.IsMatch(field))
            {
                var classic = classicFieldFormat.Match(field);
                FieldCode = classic.Groups[2].Value;
                FieldExpansion = classic.Groups[4].Value;
            }
            else if (shortFieldFormat.IsMatch(field))
            {
                var shortField = shortFieldFormat.Match(field);
                FieldCode = shortField.Groups[2].Value;
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<Field>)mapping).Apply(this);
        }

        #endregion
    }
}

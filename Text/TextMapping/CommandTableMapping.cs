using System;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using System.Xml;
using System.IO;
using b2xtranslator.txt.TextModel;

namespace b2xtranslator.txt.TextMapping
{
    public class CommandTableMapping : AbstractMapping,
        IMapping<CommandTable>
    {
        private CommandTable _tcg;
        private ConversionContext _ctx;

        public CommandTableMapping(ConversionContext ctx, IWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void Apply(CommandTable tcg)
        {
            _tcg = tcg;
            _writer.WriteStartElement("wne", "tcg", OpenXmlNamespaces.MicrosoftWordML);

            //write the keymaps
            _writer.WriteStartElement("wne", "keymaps", OpenXmlNamespaces.MicrosoftWordML);
            for (int i = 0; i < tcg.KeyMapEntries.Count; i++)
            {
                writeKeyMapEntry(tcg.KeyMapEntries[i]);
            }
            _writer.WriteEndElement();

 
            _writer.WriteEndElement();

            _writer.Flush();
        }





        private void writeKeyMapEntry(KeyMapEntry kme)
        {
            _writer.WriteStartElement("wne", "keymap", OpenXmlNamespaces.MicrosoftWordML);

            //primary KCM
            if (kme.kcm1 > 0)
            {
                _writer.WriteAttributeString("wne", "kcmPrimary",
                    OpenXmlNamespaces.MicrosoftWordML,
                    string.Format("{0:x4}", kme.kcm1));
            }

            _writer.WriteStartElement("wne", "macro", OpenXmlNamespaces.MicrosoftWordML);

            _writer.WriteAttributeString("wne", "macroName",
                OpenXmlNamespaces.MicrosoftWordML,
                _tcg.MacroNames[kme.paramCid.ibstMacro]
                );

            _writer.WriteEndElement();

            _writer.WriteEndElement();
        }
    }
}

using System;
using b2xtranslator.StructuredStorage.Common;
using Microsoft.Extensions.Logging;

namespace b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Encapsulates a directory entry
    /// Author: math
    /// </summary>
    public class DirectoryEntry : AbstractDirectoryEntry
    {
        InputHandler _fileHandler;
        Header _header;
        ILogger _logger;
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        /// <param name="sid">The sid of the directory entry</param>
        internal DirectoryEntry(Header header, InputHandler fileHandler, uint sid, string path, ILogger logger) : base(sid)
        {
            this._header = header;
            this._fileHandler = fileHandler;
            _logger = logger;
            //_sid = sid;            
            ReadDirectoryEntry();
            this._path = path;
        }


        /// <summary>
        /// Reads the values of the directory entry. The position of the file handler must be at the start of a directory entry.
        /// </summary>
        private void ReadDirectoryEntry()
        {
            string rawName = this._fileHandler.ReadString(64);
            ushort lengthOfName = this._fileHandler.ReadUInt16();
            int nameLength = (lengthOfName / 2) - 1;
            if (nameLength > rawName.Length)
                nameLength = rawName.Length;
            this.Name = rawName.Substring(0, nameLength);

            // Name length check: lengthOfName = length of the element in bytes including Unicode NULL
            // Commented out due to trouble with odd unicode-named streams in PowerPoint -- flgr
            /*if (lengthOfName != (_name.Length + 1) * 2)
            {
                throw new InvalidValueInDirectoryEntryException("_cb");
            }*/
            // Added warning - math
            if (lengthOfName != (this._name.Length + 1) * 2)
            {
                _logger.LogWarning("Length of the name (_cb) of stream '" + this.Name + "' is not correct.");
            }


            this.Type = (DirectoryEntryType)this._fileHandler.ReadByte();
            this.Color = (DirectoryEntryColor)this._fileHandler.ReadByte();
            this.LeftSiblingSid = this._fileHandler.ReadUInt32();
            this.RightSiblingSid = this._fileHandler.ReadUInt32();
            this.ChildSiblingSid = this._fileHandler.ReadUInt32();

            var array = new byte[16];
            this._fileHandler.Read(array);
            this.ClsId = new Guid(array);

            this.UserFlags = this._fileHandler.ReadUInt32();
            // Omit creation time
            this._fileHandler.ReadUInt64();
            // Omit modification time 
            this._fileHandler.ReadUInt64();
            this.StartSector = this._fileHandler.ReadUInt32();

            uint sizeLow = this._fileHandler.ReadUInt32();
            uint sizeHigh = this._fileHandler.ReadUInt32();

            if (this._header.SectorSize == 512 && sizeHigh != 0x0)
            {
                // Must be zero according to the specification. However, this requirement can be ommited.
                _logger.LogWarning("ul_SizeHigh of stream '" + this.Name + "' should be zero as sector size is 512.");
                sizeHigh = 0x0;
            }
            this.SizeOfStream = ((ulong)sizeHigh << 32) + sizeLow;
        }
    }
}

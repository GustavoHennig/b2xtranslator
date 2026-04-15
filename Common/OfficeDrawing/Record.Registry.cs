using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    public partial class Record
    {
        private static void RegisterKnownRecordFactories()
        {
            Register(0xF000, (reader, size, typeCode, version, instance) => new DrawingGroup(reader, size, typeCode, version, instance));
            Register(0xF001, (reader, size, typeCode, version, instance) => new BlipStoreContainer(reader, size, typeCode, version, instance));
            Register(0xF002, (reader, size, typeCode, version, instance) => new DrawingContainer(reader, size, typeCode, version, instance));
            Register(0xF003, (reader, size, typeCode, version, instance) => new GroupContainer(reader, size, typeCode, version, instance));
            Register(0xF004, (reader, size, typeCode, version, instance) => new ShapeContainer(reader, size, typeCode, version, instance));
            Register(0xF005, (reader, size, typeCode, version, instance) => new SolverContainer(reader, size, typeCode, version, instance));
            Register(0xF006, (reader, size, typeCode, version, instance) => new DrawingGroupRecord(reader, size, typeCode, version, instance));
            Register(0xF007, (reader, size, typeCode, version, instance) => new BlipStoreEntry(reader, size, typeCode, version, instance));
            Register(0xF008, (reader, size, typeCode, version, instance) => new DrawingRecord(reader, size, typeCode, version, instance));
            Register(0xF009, (reader, size, typeCode, version, instance) => new GroupShapeRecord(reader, size, typeCode, version, instance));
            Register(0xF00A, (reader, size, typeCode, version, instance) => new Shape(reader, size, typeCode, version, instance));
            Register(new ushort[] { 0xF00B, 0xF121, 0xF122 }, (reader, size, typeCode, version, instance) => new ShapeOptions(reader, size, typeCode, version, instance));
            Register(0xF00D, (reader, size, typeCode, version, instance) => new ClientTextbox(reader, size, typeCode, version, instance));
            Register(0xF00F, (reader, size, typeCode, version, instance) => new ChildAnchor(reader, size, typeCode, version, instance));
            Register(0xF010, (reader, size, typeCode, version, instance) => new ClientAnchor(reader, size, typeCode, version, instance));
            Register(0xF011, (reader, size, typeCode, version, instance) => new ClientData(reader, size, typeCode, version, instance));
            Register(0xF012, (reader, size, typeCode, version, instance) => new FConnectorRule(reader, size, typeCode, version, instance));
            Register(0xF014, (reader, size, typeCode, version, instance) => new FArcRule(reader, size, typeCode, version, instance));
            Register(0xF017, (reader, size, typeCode, version, instance) => new FCalloutRule(reader, size, typeCode, version, instance));
            Register(new ushort[] { 0xF01A, 0xF01B, 0xF01C }, (reader, size, typeCode, version, instance) => new MetafilePictBlip(reader, size, typeCode, version, instance));
            Register(new ushort[] { 0xF01D, 0xF01E, 0xF01F, 0xF020, 0xF021 }, (reader, size, typeCode, version, instance) => new BitmapBlip(reader, size, typeCode, version, instance));
        }
    }
}

using System.Drawing;
using static WindowsIconUtility.Windows;

namespace WindowsIconUtility
{
    public enum IconTypeHint
    {
        ICON = 0,
        PNG = 1 // Allowed since Windows Vista
    }

    public class IconFile(string origin, Windows.GRPICONDIRENTRY header, IconTypeHint typeHint, byte[] data, uint dataLen)
    {
        public string Origin { get; } = origin;
        public GRPICONDIRENTRY Header { get; } = header;
        public IconTypeHint TypeHint { get; } = typeHint;
        public byte[] Data { get; } = data;
        public uint DataLen { get; } = dataLen;

        public bool SaveTo(Stream stream)
        {
            var writer = new BinaryWriter(stream);

            switch (TypeHint)
            {
                case IconTypeHint.ICON:
                    // Prepare ICO header (ICONDIR + 1 ICONDIRENTRY)
                    // ICONDIR (6 bytes)
                    writer.Write((ushort)0);      // Reserved
                    writer.Write((ushort)1);      // Type = 1 (icon)
                    writer.Write((ushort)1);      // One image

                    // ICONDIRENTRY (16 bytes)
                    writer.Write(Header.Width);     // Width
                    writer.Write(Header.Height);    // Height
                    writer.Write(Header.ColorCount);
                    writer.Write(Header.Reserved);
                    writer.Write(Header.Planes);
                    writer.Write(Header.BitCount);
                    writer.Write((uint)DataLen); // Size of image data
                    writer.Write((uint)22);       // Offset to image data (6 + 16)

                    writer.Write(Data);

                    break;
                case IconTypeHint.PNG:
                    writer.Write(Data);
                    break;
                default:
                    return false;
            }

            writer.Seek(0, SeekOrigin.Begin);

            return true;
        }

        public bool SaveTo(string filepath, FileMode mode = FileMode.Create)
        {
            if (TypeHint == IconTypeHint.ICON && Path.GetExtension(filepath) != ".ico")
                filepath = Path.ChangeExtension(filepath, "ico");

            if (TypeHint == IconTypeHint.PNG && Path.GetExtension(filepath) != ".png")
                filepath = Path.ChangeExtension(filepath, "png");

            using (var file = new FileStream(filepath, mode))
                return SaveTo(file);
        }

        public Stream GetStream()
        {
            var stream = new MemoryStream();

            SaveTo(stream);

            stream.Seek(0, SeekOrigin.Begin);

            return stream;
        }

        public Size GetSize()
        {
            using var bmp = Bitmap.FromStream(GetStream());
            return bmp.Size;
        }
    }
}

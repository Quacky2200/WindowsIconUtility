/**
 * WindowsIconUtility
 * Copyright (C) 2025 Quacky2200
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 **/

using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace WindowsIconUtility
{
    public class Windows
    {
        ///////////////////////////////////////////////////////////////////////
        // Executable Icons                                                  //
        ///////////////////////////////////////////////////////////////////////

        /*[DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        public static extern uint ExtractIconEx(
            string lpszFile,
            int nIconIndex,
            IntPtr[] phiconLarge,
            IntPtr[] phiconSmall,
            uint nIcons);*/

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool DestroyIcon(IntPtr hIcon);

        internal const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
        internal static readonly IntPtr RT_GROUP_ICON = new IntPtr(14);
        internal static readonly IntPtr RT_ICON = new IntPtr(3);

        internal delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll")]
        internal static extern bool EnumResourceNames(IntPtr hModule, IntPtr lpszType, EnumResNameProc lpEnumFunc, IntPtr lParam);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr FindResource(IntPtr hModule, IntPtr lpName, IntPtr lpType);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr LockResource(IntPtr hResData);

        [DllImport("kernel32.dll")]
        internal static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll")]
        internal static extern bool FreeLibrary(IntPtr hModule);

        [StructLayout(LayoutKind.Sequential, Pack = 2)]
        public struct GRPICONDIRENTRY
        {
            // WARNING: Do not base width and height decisions on these as they can be wrong (especially PNGs)
            public byte Width;
            public byte Height;
            public byte ColorCount;
            public byte Reserved;
            public ushort Planes;
            public ushort BitCount;
            public uint BytesInRes;
            public ushort ID; // Resource ID for the actual RT_ICON
        }

        public static GRPICONDIRENTRY GetGrpIconEntry(byte[] buffer, int offset)
        {
            return MemoryMarshal.Read<GRPICONDIRENTRY>(buffer.AsSpan(offset));
        }

        ///////////////////////////////////////////////////////////////////////
        // File Icons                                                        //
        ///////////////////////////////////////////////////////////////////////

        [Flags]
        public enum SIIGBF
        {
            /// <summary>
            /// Resize the image to fit the requested size.
            /// </summary>
            SIIGBF_RESIZETOFIT = 0x00000000,

            /// <summary>
            /// Return a larger image if available, even if it exceeds the requested size.
            /// </summary>
            SIIGBF_BIGGERSIZEOK = 0x00000001,

            /// <summary>
            /// Only use memory cache, don't access disk.
            /// </summary>
            SIIGBF_MEMORYONLY = 0x00000002,

            /// <summary>
            /// Return only the icon, not a thumbnail.
            /// </summary>
            SIIGBF_ICONONLY = 0x00000004,

            /// <summary>
            /// Return only the thumbnail, not an icon.
            /// </summary>
            SIIGBF_THUMBNAILONLY = 0x00000008,

            /// <summary>
            /// Only return the image if it's already in the cache.
            /// </summary>
            SIIGBF_INCACHEONLY = 0x00000010,

            /// <summary>
            /// Crop the image to a square aspect ratio.
            /// </summary>
            SIIGBF_CROPTOSQUARE = 0x00000020,

            /// <summary>
            /// Stretch the image to exactly match the requested size.
            /// </summary>
            SIIGBF_STRETCHTOFIT = 0x00000040,

            /// <summary>
            /// Return the image as a 32-bit bitmap with alpha channel (transparency).
            /// </summary>
            SIIGBF_SCALEUP = 0x00000080
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        internal static extern void SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object ppv);

        [ComImport]
        [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellItemImageFactory
        {
            [PreserveSig]
            int GetImage(SIZE size, int flags, out IntPtr phbm);
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SIZE
        {
            public int cx;
            public int cy;
            public SIZE(int x, int y) => (cx, cy) = (x, y);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAP
        {
            public int bmType;
            public int bmWidth;
            public int bmHeight;
            public int bmWidthBytes;
            public ushort bmPlanes;
            public ushort bmBitsPixel;
            public IntPtr bmBits;
        }

        [DllImport("gdi32.dll")]
        public static extern int GetObject(IntPtr hObject, int nCount, ref BITMAP lpObject);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern int GetDIBits(
            IntPtr hdc,                    // Handle to device context
            IntPtr hBitmap,                // Handle to bitmap
            uint uStartScan,              // First scan line to retrieve
            uint cScanLines,              // Number of scan lines to retrieve
            IntPtr lpvBits,               // Pointer to buffer for bitmap bits
            ref BITMAPINFO lpbi,          // Pointer to BITMAPINFO structure
            uint uUsage                   // Format of bmiColors (DIB_RGB_COLORS or DIB_PAL_COLORS)
        );


        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER
        {
            public uint biSize;
            public int biWidth;
            public int biHeight;
            public ushort biPlanes;
            public ushort biBitCount;
            public uint biCompression;
            public uint biSizeImage;
            public int biXPelsPerMeter;
            public int biYPelsPerMeter;
            public uint biClrUsed;
            public uint biClrImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RGBQUAD
        {
            public byte rgbBlue;
            public byte rgbGreen;
            public byte rgbRed;
            public byte rgbReserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFO
        {
            public BITMAPINFOHEADER bmiHeader;

            // This is a placeholder for the color table.
            // For 24-bit or 32-bit images, you typically don't need it.
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
            public RGBQUAD[] bmiColors;
        }

        public static Bitmap GetDIB(IntPtr hBitmap)
        {
            using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr hdc = graphics.GetHdc();
                BITMAP bmp = new BITMAP();

                GetObject(hBitmap, Marshal.SizeOf(typeof(BITMAP)), ref bmp);

                // Setup BITMAPINFO with negative height for top-down
                BITMAPINFO bmi = new BITMAPINFO();
                bmi.bmiHeader.biSize = (uint)Marshal.SizeOf(typeof(BITMAPINFOHEADER));
                bmi.bmiHeader.biWidth = bmp.bmWidth;
                bmi.bmiHeader.biHeight = -bmp.bmHeight;
                bmi.bmiHeader.biPlanes = bmp.bmPlanes;
                bmi.bmiHeader.biBitCount = bmp.bmBitsPixel;
                bmi.bmiHeader.biCompression = 0; // BI_RGB

                int bytesPerPixel = bmp.bmBitsPixel / 8;
                int stride = ((bmp.bmWidth * bytesPerPixel + 3) & ~3); // DWORD aligned
                int imageSize = stride * bmp.bmHeight;

                IntPtr pixelData = Marshal.AllocHGlobal(imageSize);

                int result = GetDIBits(hdc, hBitmap, 0, (uint)bmp.bmHeight, pixelData, ref bmi, 0);
                graphics.ReleaseHdc(hdc);

                if (result == 0)
                {
                    Marshal.FreeHGlobal(pixelData);
                    throw new InvalidOperationException("GetDIBits failed.");
                }

                var pixelFormat = bmp.bmBitsPixel == 32 ? PixelFormat.Format32bppArgb :
                                  bmp.bmBitsPixel == 24 ? PixelFormat.Format24bppRgb :
                                  PixelFormat.Format16bppRgb565;

                Bitmap bitmap = new Bitmap(bmp.bmWidth, bmp.bmHeight, stride, pixelFormat, pixelData);

                // Copy into safe memory and release GDI unmanaged
                Bitmap final = new Bitmap(bitmap);
                Marshal.FreeHGlobal(pixelData);
                bitmap.Dispose();

                return final;
            }
        }
    }
}

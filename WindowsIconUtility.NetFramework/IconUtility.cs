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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using static WindowsIconUtility.Windows;

namespace WindowsIconUtility
{
    public class InvalidIconEntryException : Exception
    {
        public InvalidIconEntryException(string message) : base(message)
        {
        }
    }

    public class IconUtility
    {
        /// <summary>
        /// Set a minimum bit depth to filter unwanted icons.
        /// 
        /// 8 bits for colourful icons, 4 bits for 95 / NT style, 1 bit for toolbar B/W icons, 0 for all.
        /// </summary>
        public static int MinimumBitDepth = 4;

        /// <summary>
        /// Set a minimum icon size to filter any icons which may be undesirable
        /// </summary>
        public static int MinimumSize = 16;

        /// <summary>
        /// Only enumerate the first icon group to retrieve the executable icon.
        /// 
        /// Set to false to retrieve all enumerable icons. This could be undesirable.
        /// </summary>
        public static bool EnumerateFirstIconGroupOnly = true;

        /// <summary>
        /// Some icons, supposedly 256x256 may have missing width/height. If it's defined as an icon group - it should be safe.
        /// </summary>
        public static bool TreatBadIconsLightly = true;

        /// <summary>
        /// Some invalid bit depth icons may have a lower bit depth than intended (color dithering for low bitrate?).
        /// Set this lower than minimumBitDepth or 0 to unfilter these icons.
        /// </summary>
        public static int MinimumInvalidBitDepth = 8; // Eurotrucks should update their logos.

        /// <summary>
        /// Default sorting method.
        /// 
        /// Sort by PNG, highest bit depth first, followed by highest resolution.
        /// </summary>
        public static Func<List<IconFile>, IOrderedEnumerable<IconFile>> OrderMethod = l =>
        {
            // PNGs support higher width/height and may scale better than .ico - The choice is here to change if required.
            return l
            .OrderByDescending(i => i.TypeHint == IconTypeHint.PNG ? 1 : -1)
            .ThenByDescending(i => i.Header.BitCount)
            .ThenByDescending(i =>
            {
                // Nasty but most reliable method to get size information.
                var size = i.GetSize();
                return size.Width * size.Height;
            });
        };

        /// <summary>
        /// Whether to get associative icon for the file or it's thumbnail when possible
        /// </summary>
        public static bool GetFileThumbnails = true;

        /// <summary>
        /// If we fail to retrieve an executable's icon, use the associative file icon
        /// 
        /// We'd normally want this on by default but this allows you to disable fallback
        /// and can tell you when there's no detectable embedded icon.
        /// </summary>
        public static bool FallbackExeToIcon = true;

        private static bool IsValidIconFile(IconFile icon)
        {
            var entry = icon.Header;

            if (icon.TypeHint == IconTypeHint.PNG) return true; // No checks needed here.

            int width = entry.Width;
            int height = entry.Height;
            int bits = entry.BitCount;
            bool invalid = false;

            // Check invalid width / height. It could be valid.
            //
            // We will also filter if the bit depth is given and we've got a filter for invalid icons.
            //
            // This issue occurs due to various issues:
            //     - Large icons or PNGs that would likely cause byte width/height overflow
            //     - An icon group not having icon width and height set
            if (width == 0 && height == 0 && (bits == 0 || (bits > 0 && MinimumInvalidBitDepth >= 0 && bits >= MinimumInvalidBitDepth)))
            {
                invalid = true;
                if (TreatBadIconsLightly)
                {
                    // If it's listed, then it 'should be OK'
                    width = 256;
                    height = 256;
                    bits = 32;
                }
            }

            bool validDimensions =
                width >= 8 && height >= 8 &&
                width == height && width >= MinimumSize;
            // width <= 512 && height <= 512 &&
            // (width & (width - 1)) == 0 &&
            // (height & (height - 1)) == 0;

            bool validBitDepth = bits >= MinimumBitDepth; // entry.BitCount == 8 || entry.BitCount == 16 || entry.BitCount == 32;

            var method = "WindowsIconUtility.IconUtility.IsValidIconFile";
            var info = $"icon: {entry.Width}x{entry.Height}, {entry.BitCount}-bit (ID: {entry.ID})";

            if (validDimensions && validBitDepth)
            {
                Debug(method, "Valid " + info);
                return true;
            }
            else
            {
                if (invalid)
                {
                    Debug(method, "Invalid " + info);
                }
                else
                {
                    Debug(method, "Ignored " + info);
                }

                return false;
            }
        }

        private static void Debug(string method, string msg)
        {
            System.Diagnostics.Debug.WriteLine($"{method}():- ${msg}");
        }

        private static bool MagicBytesMatch(byte[] source, byte[] pattern, int offset = 0)
        {
            if (source.Length < pattern.Length) return false;

            for (int i = 0; i < pattern.Length; i++)
            {
                if (source[i + offset] != pattern[i + offset]) return false;
            }

            return true;
        }

        private static readonly byte[] PNGHeader = { 137, 80, 78, 71, 13, 10, 26, 10 };

        /// <summary>
        /// Generate an icon from a module handle using the group icon entry
        /// 
        /// You need the entry id to know which icon to extract.
        /// </summary>
        /// <param name="exePath">Executable path. Only used for reference in IconFile.</param>
        /// <param name="hModule">Module handle</param>
        /// <param name="entry">Icon Group Entry (icon set metadata)</param>
        /// <returns>True on successful save</returns>
        /// <exception cref="InvalidIconEntryException">The icon resource cannot be found</exception>
        internal static IconFile? GetIconFile(string exePath, IntPtr hModule, GRPICONDIRENTRY entry)
        {
            // Load icon raw resource
            IntPtr iconRes = FindResource(hModule, (IntPtr)entry.ID, (IntPtr)3); // RT_ICON = 3
            if (iconRes == IntPtr.Zero)
                throw new InvalidIconEntryException("Icon resource not found");

            IntPtr iconData = LoadResource(hModule, iconRes);
            IntPtr iconPtr = LockResource(iconData);
            uint iconSize = SizeofResource(hModule, iconRes);

            // Image data
            byte[] iconBytes = new byte[iconSize];
            Marshal.Copy(iconPtr, iconBytes, 0, (int)iconSize);

            IconTypeHint hint;
            if (MagicBytesMatch(iconBytes, PNGHeader))
                hint = IconTypeHint.PNG;
            else
                hint = IconTypeHint.ICON;

            var icon = new IconFile(exePath, entry, hint, iconBytes, iconSize);

            if (IsValidIconFile(icon)) return icon;
            else return null;
        }

        /// <summary>
        /// Generic method to get icons from an executable.
        /// </summary>
        /// <param name="exePath">Executable path</param>
        /// <remarks>This is a generic method. Use this if you want metadata, otherwise use the GetIcons methods.</remarks>
        /// <returns>List of IconFile which can have 0 results or return null. NOTE: Check for null before use!</returns>
        public static List<IconFile>? FindExeIcons(string exePath)
        {
            IntPtr hModule = LoadLibraryEx(exePath, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE);
            if (hModule == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load module {exePath}.");
                return null;
            }

            List<IconFile> iconFiles = new List<IconFile>();

            try
            {
                bool enumeratedFirstGroup = false;

                // Enumerate RT_GROUP_ICON resources
                EnumResourceNames(hModule, RT_GROUP_ICON, (h, type, name, param) =>
                {
                    if (enumeratedFirstGroup && EnumerateFirstIconGroupOnly) return false; // Only wanted the first set (most likely app icons)
                    enumeratedFirstGroup = true;

                    IntPtr grpRes = FindResource(hModule, name, RT_GROUP_ICON);
                    IntPtr grpData = LoadResource(hModule, grpRes);
                    IntPtr grpPtr = LockResource(grpData);
                    uint grpSize = SizeofResource(hModule, grpRes);
                    byte[] grpBytes = new byte[grpSize];
                    Marshal.Copy(grpPtr, grpBytes, 0, (int)grpSize);

                    bool isIcon = BitConverter.ToBoolean(grpBytes, 2);
                    ushort count = BitConverter.ToUInt16(grpBytes, 4);

                    if (!isIcon || count == 0) return false;

                    for (int i = 0; i < count; i++)
                    {
                        var grpIcon = GetGrpIconEntry(grpBytes, 6 + i * Marshal.SizeOf<GRPICONDIRENTRY>());

                        var icon = GetIconFile(exePath, h, grpIcon);

                        if (icon != null) iconFiles.Add(icon);
                    }

                    return true;
                }, IntPtr.Zero);
            }
            finally
            {
                FreeLibrary(hModule);
            }

            if (iconFiles.Count == 0 && FallbackExeToIcon) iconFiles = [ GetAssociativeFileIcon(exePath) ];

            return OrderMethod(iconFiles).ToList();
        }

        /// <summary>
        /// Get icons of an executable as a list of Icons
        /// </summary>
        /// <param name="exePath">EXE file to retrieve icons</param>
        /// <returns>List of icons or empty list</returns>
        public static List<Icon> GetExeIcons(string exePath)
        {
            List<IconFile>? iconFiles = FindExeIcons(exePath);

            List<Icon> iconObjs = [];

            if (iconFiles == null) return iconObjs; // Empty

            foreach (IconFile iconFile in iconFiles)
            {
                var ms = new MemoryStream();

                iconFile.SaveTo(ms);

                iconObjs.Add(new Icon(ms));
            }

            return iconObjs;
        }

        /// <summary>
        /// Save icons of an executable to a directory.
        /// </summary>
        /// <param name="exePath">EXE file to retrieve icons</param>
        /// <param name="filePathFormatter">Formatter callback (IconFile as arg)</param>
        /// <param name="mode">File creation mode</param>
        /// <returns>True on successful save for 1+ icons</returns>
        public static bool SaveExeIcons(string exePath, Func<IconFile, string> filePathFormatter, FileMode mode = FileMode.Create)
        {
            bool success = false;

            List<IconFile>? iconFiles = FindExeIcons(exePath);

            if (iconFiles == null || (iconFiles != null && iconFiles.Count == 0)) return false;

            if (iconFiles != null) foreach (IconFile iconFile in iconFiles)
            {
                var filepath = filePathFormatter(iconFile);

                if (iconFile.SaveTo(filepath) && !success) success = true;
            }

            return success;
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// Get best icons
        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static Bitmap ResizeToWidth(Image original, int targetWidth)
        {
            double ratio = (double)targetWidth / original.Width;
            int targetHeight = (int)(original.Height * ratio);

            var dest = new Bitmap(targetWidth, targetHeight);
            using (var g = Graphics.FromImage(dest))
            {
                g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                g.DrawImage(original, 0, 0, targetWidth, targetHeight);
            }

            return dest;
        }

        public static Bitmap? GetBestIcon(string filePath, int targetWidth, float downscaleThreshold = 1.25f, float upscaleThreshold = 0.1f)
        {
            var exe = getExePath(filePath);

            List<IconFile>? icons = (
                exe != null ?
                FindExeIcons(exe) :
                [ GetAssociativeFileIcon(filePath, Math.Min(1024, targetWidth)) ]
            );

            IconFile? best = null;
            int bestDiff = int.MaxValue;
            int bestWidth = 0;

            if (icons != null) foreach (var img in icons)
                {
                    int width = img.GetSize().Width;
                    int diff = Math.Abs(width - targetWidth);

                    if (best == null || diff < bestDiff || (diff == bestDiff && width > bestWidth))
                    {
                        best = img;
                        bestDiff = diff;
                        bestWidth = width;
                    }
                }

            if (best == null)
                return null;

            var image = new Bitmap(best.GetStream());

            var percOfTarget = (bestWidth / targetWidth);

            // If the image is too big or too small, downscale and upscale accordingly.
            if (percOfTarget > downscaleThreshold || (1.0f - percOfTarget) > upscaleThreshold)
                return ResizeToWidth(image, targetWidth);
            else
                return image;
        }

        public static IconFile? GetAssociativeFileIcon(string filePath, int size = 256)
        {
            var iidFactory = typeof(IShellItemImageFactory).GUID;
            SHCreateItemFromParsingName(filePath, IntPtr.Zero, ref iidFactory, out var nativeItem);

            //if (nativeItem.GetType() != typeof(IShellItemImageFactory))
            //    return null;
            int flags = (int)(SIIGBF.SIIGBF_RESIZETOFIT | (GetFileThumbnails ? 0 : SIIGBF.SIIGBF_ICONONLY));

            IShellItemImageFactory factory = (IShellItemImageFactory)nativeItem;
            factory.GetImage(new SIZE(size, size), flags, out var hBitmap);

            if (hBitmap == IntPtr.Zero)
                return null;

            using (var bitmap = GetDIB(hBitmap))
            {
                using (var ms = new MemoryStream())
                {
                    bitmap.Save(ms, ImageFormat.Png);

                    uint length = 0;
                    GRPICONDIRENTRY entry = new GRPICONDIRENTRY()
                    {
                        ID = 1,
                        Planes = 0,
                        BitCount = 32,
                    };

                    if (ms.Length < uint.MaxValue) entry.BytesInRes = (uint)ms.Length;
                    if (bitmap.Width < byte.MaxValue) entry.Width = (byte)bitmap.Width;
                    if (bitmap.Height < byte.MaxValue) entry.Height = (byte)bitmap.Height;
                    if (ms.Length < uint.MaxValue) length = (uint)ms.Length;

                    return new IconFile(filePath, entry, IconTypeHint.PNG, ms.ToArray(), length);
                }

            }
        }

        public static string? getExePath(string file)
        {
            if (Path.GetExtension(file).Equals(".exe", StringComparison.OrdinalIgnoreCase)) return file;

            if (Path.GetExtension(file).Equals(".lnk", StringComparison.OrdinalIgnoreCase))
            {
                var target = GetLnkTargetPath(file);

                if (target != null && Path.GetExtension(target).Equals(".exe", StringComparison.OrdinalIgnoreCase))
                    return target;
            }

            return null;
        }
        public static string? GetLnkTargetPath(string lnkPath)
        {
            // You'll only get 32-bit paths if building this app as x86. Make sure to untick 'Prefer x86'
            // The majority will be running x64 since Windows 10.
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");

            if (shellType != null)
            {
                dynamic shell = Activator.CreateInstance(shellType);
                dynamic shortcut = shell.CreateShortcut(lnkPath);

                string targetPath = shortcut.TargetPath;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);

                return targetPath;
            }

            return null;
        }
    }
}

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
using System.Diagnostics;
using WindowsIconUtility;

namespace WindowsIconUtilityTestApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_ShowIcons(string exePath)
        {
            Application.EnableVisualStyles();

            this.Controls.Clear();

            FlowLayoutPanel panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                Margin = new Padding(0),
                Padding = new Padding(0),
                MinimumSize = new System.Drawing.Size(10, 10)
            };

            panel.WrapContents = false;

            this.Controls.Add(panel);

            panel.Controls.Add(new Label()
            {
                Text = "Filename: " + Path.GetFileName(exePath),
                Width = this.Width,
                Margin = new Padding(10)
            });

            void AddIcon(Bitmap? bmp)
            {
                if (bmp == null) return;

                try
                {
                    // Resident Evil 3 gives me a Bitmap overflow from Icon. Instead saving to a MemoryStream and loading this through Image.FromStream works fine. Hmm.
                    //Bitmap bmp = new Bitmap(Image.FromStream(icon.GetStream()));
                    /*Bitmap.FromStream(icon.Save())
                    using (var fs = new FileStream("temp.ico", FileMode.Create))
                    {
                        icon.Save(fs);
                    }
                    var bitmap = icon.ToBitmap();*/

                    PictureBox pb = new PictureBox()
                    {
                        Width = Math.Min(bmp.Width, 512),
                        Height = Math.Min(bmp.Height, 512),
                        Image = bmp,
                        SizeMode = PictureBoxSizeMode.CenterImage,
                        BorderStyle = BorderStyle.FixedSingle,
                        Margin = new Padding(10),
                    };

                    pb.SizeMode = PictureBoxSizeMode.Zoom;
                    panel.Controls.Add(pb);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Unable to make picturebox due to error: {ex.Message}");
                }
            }

            AddIcon(IconUtility.GetBestIcon(exePath, 256));

            //var icons = IconUtility.GetExeIcons(exePath);
            /* if (icons != null) foreach (var icon in icons)
             {
                 AddIcon(icon);
             }*/
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] files)
            {
                Form1_ShowIcons(files[0]);
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Allow any file
                e.Effect = DragDropEffects.Copy;
                return;

                /*string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                var exe = files.Length > 0 ? IconUtility.getExePath(files[0]) : null;

                if (exe != null)
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }*/
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
    }
}

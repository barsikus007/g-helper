// Source thanks to https://github.com/vddCore/Starlight with some adjustments from me

using GHelper.AnimeMatrix.Communication;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Management;
using System.Text;

namespace Starlight.AnimeMatrix
{
    public class BuiltInAnimation
    {
        public enum Startup
        {
            GlitchConstruction,
            StaticEmergence
        }

        public enum Shutdown
        {
            GlitchOut,
            SeeYa
        }

        public enum Sleeping
        {
            BannerSwipe,
            Starfield
        }

        public enum Running
        {
            BinaryBannerScroll,
            RogLogoGlitch
        }

        public byte AsByte { get; }

        public BuiltInAnimation(
            Running running,
            Sleeping sleeping,
            Shutdown shutdown,
            Startup startup
        )
        {
            AsByte |= (byte)(((int)running & 0x01) << 0);
            AsByte |= (byte)(((int)sleeping & 0x01) << 1);
            AsByte |= (byte)(((int)shutdown & 0x01) << 2);
            AsByte |= (byte)(((int)startup & 0x01) << 3);
        }
    }

    internal class AnimeMatrixPacket : Packet
    {
        public AnimeMatrixPacket(byte[] command)
            : base(0x5E, 640, command)
        {
        }
    }

    public enum AnimeType
    {
        GA401,
        GA402,
        GU604
    }



    public enum BrightnessMode : byte
    {
        Off = 0,
        Dim = 1,
        Medium = 2,
        Full = 3
    }



    public class AnimeMatrixDevice : Device
    {
        int UpdatePageLength = 490;
        int LedCount = 1450;

        byte[] _displayBuffer;
        List<byte[]> frames = new List<byte[]>();

        public int MaxRows = 61;
        //public int FullRows = 11;
        //public int FullEvenRows = -1;

        public int dx = 0;
        public int MaxDiagonalRows = 36;
        public int MaxColumns = 34;
        public int LedStart = 0;

        public int TextShift = 8;

        private int frameIndex = 0;

        private static AnimeType _model = AnimeType.GA402;

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern IntPtr AddFontMemResourceEx(IntPtr pbFont, uint cbFont, IntPtr pdv, [System.Runtime.InteropServices.In] ref uint pcFonts);
        private PrivateFontCollection fonts = new PrivateFontCollection();

        public AnimeMatrixDevice() : base(0x0B05, 0x193B, 640)
        {
            string model = GetModel();

            if (model.Contains("401"))
            {
                _model = AnimeType.GA401;

                MaxColumns = 33;
                dx = 1;

                MaxRows = 55;
                MaxDiagonalRows = 36;
                LedCount = 1245;

                UpdatePageLength = 410;

                TextShift = 11;

                LedStart = 1;
            }

            if (model.Contains("GU604"))
            {
                _model = AnimeType.GU604;

                MaxColumns = 39;
                MaxRows = 92;
                LedCount = 1711;
                UpdatePageLength = 630;

                TextShift = 10;
            }

            _displayBuffer = new byte[LedCount];

            /*
            for (int i = 0; i < MaxRows; i++)
            {
                _model = AnimeType.GA401;
                Logger.WriteLine(FirstX(i) + " " + Pitch(i));
            }
            */

            LoadMFont();

        }

        private void LoadMFont()
        {
            byte[] fontData = GHelper.Properties.Resources.MFont;
            IntPtr fontPtr = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(fontData.Length);
            System.Runtime.InteropServices.Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
            uint dummy = 0;

            fonts.AddMemoryFont(fontPtr, GHelper.Properties.Resources.MFont.Length);
            AddFontMemResourceEx(fontPtr, (uint)GHelper.Properties.Resources.MFont.Length, IntPtr.Zero, ref dummy);
            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(fontPtr);
        }

        public string GetModel()
        {
            using (var searcher = new ManagementObjectSearcher(@"Select * from Win32_ComputerSystem"))
            {
                foreach (var process in searcher.Get())
                    return process["Model"].ToString();
            }

            return null;

        }

        public byte[] GetBuffer()
        {
            return _displayBuffer;
        }

        public void PresentNextFrame()
        {
            if (frameIndex >= frames.Count) frameIndex = 0;
            _displayBuffer = frames[frameIndex];
            Present();
            frameIndex++;
        }

        public void ClearFrames()
        {
            frames.Clear();
            frameIndex = 0;
        }

        public void AddFrame()
        {
            frames.Add(_displayBuffer.ToArray());
        }

        public void SendRaw(params byte[] data)
        {
            Set(Packet<AnimeMatrixPacket>(data));
        }


        public int Width()
        {
            switch (_model)
            {
                case AnimeType.GA401:
                    return 33;
                case AnimeType.GU604:
                    return 39;
                default:
                    return 34;
            }
        }

        public int FirstX(int y)
        {
            switch (_model)
            {
                case AnimeType.GA401:
                    if (y < 5 && y % 2 == 0)
                    {
                        return 1;
                    }
                    return (int)Math.Ceiling(Math.Max(0, y - 5) / 2F);
                case AnimeType.GU604:
                    if (y < 9 && y % 2 == 0)
                    {
                        return 1;
                    }
                    return (int)Math.Ceiling(Math.Max(0, y - 9) / 2F);

                default:
                    return (int)Math.Ceiling(Math.Max(0, y - 11) / 2F);
            }
        }


        public int Pitch(int y)
        {
            switch (_model)
            {
                case AnimeType.GA401:
                    switch (y)
                    {
                        case 0:
                        case 2:
                        case 4:
                            return 33;
                        case 1:
                        case 3:
                            return 35;
                        default:
                            return 36 - y / 2;
                    }

                case AnimeType.GU604:
                    switch (y)
                    {
                        case 0:
                        case 2:
                        case 4:
                        case 6:
                        case 8:
                            return 38;

                        case 1:
                        case 3:
                        case 5:
                        case 7:
                        case 9:
                            return 39;

                        default:
                            return Width() - FirstX(y);
                    }


                default:
                    return Width() - FirstX(y);
            }
        }


        public int RowToLinearAddress(int y)
        {
            int ret = LedStart;
            for (var i = 0; i < y; i++)
                ret += Pitch(i);

            return ret;
        }

        public void SetLedPlanar(int x, int y, byte value)
        {
            if (!IsRowInRange(y)) return;

            if (x >= FirstX(y) && x < Width())
                SetLedLinear(RowToLinearAddress(y) - FirstX(y) + x, value);
        }

        public void WakeUp()
        {
            Set(Packet<AnimeMatrixPacket>(Encoding.ASCII.GetBytes("ASUS Tech.Inc.")));
        }

        public void SetLedLinear(int address, byte value)
        {
            if (!IsAddressableLed(address)) return;
            _displayBuffer[address] = value;
        }

        public void SetLedLinearImmediate(int address, byte value)
        {
            if (!IsAddressableLed(address)) return;
            _displayBuffer[address] = value;

            Set(Packet<AnimeMatrixPacket>(0xC0, 0x02)
                .AppendData(BitConverter.GetBytes((ushort)(address + 1)))
                .AppendData(BitConverter.GetBytes((ushort)0x0001))
                .AppendData(value)
            );

            Set(Packet<AnimeMatrixPacket>(0xC0, 0x03));
        }



        public void Clear(bool present = false)
        {
            for (var i = 0; i < _displayBuffer.Length; i++)
                _displayBuffer[i] = 0;

            if (present)
                Present();
        }

        public void Present()
        {

            int page = 0;
            int start, end;

            while (page * UpdatePageLength < LedCount)
            {
                start = page * UpdatePageLength;
                end = Math.Min(LedCount, (page + 1) * UpdatePageLength);

                Set(Packet<AnimeMatrixPacket>(0xC0, 0x02)
                    .AppendData(BitConverter.GetBytes((ushort)(start + 1)))
                    .AppendData(BitConverter.GetBytes((ushort)(end - start)))
                    .AppendData(_displayBuffer[start..end])
                );

                page++;
            }

            Set(Packet<AnimeMatrixPacket>(0xC0, 0x03));
        }

        public void SetDisplayState(bool enable)
        {
            if (enable)
            {
                Set(Packet<AnimeMatrixPacket>(0xC3, 0x01)
                    .AppendData(0x00));
            }
            else
            {
                Set(Packet<AnimeMatrixPacket>(0xC3, 0x01)
                    .AppendData(0x80));
            }
        }

        public void SetBrightness(BrightnessMode mode)
        {
            Set(Packet<AnimeMatrixPacket>(0xC0, 0x04)
                .AppendData((byte)mode)
            );
        }

        public void SetBuiltInAnimation(bool enable)
        {
            var enabled = enable ? (byte)0x00 : (byte)0x80;
            Set(Packet<AnimeMatrixPacket>(0xC4, 0x01, enabled));
        }

        public void SetBuiltInAnimation(bool enable, BuiltInAnimation animation)
        {
            SetBuiltInAnimation(enable);
            Set(Packet<AnimeMatrixPacket>(0xC5, animation.AsByte));
        }


        public void PresentClock()
        {
            string second = (DateTime.Now.Second % 2 == 0) ? ":" : "  ";
            string time = DateTime.Now.ToString("HH" + second + "mm");

            Clear();
            TextDiagonal(time, 15, 12, TextShift + 11);
            TextDiagonal(DateTime.Now.ToString("yy'. 'MM'. 'dd"), 11.5F, 3, TextShift);
            Present();

        }

        public void TextDiagonal(string text, float fontSize = 10, int deltaX = 0, int deltaY = 10)
        {

            int maxX = (int)Math.Sqrt(MaxRows * MaxRows + MaxColumns * MaxColumns);
            int textHeight;

            using (Bitmap bmp = new Bitmap(maxX, MaxRows))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CompositingQuality = CompositingQuality.HighQuality;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.SingleBitPerPixel;

                    using (Font font = new Font(fonts.Families[0], fontSize, FontStyle.Regular, GraphicsUnit.Pixel))
                    {
                        SizeF textSize = g.MeasureString(text, font);
                        textHeight = (int)textSize.Height;
                        g.DrawString(text, font, Brushes.White, 0, 0);
                    }
                }

                for (int y = 0; y < bmp.Height; y++)
                {
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        var pixel = bmp.GetPixel(x, y);
                        var color = (pixel.R + pixel.G + pixel.B) / 3;
                        if (color > 100) SetLedDiagonal(x, y, (byte)color, deltaX, deltaY);
                    }
                }
                bmp.Save(@"D:\projects\g-helper\app\bin\test-diag.bmp");
            }
        }


        public void PresentText(string text1, string text2 = "")
        {
            using (Bitmap bmp = new Bitmap(MaxColumns * 3, MaxRows))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CompositingQuality = CompositingQuality.HighQuality;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.SingleBitPerPixel;

                    using (Font font = new Font("Consolas", 22F, FontStyle.Regular, GraphicsUnit.Pixel))
                    {
                        SizeF textSize = g.MeasureString(text1, font);
                        g.DrawString(text1, font, Brushes.White, (MaxColumns * 3 - textSize.Width) + 3, -4);
                    }

                    if (text2.Length > 0)
                        using (Font font = new Font("Consolas", 18F, GraphicsUnit.Pixel))
                        // using (Font font = new Font("ROG Fonts v1.5", 18F, GraphicsUnit.Pixel))
                        {
                            SizeF textSize = g.MeasureString(text2, font);
                            g.DrawString(text2, font, Brushes.White, (MaxColumns * 3 - textSize.Width) + 1, 25);
                        }

                }

                bmp.Save("test.bmp", ImageFormat.Bmp);

                GenerateFrame(bmp);
                Present();
                bmp.Save(@"D:\projects\g-helper\app\bin\test.bmp");
            }

        }

        public void GenerateFrame(Image image, float zoom = 100, int panX = 0, int panY = 0, InterpolationMode quality = InterpolationMode.Default)
        {

            int width = MaxColumns / 2 * 6;
            int height = MaxRows;

            int targetWidth = MaxColumns * 2;

            float scale;

            using (Bitmap bmp = new Bitmap(targetWidth, height))
            {
                scale = Math.Min((float)width / (float)image.Width, (float)height / (float)image.Height) * zoom / 100;

                using (var graph = Graphics.FromImage(bmp))
                {
                    var scaleWidth = (float)(image.Width * scale);
                    var scaleHeight = (float)(image.Height * scale);

                    graph.InterpolationMode = quality;
                    graph.CompositingQuality = CompositingQuality.HighQuality;
                    graph.SmoothingMode = SmoothingMode.AntiAlias;

                    graph.DrawImage(image, (float)Math.Round(targetWidth - (scaleWidth + panX) * targetWidth / width), -panY, (float)Math.Round(scaleWidth * targetWidth / width), scaleHeight);

                }

                for (int y = 0; y < bmp.Height; y++)
                {
                    for (int x = 0; x < bmp.Width; x++)
                        // if (x % 2 == y % 2)
                        {
                            var pixel = bmp.GetPixel(x, y);
                            var color = (pixel.R + pixel.G + pixel.B) / 3;
                            if (color < 10) color = 0;
                            // SetLedPlanar(x / 2, y, (byte)color);
                            SetLedDiagonalGA401(x, y, (byte)color);
                        }
                }
            }
        }


        public void SetLedDiagonal(int x, int y, byte color, int deltaX = 0, int deltaY = 10)
        {
            x += deltaX;
            y -= deltaY;

            int plX = (x - y) / 2;
            int plY = x + y;
            SetLedPlanar(plX, plY, color);
        }


        public void SetLedDiagonalGA401(int x, int y, byte color)
        {
            SetLedDiagonal(x, y, color, 0, 34);
        }


        private bool IsRowInRange(int row)
        {
            return (row >= 0 && row < MaxRows);
        }

        private bool IsAddressableLed(int address)
        {
            return (address >= 0 && address < LedCount);
        }
        
        public void PresentTextDiagonal(string text)
        {

            Clear();


            InstalledFontCollection installedFontCollection = new InstalledFontCollection();


            string familyName;
            string familyList = "";
            FontFamily[] fontFamilies;
            // Get the array of FontFamily objects.
            fontFamilies = installedFontCollection.Families;

            int count = fontFamilies.Length;
            for (int j = 0; j < count; ++j)
            {
                familyName = fontFamilies[j].Name;
                familyList = familyList + familyName;
                familyList = familyList + ",  ";
            }

            int maxX = (int)Math.Sqrt(MaxRows * MaxRows + MaxColumns * MaxColumns);

            using (Bitmap bmp = new Bitmap(maxX, MaxRows))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CompositingQuality = CompositingQuality.HighQuality;
                    g.SmoothingMode = SmoothingMode.AntiAlias;

                    using (Font font = new Font("Consolas", 13F, FontStyle.Regular, GraphicsUnit.Pixel))
                    {
                        SizeF textSize = g.MeasureString(text, font);
                        g.DrawString(text, font, Brushes.White, 4, 1);
                    }
                }

                for (int y = 0; y < bmp.Height; y++)
                {
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        var pixel = bmp.GetPixel(x, y);
                        var color = (pixel.R + pixel.G + pixel.B) / 3;
                        SetLedDiagonal(x, y, (byte)color);
                    }
                }
            }

            Present();
        }

        static (int, int) GetRowData(int row)
        {
            /*
                short - long - ...
                1   - 33
                34  - 66   67-68 do nothing
                69  - 101
                102 - 134  135-136 do nothing
                137 - 169
                170 - 202  203 do nothing
                204 - 236
                237 - 268  269 do nothing -1 starts
                270 - 301
                302 - 332  333 do nothing -1
                334 - 364
                ...
             */

            int row_leds = 33; // -1 every 2 rows starting from 8
            int row_modifier_ = row > 7 ? (row - 6) / 2 : 0;
            int row_modifier = (int)Math.Pow(row_modifier_, 2) + row_modifier_ * (row % 2);
            int last = row * row_leds + (row > 2 ? (row > 4) ? 4 : 2 : 0) + (row > 6 ? (row - 5) / 2 : 0) - row_modifier;
            int first = last - row_leds + 1 + row_modifier_;
            // Logger.WriteLine($"{row}: {first}-{last} {row_modifier_} {row_modifier}");
            return (first, last);
        }

        public void PresentBorder()
        {
            Clear();
            for (int row = 1; row <= 55; row++)
            {
                (int first, int last) = GetRowData(row);
                SetLedLinear(first, 255);
                SetLedLinear(last, 255);
                if (row == 1 || row == 55) for (int i = first + 1; i < last; i++) SetLedLinear(i, 255);
            }
            Present();
        }

        public void SetLedDiagonalZ(int x = 1, int y = 1, byte color = 255)
        {
            int row = x+3;
            (int first, int last) = GetRowData(row);
            int set_value = first + (row - y);
            if (row < 6) set_value--;
            if (row < 4) set_value--;
            if (row < 2) set_value--;
            if (y == row && y == 4) return;
            if (set_value > last) return;
            SetLedLinear(set_value, color);
        }

        public void PresentTextZ(string text)
        {
            // int second = DateTime.Now.Second;
            // string time;
            Clear();
            int kek = int.Parse(text);
            // for (int z = 1; z <= kek; z++)
            // {
            //     // Logger.WriteLine($"{z} 1 - {kek}");

            //     // SetLedLinear(z, (byte)(z % 256));
            //     SetLedLinear(z, 255);
            // }
            /*
                short - long - ...
                1   - 33
                34  - 66   67-68 do nothing
                69  - 101
                102 - 134  135-136 do nothing
                137 - 169
                170 - 202  203 do nothing
                204 - 236
                237 - 268  269 do nothing -1 starts
                270 - 301
                302 - 332  333 do nothing -1
                334 - 364
                ...
             */
            if (kek == 0)
            {
                for (int lol = 1; lol <= LedCount; lol++)
                {
                    SetLedLinear(lol, 255);
                }
                Present();
                return;
            }
            if (kek == -1)
            {
                for (int ledX = 1; ledX <= 60; ledX++) for (int ledY = 1; ledY <= 36; ledY++) SetLedDiagonalGA401(ledX, ledY, 255);
                Present();
                return;
            }
            // PresentBorder();
            kek += 3;
            // SetLedDiagonal(1, 33, 255, 0, 34);
            // SetLedDiagonal(1, 34, 255, 0, 34);
            // SetLedDiagonal(2, 35, 255, 0, 34);
            // for (int ledX = 3; ledX < 53; ledX++) SetLedDiagonal(ledX, 36, 255, 0, 34);
            // SetLedDiagonal(53, 35, 255, 0, 34);
            // SetLedDiagonal(54, 34, 255, 0, 34);
            // SetLedDiagonal(55, 33, 255, 0, 34);
            // SetLedDiagonal(56, 32, 255, 0, 34);
            // SetLedDiagonal(57, 31, 255, 0, 34);
            // SetLedDiagonal(58, 30, 255, 0, 34);
            // SetLedDiagonal(59, 29, 255, 0, 34);
            // SetLedDiagonal(60, 28, 255, 0, 34);
            for (int ledX = 1; ledX <= 60; ledX++) {
                for (int ledY = 1; ledY <= 36; ledY++) {
                    // if (((ledX*60 + ledY) % 2) == 0) SetLedDiagonal(ledX, ledY, 255, 0, 34);
                    // if (((ledX*60 + ledY) % 2) == 1) SetLedDiagonal(ledX, ledY, 255, 0, 34);
                    // if (((ledX + ledY*60) % 2) == 0) SetLedDiagonal(ledX, ledY, 255, 0, 34);
                    // if (((ledX + ledY*60) % 2) == 1) SetLedDiagonal(ledX, ledY, 255, 0, 34);
                    // checkmates led pattern
                    // if (((ledX + ledY) % 2) == 0) SetLedDiagonal(ledX, ledY, 255, 0, 34);
                    if (((ledX + ledY) % 2) == 1) SetLedDiagonalGA401(ledX, ledY, 255);
                }
            }
            for (int row = kek; row > 0; row--)
            {
                break;
                SetLedDiagonal(kek - 3, row, 255);
                continue;
                // 4 <= row <= 63  kek ~= 01-60
                (int first, int last) = GetRowData(row);
                int set_value = first + (kek - row);
                if (row < 6) set_value--;
                if (row < 4) set_value--;
                if (row < 2) set_value--;
                if (kek == row && kek == 4) continue;
                if (set_value > last) break;
                SetLedLinear(set_value, 255);
                // if (row == 1 || row == 55) for (int i = first + 1; i < last; i++) SetLedLinear(i, 255);
            }
            Present();
        }
    }
}
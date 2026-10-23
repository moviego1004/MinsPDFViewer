using System;
using System.Collections.Generic;
using System.Globalization;

namespace MinsPDFViewer
{
    /// <summary>용지에 페이지를 얹는 방식.</summary>
    public enum PrintScaleMode
    {
        /// <summary>인쇄 가능 영역에 맞춰 키우거나 줄임</summary>
        Fit,
        /// <summary>인쇄 가능 영역보다 큰 페이지만 줄임</summary>
        ShrinkOversized,
        /// <summary>100% (용지 가운데)</summary>
        ActualSize
    }

    public enum PrintOrientationMode
    {
        /// <summary>페이지마다 가로·세로를 페이지 모양에 맞춤</summary>
        Auto,
        Portrait,
        Landscape
    }

    /// <summary>1/100 인치 단위 사각형 (System.Drawing.Printing과 같은 단위). 원점 = 용지 왼쪽 위.</summary>
    public readonly record struct PrintRect(double X, double Y, double Width, double Height);

    /// <summary>인쇄 배치 계산. 미리보기와 실제 인쇄가 같은 함수를 쓴다 (화면 = 종이).</summary>
    public static class PrintLayout
    {
        /// <summary>PDF 포인트 → 1/100 인치</summary>
        public const double HundredthsPerPoint = 100.0 / 72.0;

        public static bool IsLandscape(double pageWidthPt, double pageHeightPt, PrintOrientationMode mode, bool defaultLandscape = false)
        {
            return mode switch
            {
                PrintOrientationMode.Portrait => false,
                PrintOrientationMode.Landscape => true,
                _ => pageWidthPt > pageHeightPt * 1.0001 || (Math.Abs(pageWidthPt - pageHeightPt) < 0.01 && defaultLandscape)
            };
        }

        /// <summary>
        /// 드라이버가 페이지별 용지 방향을 무시해서 원하는 방향과 실제 용지 방향이 다른지.
        /// 그러면 내용을 90° 돌려 찍는다 (가로 페이지가 세로 용지에 작게 나오는 대신 꽉 차게).
        /// </summary>
        public static bool NeedsContentRotation(double pageWidthPt, double pageHeightPt, double paperWidth, double paperHeight,
            PrintOrientationMode mode, bool defaultLandscape = false)
        {
            if (Math.Abs(paperWidth - paperHeight) < 1.0)
                return false;
            bool wantLandscape = IsLandscape(pageWidthPt, pageHeightPt, mode, defaultLandscape);
            bool paperLandscape = paperWidth > paperHeight;
            return wantLandscape != paperLandscape && Math.Abs(pageWidthPt - pageHeightPt) >= 0.01;
        }

        /// <summary>
        /// 페이지(포인트)를 용지(1/100 인치) 위에 놓을 사각형. 맞춤·축소는 인쇄 가능 영역 가운데,
        /// 실제 크기는 용지 가운데 (원본과 같은 위치·크기, 넘치면 잘림).
        /// </summary>
        public static PrintRect Compute(
            double pageWidthPt, double pageHeightPt,
            double paperWidth, double paperHeight,
            PrintRect printable,
            PrintScaleMode mode)
        {
            double pageW = Math.Max(0.01, pageWidthPt) * HundredthsPerPoint;
            double pageH = Math.Max(0.01, pageHeightPt) * HundredthsPerPoint;

            if (printable.Width <= 0 || printable.Height <= 0)
                printable = new PrintRect(0, 0, paperWidth, paperHeight);

            double fit = Math.Min(printable.Width / pageW, printable.Height / pageH);
            double scale = mode switch
            {
                PrintScaleMode.Fit => fit,
                PrintScaleMode.ShrinkOversized => Math.Min(1.0, fit),
                _ => 1.0
            };

            double w = pageW * scale;
            double h = pageH * scale;
            if (mode == PrintScaleMode.ActualSize)
                return new PrintRect((paperWidth - w) / 2.0, (paperHeight - h) / 2.0, w, h);

            return new PrintRect(
                printable.X + (printable.Width - w) / 2.0,
                printable.Y + (printable.Height - h) / 2.0,
                w, h);
        }
    }

    /// <summary>"1-3, 5, 8-" 같은 쪽 범위 (1부터, 쉼표·공백 구분). "-4" = 1~4, "8-" = 8~끝.</summary>
    public static class PrintPageRange
    {
        public static bool TryParse(string? text, int pageCount, out List<int> pageIndices, out string? error)
        {
            pageIndices = new List<int>();
            error = null;
            if (pageCount <= 0)
            {
                error = "페이지 없음";
                return false;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                error = "범위 없음";
                return false;
            }

            foreach (string rawPart in text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string part = rawPart.Trim().Replace('~', '-').Replace('–', '-');
                if (part.Length == 0)
                    continue;

                int start, end;
                int dash = part.IndexOf('-');
                if (dash < 0)
                {
                    if (!TryPage(part, out start))
                    {
                        error = $"'{part}'";
                        return false;
                    }
                    end = start;
                }
                else
                {
                    string left = part.Substring(0, dash).Trim();
                    string right = part.Substring(dash + 1).Trim();
                    start = 1;
                    end = pageCount;
                    if (left.Length > 0 && !TryPage(left, out start))
                    {
                        error = $"'{part}'";
                        return false;
                    }
                    if (right.Length > 0 && !TryPage(right, out end))
                    {
                        error = $"'{part}'";
                        return false;
                    }
                }

                if (start < 1 || end < 1 || start > pageCount || end > pageCount)
                {
                    error = $"1~{pageCount}";
                    return false;
                }

                if (start <= end)
                {
                    for (int p = start; p <= end; p++)
                        pageIndices.Add(p - 1);
                }
                else
                {
                    // "5-3" = 역순
                    for (int p = start; p >= end; p--)
                        pageIndices.Add(p - 1);
                }
            }

            if (pageIndices.Count == 0)
            {
                error = "범위 없음";
                return false;
            }

            return true;
        }

        private static bool TryPage(string s, out int page)
        {
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out page);
        }
    }
}

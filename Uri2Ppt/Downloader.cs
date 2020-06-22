using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;
using System.Threading.Tasks;
using NetOffice.PowerPointApi;
using ImageResizer;

namespace Uri2Ppt
{
    public static class Downloader
    {
        static Downloader()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        public static Task<int> GetLastRow(string fileName)
        {
            return Task.Run(() => {
                using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                    return worksheet.Dimension.End.Row;
                }                   
            });
        }
        public static async Task<List<DownloadedItem>> ReadAndWrite(ViewMainModel model, IProgress<int> progress)
        {
            string tempDir = Path.GetDirectoryName(model.OpenedFile);
            tempDir = Path.Combine(tempDir, "TEMP");
            if (!Directory.Exists(tempDir))
                Directory.CreateDirectory(tempDir);

            int rowsCount = model.RowFinish - model.RowStart + 1;
            List<DownloadedItem> items = new List<DownloadedItem>(rowsCount);

            using (FileStream stream = File.Open(model.OpenedFile, FileMode.Open, FileAccess.Read))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            using (PowerPoint.Application powerApplication = new PowerPoint.Application())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
                presentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeA4Paper;

                int pagesCreated = 0;


                for (int row = model.RowStart; row <= model.RowFinish; row++)
                {
                    DownloadedItem item = new DownloadedItem();
                    if (!string.IsNullOrEmpty(model.ColumnText))
                    {
                        var cell = worksheet.Cells[row, ExcelColumnNameToNumber(model.ColumnText)];
                        item.Text = GetCellValue(cell);
                    }
                    if (!string.IsNullOrEmpty(model.ColumnHyperlink))
                    {
                        var cell = worksheet.Cells[row, ExcelColumnNameToNumber(model.ColumnHyperlink)];
                        item.Hyperlink = GetUrlFromCell(cell);
                    }
                    if (!string.IsNullOrEmpty(model.ColumnURI1))
                    {
                        string bitmap = await DownloadBitmap(worksheet.Cells[row, ExcelColumnNameToNumber(model.ColumnURI1)], tempDir);
                        if (bitmap != null)
                            item.Bitmaps.Add(bitmap);
                    }
                    if (!string.IsNullOrEmpty(model.ColumnURI2))
                    {
                        string bitmap = await DownloadBitmap(worksheet.Cells[row, ExcelColumnNameToNumber(model.ColumnURI2)], tempDir);
                        if (bitmap != null)
                            item.Bitmaps.Add(bitmap);
                    }
                    if (!string.IsNullOrEmpty(model.ColumnURI3))
                    {
                        string bitmap = await DownloadBitmap(worksheet.Cells[row, ExcelColumnNameToNumber(model.ColumnURI3)], tempDir);
                        if (bitmap != null)
                            item.Bitmaps.Add(bitmap);
                    }
                    if (!string.IsNullOrEmpty(model.ColumnURI4))
                    {
                        string bitmap = await DownloadBitmap(worksheet.Cells[row, ExcelColumnNameToNumber(model.ColumnURI4)], tempDir);
                        if (bitmap != null)
                            item.Bitmaps.Add(bitmap);
                    }

                    var newSlide = presentation.Slides.Add(pagesCreated + 1, PpSlideLayout.ppLayoutTitleOnly);

                    if (!string.IsNullOrEmpty(item.Text))
                        DrowItemText(newSlide, item.Text);

                    newSlide.FollowMasterBackground = MsoTriState.msoFalse;
                    SetPhotosOnPage(newSlide, item);

                    if (item.Hyperlink != null)
                        DrowHyperlinkText(newSlide, item.Hyperlink);

                    progress.Report(++pagesCreated * 100 / rowsCount);
                }

            }
            if (Directory.Exists(tempDir))
            {
                try
                {
                    Directory.Delete(tempDir, true);
                }
                catch  {}
            }
            return items;
        }   

        private static void DrowItemText(Slide newSlide, string text)
        {
            PowerPoint.Shape label = newSlide.Shapes[1];
            label.Left = 20;
            label.Top = 20;
            label.Width = 750;
            label.Height = 20;
            label.TextFrame.TextRange.Font.Size = 12;
            label.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            label.TextFrame.TextRange.Text = text;
            label.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
        }
        private static void DrowHyperlinkText(Slide newSlide, Uri uri)
        {
            PowerPoint.Shape label = newSlide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 20, 500, 750, 20);
            label.Left = 20;
            label.Width = 750;
            label.TextFrame.TextRange.Font.Size = 12;
            label.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
            label.TextFrame.TextRange.Font.Color.SchemeColor = PpColorSchemeIndex.ppAccent2;// .RGB = System.Drawing.Color.FromArgb(100, 51, 102, 187).ToArgb();
            label.TextFrame.TextRange.Text = uri.ToString();
            label.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            label.ActionSettings[PpMouseActivation.ppMouseClick].Action = PpActionType.ppActionHyperlink;
            label.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.Address = uri.ToString();
        }
        private static void SetPhotosOnPage(PowerPoint.Slide slide, DownloadedItem item)
        {
            //проверяем, сколько фотографий пойдет на слайд
            int photoOnPage = item.Bitmaps.Count;
            if (photoOnPage == 0) //tckb 0, завершаем
                return;

            byte numOfPhoto = 1;

            foreach (var photo in item.Bitmaps)
            {
                if (InsertPhotoToPage(slide, photo, numOfPhoto, photoOnPage))
                    numOfPhoto++;
            }   
        }
        private static bool InsertPhotoToPage(PowerPoint.Slide slide, string path, byte photoNum, int photoOnPage)
        {
            if (!System.IO.File.Exists(path))
                return false;

            var resizeSettings = CreateResizerSettings(photoNum, photoOnPage);
            var newImageSize = ImageResizer.Resizer.CalculateNewSize(path, resizeSettings);
            var newImageLocation = CreateImageLocation(photoNum, photoOnPage);
            try
            {
                slide.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, left: newImageLocation.X, top: newImageLocation.Y, width: newImageSize.Width, height: newImageSize.Height);
                return true;
            }
            catch
            {
                return false;
            }
        }
        private static ResizeSettings CreateResizerSettings(byte photoNum, int photoOnPage)
        {
            ResizeSettings settings = new ResizeSettings { ResizeMode = ResizeMode.Rectangle, ScaleMode = ScaleMode.Both };

            if (photoOnPage == 4)
            {
                settings.Height = 204;
                settings.Width = 272;
            }
            else if (photoOnPage == 3)
            {
                if (photoNum == 1)
                {
                    settings.Height = 300;
                    settings.Width = 400;
                }
                else
                {
                    settings.Height = 204;
                    settings.Width = 272;
                }
            }
            else if (photoOnPage == 2)
            {
                settings.Height = 255;
                settings.Width = 340;
            }
            else
            {
                settings.Height = 346;
                settings.Width = 462;
            }
            return settings;
        }
        private static System.Drawing.Point CreateImageLocation(byte photoNum, int photoOnPage)
        {
            if (photoOnPage == 4)
            {
                if (photoNum == 1) return new System.Drawing.Point(40, 60);
                if (photoNum == 2) return new System.Drawing.Point(470, 60);
                if (photoNum == 3) return new System.Drawing.Point(40, 280);                
                if (photoNum == 4) return new System.Drawing.Point(470, 280);
            }
            else if (photoOnPage == 3)
            {
                if (photoNum == 1) return new System.Drawing.Point(40, 185);
                if (photoNum == 2) return new System.Drawing.Point(470, 60);
                if (photoNum == 3) return new System.Drawing.Point(470, 280);
            }
            else if (photoOnPage == 2)
            {
                if (photoNum == 1) return new System.Drawing.Point(40, 185);
                if (photoNum == 2) return new System.Drawing.Point(400, 185);
            }
            return new System.Drawing.Point(40, 160);
        }

        private static async Task<string> DownloadBitmap(ExcelRange cell, string dir)
        {
            Uri uri = GetUrlFromCell(cell);
            if (uri != null)
            {
                string path = await DownloadBitmap(uri, dir);
                return path;
            }
            return null;            
        }
        private static Task<string> DownloadBitmap(Uri uri, string dir)
        {
            return Task.Run(() => {
                try
                {
                    using (var info = ImageResizer.ImageInfo.Build(uri))
                    {
                        string destinationfile = dir + @"\" + Guid.NewGuid() + "." + info.SourceExtention;
                        info.SaveAs(destinationfile);
                        return destinationfile;
                    }
                }
                catch
                {
                    return null;
                }
            });
        }


        private static string GetCellValue(ExcelRange cell)
        {            
            return cell?.Value?.ToString();
        }

        internal static Uri GetUrlFromCell(ExcelRange cell)
        {
            if (cell == null) return null;

            var hyperlink = cell.Hyperlink;
            if (hyperlink != null)
                return hyperlink;

            if (cell.Value != null)
            {
                string url = cell.Value.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(url) && Uri.TryCreate(url, UriKind.Absolute, out hyperlink))
                    return hyperlink;
            }

            if (cell.Formula != null)
            {
                string formula = cell.Formula;
                if (formula.Contains("HYPERLINK"))
                {
                    int Start, End;
                    Start = formula.IndexOf('"', 0) + 1;
                    End = formula.IndexOf('"', Start);
                    formula = formula.Substring(Start, End - Start);
                    if (Uri.TryCreate(formula, UriKind.Absolute, out hyperlink))
                        return hyperlink;
                }
            }
            return null;           
        }
        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += columnName[i] - 'A' + 1;
            }
            return sum;
        }
    }
}

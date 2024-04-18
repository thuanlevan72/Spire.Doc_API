
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Text;
using WebApplication7.Commons;
using WebApplication7.ViewModel;
using Spire.Barcode;
using System.Drawing;
using System;
using System.Globalization;
using System.Reflection;






namespace WebApplication7.Helpers
{
    public class ConvertWordToPDF
    {
        private readonly Document _doc;
        private readonly IConfigurationBuilder _builder;
        public ConvertWordToPDF()
        {
            _doc = new Document();
            _builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json");
        }
        public string GetListString(List<string> data)
        {
            StringBuilder noiDung = new StringBuilder();
            for (int i = 0; i < data.Count; i++)
            {
                if (i == data.Count - 1)
                {
                    noiDung.Append($"{data[i]}");
                }
                else
                {
                    noiDung.Append($"{data[i]} \n");
                }
            }

            return noiDung.ToString();
        }
        private Image CreateBarcode(string code)
        {
            BarcodeSettings settings = new BarcodeSettings();
            settings.Type = BarCodeType.Code128;
            settings.Data = code;
            settings.ShowTopText = false;
            settings.ShowText = false;
            settings.ImageWidth = 10;
            settings.ImageHeight = 10;

            BarCodeGenerator generator = new BarCodeGenerator(settings);
            Image barcodeImg = generator.GenerateImage();
            return barcodeImg;
        }
        private void HandleImages(TextSelection selection, Image image)
        {
            TextRange range = selection.GetAsOneRange();
            Paragraph paragraph = range.OwnerParagraph;

            int index = paragraph.ChildObjects.IndexOf(range);

            Size size = new Size(200, 20);
            Image processedImage = CropImageToFrame(image, size);

            DocPicture picture = paragraph.AppendPicture(processedImage);
            picture.TextWrappingType = TextWrappingType.Left;
            picture.Height = size.Height;
            picture.Width = size.Width;
            picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            picture.VerticalPosition = -2.0F;

            paragraph.ChildObjects.Insert(index, picture);
            paragraph.ChildObjects.Remove(range);
        }
        public static Image CropImageToFrame(Image img, Size size)
        {
            if (img.Height < size.Height && img.Width < size.Width)
                return img;
            Bitmap res = new Bitmap(size.Width, size.Height);
            using (var g = Graphics.FromImage(res))
            {
                var scale = Math.Max((float)size.Width / img.Width, (float)size.Height / img.Height);
                var scaleWidth = (int)(img.Width * scale);
                var scaleHeight = (int)(img.Height * scale);
                g.DrawImage(img, new Rectangle((size.Width - scaleWidth) / 2, (size.Height - scaleHeight) / 2, scaleWidth, scaleHeight));
            }
            return res;
        }
        public void HandleInsertCheckBox(PhieuKhamBenhViewModel data)
        {
            // find insert check box
            TextSelection[] textSelections = _doc.FindAllString(PhieuKhamBenhKEY.CHECK_THUONG, true, true);
            foreach (TextSelection textSelection in textSelections)
            {
                // Tạo hình ảnh mới từ file
                DocPicture picture = new DocPicture(_doc);
                if (data.CheckKhanCap)
                {
                    picture.LoadImage(Image.FromFile(Constants.UrlImageUnCheck));
                }
                else
                {
                    picture.LoadImage(Image.FromFile(Constants.UrlImageCheck));
                }

                // Đặt kiểu cho hình ảnh

                picture.Width = 10;
                picture.Height = 10;

                // Lấy đoạn văn bản đang chứa textSelection
                Paragraph paragraph = textSelection.GetAsOneRange().OwnerParagraph;

                // Xác định vị trí của đoạn văn bản cần thay thế trong đoạn văn bản
                int index = paragraph.ChildObjects.IndexOf(textSelection.GetAsOneRange());

                // Thêm hình ảnh vào vị trí của đoạn văn bản cần thay thế
                paragraph.ChildObjects.Insert(index, picture);

                // Xóa đoạn văn bản cần thay thế
                paragraph.ChildObjects.Remove(textSelection.GetAsOneRange());
            }
            // find insert check box khẩn cấp
            TextSelection[] textSelectionhelp = _doc.FindAllString(PhieuKhamBenhKEY.CHECK_KHAN_CAP, true, true);
            foreach (TextSelection textSelection in textSelectionhelp)
            {
                // Tạo hình ảnh mới từ file
                DocPicture picture = new DocPicture(_doc);
                if (!data.CheckKhanCap)
                {
                    picture.LoadImage(Image.FromFile(Constants.UrlImageUnCheck));
                }
                else
                {
                    picture.LoadImage(Image.FromFile(Constants.UrlImageCheck));
                }

                // Đặt kiểu cho hình ảnh

                picture.Width = 10;
                picture.Height = 10;


                // Lấy đoạn văn bản đang chứa textSelection
                Paragraph paragraph = textSelection.GetAsOneRange().OwnerParagraph;

                // Xác định vị trí của đoạn văn bản cần thay thế trong đoạn văn bản
                int index = paragraph.ChildObjects.IndexOf(textSelection.GetAsOneRange());

                // Thêm hình ảnh vào vị trí của đoạn văn bản cần thay thế
                paragraph.ChildObjects.Insert(index, picture);

                // Xóa đoạn văn bản cần thay thế
                paragraph.ChildObjects.Remove(textSelection.GetAsOneRange());
            }
        }
        public void ReplaceTextInWord(PhieuKhamBenhViewModel data)
        {
            var configuration = _builder.Build();
            var hashData = configuration.GetSection("HashData");


            string noiDungChuanDoanSoBo = GetListString(data.ChuanDoanSoBo);
            string noiDungCacBoPhan = GetListString(data.KhamLamSan.CacBoPhan);
            string noiDungGhiChu = GetListString(data.GhiChus);
            _doc.Replace(PhieuKhamBenhKEY.SO_DIEN_THOAI_NGUOI_NHA, data.SoDienThoaiNguoiNha, false, true);
            _doc.Replace(PhieuKhamBenhKEY.HAN_BAO_HIEM, data.HanBaoHiem.ToString("dd/MM/yyyy"), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NOI_LAM_VIEC, data.NoiLamViec, false, true);
            _doc.Replace(PhieuKhamBenhKEY.DIA_CHI, data.DiaChi, false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGAY_SINH, data.NgaySinh, false, true);
            _doc.Replace(PhieuKhamBenhKEY.DAN_TOC, data.DanToc, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TEN_NGUOI_BENH, data.TenNguoiBenh, false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGHE_NGHIEP, data.NgheNghiep, false, true);
            _doc.Replace(PhieuKhamBenhKEY.QUOC_TICH, data.QuocTich, false, true);
            _doc.Replace(PhieuKhamBenhKEY.SO_PHIEU, data.SoPhieu, false, true);
            _doc.Replace(PhieuKhamBenhKEY.MA_NB, data.MA_NB, false, true);// hash trước đã
            _doc.Replace(PhieuKhamBenhKEY.QUOC_TICH, data.QuocTich, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TUOI, data.Tuoi.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NOI_DUNG_GHI_CHU, noiDungGhiChu.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.THOI_GIAN_DEN_KHAM, data.ThoiGianDenKham.ToString(Constants.TimeFormat), false, true);
            _doc.Replace(PhieuKhamBenhKEY.THOI_GIAN_BAT_DAU_KHAM, data.ThoiGianBatDauKham.ToString(Constants.TimeFormat), false, true);
            _doc.Replace(PhieuKhamBenhKEY.THOI_GIAN_BAT_DAU_KHAM, noiDungGhiChu.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGAY_TAO, data.NgayTaoDon.ToString(Constants.DateFormat, new CultureInfo("vi-VN")), false, true);
            _doc.Replace(PhieuKhamBenhKEY.KHAM_TOAN_THAN, data.KhamLamSan.ToanThan, false, true);
            _doc.Replace(PhieuKhamBenhKEY.LY_DO_KHAM_BENH, data.LyDoKhanhBenh, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TIEN_SU_BAN_THAN, GetListString(data.TienSuBenh.TienSuBanThan), false, true);
            _doc.Replace(PhieuKhamBenhKEY.TIEN_SU_GIA_DINH, GetListString(data.TienSuBenh.TienSuGiaDinh), false, true);
            _doc.Replace(PhieuKhamBenhKEY.BENH_SU, data.BenhSu, false, true);
            _doc.Replace(PhieuKhamBenhKEY.SO_BAO_HIEM_Y_TE, data.SoBHYT, false, true);
            _doc.Replace(PhieuKhamBenhKEY.CHI_DINH_XET_NGHIEM, data.ChiDinhCanLamSan.XetNghiem, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TDCN, data.ChiDinhCanLamSan.ChuanDoanHinhAnh, false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGUOI_LAP_DON, data.NguoiLapDon, false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGUOI_GIOI_THIEU, data.NguoiGioiThieu, false, true);
            _doc.Replace(PhieuKhamBenhKEY.CHUAN_DOAN_SO_BO, noiDungChuanDoanSoBo, false, true);
            _doc.Replace(PhieuKhamBenhKEY.NOI_DUNG_SU_CHI, data.NoiDungSuLy, false, true);
            _doc.Replace(PhieuKhamBenhKEY.KHAM_BO_PHAN, noiDungCacBoPhan, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TOM_TAC_KET_QUA, data.TomtacKetQua, false, true);
            _doc.Replace(PhieuKhamBenhKEY.THONG_TIN_NGUOI_NHA, data.ThongTinNguoiNha, false, true);
            _doc.Replace(PhieuKhamBenhKEY.GIOI_TINH, data.GioiTinh == true ? "Nữ" : "Nam", false, true);
            _doc.Replace(PhieuKhamBenhKEY.BMI, data.ThongSoSucKhoe.BMI.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.HUYET_AP, data.ThongSoSucKhoe.HuyetAp.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.MANH_DAP, data.ThongSoSucKhoe.ManhDap.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NHIET_DO, data.ThongSoSucKhoe.NhietDo.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.CAN_NANG, data.ThongSoSucKhoe.CanNang.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NHIP_THO, data.ThongSoSucKhoe.NhipTho.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.CHIEU_CAO, data.ThongSoSucKhoe.ChieuCao.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.SP02, data.ThongSoSucKhoe.SPO2.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NOI_DUNG_GHI_CHU, noiDungGhiChu, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TEN_SO_Y_TE, Constants.SoYTE, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TEN_BENH_VIEN, Constants.TenBenhVien, false, true);
        }
        public byte[] ReplaceTextInWord(PhieuKhamBenhViewModel data, string path)
        {
            var configuration = _builder.Build();
            var hashData = configuration.GetSection("HashData");

            _doc.LoadFromFile(path);

            Image barcodeImg = CreateBarcode(data.MA_NB);
            TextSelection[] selections = _doc.FindAllString(PhieuKhamBenhKEY.IMAGE_BARCODE, false, true);
            HandleImages(selections[0], barcodeImg);
            HandleInsertCheckBox(data);
            // Replace Text In Word
            ReplaceTextInWord(data);
            AddTableToDoc(PhieuKhamBenhKEY.TABLE_BENH_CHINH, data.chanDoanXacDinh);
            //CreateTableFromHTMLs(PhieuKhamBenhKEY.TABLE_BENH_CHINH, data.chanDoanXacDinh);


            string outputFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Temp", "Output.docx");
            _doc.SaveToFile(outputFilePath, FileFormat.Docx2016);
            using (MemoryStream stream = new MemoryStream())
            {
                _doc.SaveToStream(stream, FileFormat.Docx2016);
                byte[] result = stream.ToArray();

                _doc.Close();

                return result;
            }
        }
        #region chuyển file DOCX SANG Pdf
        public (MemoryStream, string) ConvertDocumentToPdf(byte[] fileData)
        {
            // Create a MemoryStream from the fileData
            MemoryStream ms = new MemoryStream(fileData);
            _doc.LoadFromStream(ms, FileFormat.Docx2016);
            MemoryStream stream = new MemoryStream();
            _doc.SaveToStream(stream, FileFormat.PDF);
            stream.Position = 0; // Reset the stream position to the beginning.

            _doc.Dispose();

            string fileName = @"Document-" + DateTime.Now.ToString("dd-MM-yyyy") + "-" + Guid.NewGuid().ToString() + @"-Converted.pdf";

            //string directoryPath = @"C:\Users\thuan\source\repos\WebApplication7\WebApplication7\PDFResult\";
            string fullPath = Path.Combine(Constants.PDFResultPath, fileName);

            // Lưu file vào thư mục
            File.WriteAllBytes(fullPath, stream.ToArray());
            var memory = new MemoryStream();
            using (var streams = new FileStream(fullPath, FileMode.Open))
            {
                streams.CopyTo(memory);
            }
            memory.Position = 0;
            return (memory, fileName);
        }
        #endregion
        private void AddTableToDoc(string placeholder, ChanDoanXacDinh inputData)
        {
            String[] Header = { $"- Bệnh chính: {inputData.BenhChinh.TenBenh}", $"Mã ICD: {inputData.BenhChinh.ICD_CODE}" };

            List<String[]> tempList = new List<string[]>()
            {
                new String[] { "- Bệnh kèm theo: ", null }
            };
            for (int i = 0; i < inputData.benhPhu.Count; i++)
            {
                tempList.Add(new String[] { $"{inputData.benhPhu[i].TenBenh}", $"Mã ICD: {inputData.benhPhu[i].ICD_CODE}" });
            }
            String[][] data = tempList.ToArray();

            TextSelection selection = _doc.FindString(placeholder, false, true);
            TextRange range = selection.GetAsOneRange();

            Table table = range.OwnerParagraph.OwnerTextBody.AddTable(true);

            table.ResetCells(data.Length + 1, Header.Length);

            TableRow FRow = table.Rows[0];

            FRow.Height = 20;

            FRow.Cells[0].Width = (float)(table.Width * 0.80);
            FRow.Cells[1].Width = (float)(table.Width * 0.20);

            for (int i = 0; i < Header.Length; i++)
            {
                Paragraph p = FRow.Cells[i].AddParagraph();
                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Top;
                p.Format.HorizontalAlignment = HorizontalAlignment.Left;

                TextRange TR = p.AppendText(Header[i]);
                TR.CharacterFormat.FontName = "Times New Roman";
                TR.CharacterFormat.FontSize = 12;

            }
            table.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.None;
            for (int r = 0; r < data.Length; r++)
            {
                TableRow DataRow = table.Rows[r + 1];
                DataRow.Height = 15;


                DataRow.Cells[0].Width = (float)(table.Width * 0.80);
                DataRow.Cells[1].Width = (float)(table.Width * 0.20);

                for (int c = 0; c < data[r].Length; c++)
                {
                    table.Rows[r].Cells[c].CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.None;
                    DataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    Paragraph p2 = DataRow.Cells[c].AddParagraph();

                    if (data[r][c] != null)
                    {
                        TextRange TR2 = p2.AppendText(data[r][c]);
                        p2.Format.HorizontalAlignment = HorizontalAlignment.Left;


                        TR2.CharacterFormat.FontName = "Times New Roman";
                        TR2.CharacterFormat.FontSize = 12;

                    }
                }
            }
            Paragraph paragraph = range.OwnerParagraph;

            Paragraph newParagraph = new Paragraph(_doc);
            

            //// Add the new paragraph at the position of the placeholder
            //paragraph.OwnerTextBody.ChildObjects.Insert(0, newParagraph);

            // Now remove the paragraph containing the placeholder
            paragraph.OwnerTextBody.ChildObjects.Remove(paragraph);

        }
        public void CreateTableFromHTMLs(string placeholder, ChanDoanXacDinh inputData)
        {
            string benhPhuHTML = "";
            for (int i = 0; i < inputData.benhPhu.Count; i++)
            {
                benhPhuHTML += $@"<tr>
            <td style=""width: 70%"">{inputData.benhPhu[i].TenBenh}</td>
            <td style=""width: 30%"">Mã ICD: {inputData.benhPhu[i].ICD_CODE}</td>
            </tr>";
            }
            string htmlString = $@"
            <table style=""width: 100%;  font-size: 16px; "">

            <tbody>
            <tr>
            <td style=""width: 70%"">- Bệnh chính: {inputData.BenhChinh.TenBenh}</td>
            <td style=""width: 30%"">Mã ICD: {inputData.BenhChinh.ICD_CODE}</td>
            </tr>
            <tr>
            <td style=""width: 70%"">- Bệnh kèm theo: </td>
            <td style=""width: 30%""></td>
            </tr>
            {benhPhuHTML}
            </tbody>
            </table>";

            //Create a Word document


            //Add a section
            TextSelection selection = _doc.FindString(placeholder, false, true);
            TextRange range = selection.GetAsOneRange();

            //Add a paragraph and append html string
            range.Text = "";
            range.OwnerParagraph.AppendHTML(htmlString);
            // Apply the style to each TextRange in the paragraph
            for (int i = 0; i < range.OwnerParagraph.ChildObjects.Count; i++)
            {
                if (range.OwnerParagraph.ChildObjects[i] is TextRange)
                {
                    TextRange text = (TextRange)range.OwnerParagraph.ChildObjects[i];
                    text.CharacterFormat.Bold = false;
                }
            }
            //Paragraph paragraph = range.OwnerParagraph;
            //int index = paragraph.OwnerTextBody.ChildObjects.IndexOf(paragraph);
            //if (index >= 0)
            //{
            //    paragraph.OwnerTextBody.ChildObjects.Insert(index, table);
            //}

            range.OwnerParagraph.ChildObjects.Remove(range);



        }
    }


}

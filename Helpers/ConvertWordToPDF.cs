using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Text;
using WebApplication7.Commons;
using WebApplication7.ViewModel;
using Spire.Barcode;
using System.Drawing;
using System;



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
        private string GetChuanDoanSoBo(List<string> chuanDoanSoBos)
        {
            StringBuilder noiDungChuanDoanSoBo = new StringBuilder();
            for (int i = 0; i < chuanDoanSoBos.Count; i++)
            {
                if (i == chuanDoanSoBos.Count - 1)
                {
                    noiDungChuanDoanSoBo.Append($"- {chuanDoanSoBos[i]}");
                }
                else
                {
                    noiDungChuanDoanSoBo.Append($"- {chuanDoanSoBos[i]} \n");
                }
            }
          
            return noiDungChuanDoanSoBo.ToString();
        }
        public string GetListString(List<string> data)
        {
            StringBuilder noiDung = new StringBuilder();
            for(int i = 0; i< data.Count; i++)
            {
               if(i == data.Count - 1)
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

        public byte[] ReplaceTextInWord(PhieuKhamBenhViewModel data, string path)
        {
            var configuration = _builder.Build();
            var hashData = configuration.GetSection("HashData");

            StringBuilder noiDungGhiChu = new StringBuilder();
       
            data.GhiChus.ForEach(title =>
            {
                noiDungGhiChu.Append($" {title} \n");
             
            });
            string noiDungChuanDoanSoBo = GetListString(data.ChuanDoanSoBo);
            string noiDungCacBoPhan = GetListString(data.KhamLamSan.CacBoPhan);
            _doc.LoadFromFile(path);
            BarcodeSettings settings = new BarcodeSettings();
            settings.Type = BarCodeType.Code128;
            settings.Data = Guid.NewGuid().ToString();
            BarCodeGenerator generator = new BarCodeGenerator(settings);
            Image barcodeImg = generator.GenerateImage();

            // đang hơi chưa biết nên đã hash code
            //_doc.Replace(PhieuKhamBenhKEY.TEN_SO_Y_TE, hashData["SoYTE"], false, true);
            //TextSelection[] selections = _doc.FindAllString("${barcodeImg}", false, true);
            //foreach (TextSelection local in selections)
            //{
            //    TextRange rangeIndex = local.GetAsOneRange();
            //    Paragraph para = rangeIndex.OwnerParagraph;
            //    int index = para.ChildObjects.IndexOf(rangeIndex);
            //    DocPicture picture = para.AppendPicture(barcodeImg);
         
            //    picture.Width = 50; // Sửa độ rộng ảnh
            //    picture.Height = 30; // Sửa độ cao ảnh
            //    picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            //    para.ChildObjects.Insert(index, picture);
            //    para.ChildObjects.Remove(rangeIndex);
            //}
            _doc.Replace(PhieuKhamBenhKEY.NOI_LAM_VIEC, data.NoiLamViec, false, true);
            _doc.Replace(PhieuKhamBenhKEY.DIA_CHI,data.DiaChi , false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGAY_SINH, data.NgaySinh, false, true);
            _doc.Replace(PhieuKhamBenhKEY.DAN_TOC, data.DanToc, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TEN_NGUOI_BENH, data.TenNguoiBenh, false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGHE_NGHIEP, data.NgheNghiep, false, true);
            _doc.Replace(PhieuKhamBenhKEY.QUOC_TICH, data.QuocTich, false, true);
            _doc.Replace(PhieuKhamBenhKEY.SO_PHIEU, data.SoPhieu, false, true);
            _doc.Replace(PhieuKhamBenhKEY.MA_NB, data.MA_NB, false, true);// hash trước đã
            _doc.Replace(PhieuKhamBenhKEY.QUOC_TICH, data.QuocTich, false, true);
            _doc.Replace(PhieuKhamBenhKEY.TUOI, data.Tuoi.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NOI_DUNG_GHI_CHU,noiDungGhiChu.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.THOI_GIAN_DEN_KHAM, data.ThoiGianDenKham.ToString("'ngày 'dd' tháng 'MM' năm 'yyyy', 'HH' giờ 'mm' phút'"), false, true);
            _doc.Replace(PhieuKhamBenhKEY.THOI_GIAN_DEN_KHAM, data.ThoiGianBatDauKham.ToString("'ngày 'dd' tháng 'MM' năm 'yyyy', 'HH' giờ 'mm' phút'"), false, true);
            _doc.Replace(PhieuKhamBenhKEY.THOI_GIAN_BAT_DAU_KHAM, noiDungGhiChu.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NGAY_TAO, data.NgayTaoDon.ToString("'dd' tháng 'MM' năm 'yyyy'"), false, true);
            _doc.Replace(PhieuKhamBenhKEY.KHAM_TOAN_THAN, data.KhamLamSan.ToanThan, false, true);
            _doc.Replace(PhieuKhamBenhKEY.LY_DO_KHAM_BENH, data.LyDoDenKham, false, true);
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
            _doc.Replace(PhieuKhamBenhKEY.GIOI_TINH, data.GioiTinh == true ? "Nữ" : "Nam", false, true);
            //
            _doc.Replace(PhieuKhamBenhKEY.BMI, data.ThongSoSucKhoe.BMI.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.HUYET_AP, data.ThongSoSucKhoe.HuyetAp.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.MANH_DAP, data.ThongSoSucKhoe.ManhDap.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NHIET_DO, data.ThongSoSucKhoe.NhietDo.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.CAN_NANG, data.ThongSoSucKhoe.CanNang.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.NHIP_THO, data.ThongSoSucKhoe.NhipTho.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.CHIEU_CAO, data.ThongSoSucKhoe.ChieuCao.ToString(), false, true);
            _doc.Replace(PhieuKhamBenhKEY.SP02, data.ThongSoSucKhoe.SPO2.ToString(), false, true);
            //
            TextSelection selection = _doc.FindString(PhieuKhamBenhKEY.TEN_BENH_VIEN, false, true);
            TextRange range = selection.GetAsOneRange();

            range.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
            range.Text = hashData["TenBenhVien"];

            _doc.Replace("${NOIDUNGGHICHU}", noiDungGhiChu.ToString(), false, true);
            //_doc.Replace("${BENH_KEM_THEO}", noiDungGhiChu.ToString(), false, true);
            //AddTable();
            AddTableToDoc("${TableBenhChinh}",data.chanDoanXacDinh);

            MemoryStream stream = new MemoryStream();
            _doc.SaveToStream(stream, FileFormat.Docx2016);
            byte[] result = stream.ToArray();

            _doc.Close();

            return result;
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
            string directoryPath = @"C:\Users\Admin\Source\Repos\Spire.Doc_API\PDFResult\"; // Thay thế bằng đường dẫn thư mục mà bạn muốn lưu file vào
            string fullPath = Path.Combine(directoryPath, fileName);

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
            String[] Header = { $"Bệnh chính: {inputData.BenhChinh.TenBenh}", $"Mã ICD: {inputData.BenhChinh.ICD_CODE}" };
            //String[][] data = {
            //          new String[]{ "Bệnh kèm theo: ", null},
            //          new String[]{ "Aceton niệu Bệnh kèm theo", "Mã ICD: R82.4"},
            //          new String[]{ "Acidoniu Bệnh kèm theo", "Mã ICD: R82.4"},
            //      };

            List<String[]> tempList = new List<string[]>()
            {
                new String[] { "Bệnh kèm theo: ", null }
            };
            for (int i = 0;i< inputData.benhPhu.Count; i++)
            {
                tempList.Add(new String[] { $"{inputData.benhPhu[i].TenBenh}", $"Mã ICD: {inputData.benhPhu[i].ICD_CODE}" });
            }
            String[][] data = tempList.ToArray();

            TextSelection selection = _doc.FindString(placeholder, false, true);
            TextRange range = selection.GetAsOneRange();

            Table table = range.OwnerParagraph.OwnerTextBody.AddTable(true);

            table.ResetCells(data.Length + 1, Header.Length);

            TableRow FRow = table.Rows[0];

            FRow.Height = 23;
 

            // Set the width for the cells
            FRow.Cells[0].Width = (float)(table.Width * 0.7);
            FRow.Cells[1].Width = (float)(table.Width * 0.3);

            for (int i = 0; i < Header.Length; i++)
            {
                Paragraph p = FRow.Cells[i].AddParagraph();
                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Left;

                TextRange TR = p.AppendText(Header[i]);
                TR.CharacterFormat.FontName = "Times New Roman";
                TR.CharacterFormat.FontSize = 12;
              
            }
            table.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.None;
            for (int r = 0; r < data.Length; r++)
            {
                TableRow DataRow = table.Rows[r + 1];
                DataRow.Height = 20;

                // Set the width for the cells
                DataRow.Cells[0].Width = (float)(table.Width * 0.7);
                DataRow.Cells[1].Width = (float)(table.Width * 0.3);

                for (int c = 0; c < data[r].Length; c++)
                {
                    table.Rows[r].Cells[c].CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.None;
                    DataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    Paragraph p2 = DataRow.Cells[c].AddParagraph();
                    // check if the cell data is not null before appending it
                    if (data[r][c] != null)
                    {
                        TextRange TR2 = p2.AppendText(data[r][c]);
                        p2.Format.HorizontalAlignment = HorizontalAlignment.Left;

                        //Set data format
                        TR2.CharacterFormat.FontName = "Times New Roman";
                        TR2.CharacterFormat.FontSize = 12;

                    }
                }
            }

            range.OwnerParagraph.ChildObjects.Remove(range);
        }
    }
}

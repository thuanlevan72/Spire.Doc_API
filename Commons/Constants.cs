namespace WebApplication7.Commons
{
    public class Constants
    {
        public const string DateFormat = "'Ngày' dd 'tháng' MM 'năm' yyyy";
        public const string TimeFormat = "HH 'giờ' mm 'phút', 'ngày' dd 'tháng' MM 'năm' yyyy";
        public static readonly string UrlImageCheck = Path.Combine(Directory.GetCurrentDirectory(), "Images", "check.png");
        public static readonly string UrlImageUnCheck = Path.Combine(Directory.GetCurrentDirectory(), "Images", "square_11067333.png");
        public static readonly string PDFResultPath = Path.Combine(Directory.GetCurrentDirectory(), "PDFResult");
        public static readonly string PDFPullPath = Path.Combine(Directory.GetCurrentDirectory(), "DocxTemplates", "Phieu-kham-vao-vien-v2_20240416.docx");
        public static readonly string PDFPullTestPath = Path.Combine(Directory.GetCurrentDirectory(), "Temp", "Phieu-kham-vao-vien-v2_cover.docx");

        public const string TenBenhVien = "BỆNH VIỆN ĐẠI HỌC PHENIKAA";
        public const string SoYTE = "SỞ Y TẾ THÀNH PHỐ HÀ NỘI";
    }
}

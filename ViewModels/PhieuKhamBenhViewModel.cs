using Spire.Additions.Xps.Schema;
using System.Security.Cryptography.X509Certificates;

namespace WebApplication7.ViewModel
{
    public class PhieuKhamBenhViewModel
    {
        private List<string> _ghiChu;

        public string TenNguoiBenh { get; set; }

        public string NgaySinh { get; set; }

        public string ThongTinNguoiNha {  get; set; }

        public int Tuoi { get; set; }

        public DateTime ThoiGianBatDauKham {  get; set; }= DateTime.Now;

        public DateTime ThoiGianDenKham {  get; set; }= DateTime.Now;

        public string NguoiGioiThieu { get; set; }

        public string TomtacKetQua {  get; set; }                                                                   

        public string QuocTich { get; set; }
        public string DanToc { get; set; }

        public string SoPhieu { get; set; }
        public string MA_NB { get; set; }

        public string MaNB { get; set; }

        public DateTime NgayTaoDon { get; set; }

        public string NgheNghiep { get; set; }

        public string DiaChi { get; set; }

        public string NoiLamViec { get; set; }

        public string SoBHYT { get; set; }

        public string DienThoaiNguoiBenh { get; set; }

        public string TomTacKetQuaLamSan { get; set; }
        public List<string> ChanDoanSoBo { get; set; }
        public ChiDinhCanLamSan ChiDinhCanLamSan { get; set; }

        public KhamLamSan KhamLamSan { get; set; }

        public string TenBacSiKhamBenh { get; set; }

        public string NoiDungSuLy { get; set; }

        public bool GioiTinh {  get; set; }

        public string LyDoKhanhBenh { get; set; }

        public string NguoiLapDon { get; set; }

        public List<string> ChuanDoanSoBo { get; set; }

        public ThongSoSucKhoe ThongSoSucKhoe { get; set; }

        public TienSuBenh TienSuBenh { get; set; }

        public string BenhSu { get; set; }

        public string LyDoDenKham { get; set; }
        public ChanDoanXacDinh chanDoanXacDinh { get; set; }

        // get set 
        public List<string> GhiChus
        {
            get { return _ghiChu; }
            set
            {
                _ghiChu = value;
            }
        }

    }
    public class TienSuBenh
    {
        public List<string> TienSuBanThan { get; set; }

        public List<string> TienSuGiaDinh { get; set; }
    }

    
    public class ChiDinhCanLamSan
    {
        public string XetNghiem { get; set; }

        public string ChuanDoanHinhAnh { get; set; }
    }
    public class KhamLamSan
    {
        public string ToanThan { get; set; }
        public List<string> CacBoPhan { get; set; }
    }

    public class ChanDoanXacDinh
    {
        public Benh BenhChinh { get; set; }
        public List<Benh>  benhPhu {get;set;}
    }
    public class Benh
    {
        public string TenBenh { get; set; }
        public string ICD_CODE { get; set; }
    }
    public class ThongSoSucKhoe
    {
        public int ManhDap { get; set; }
        public float NhietDo { get; set; }
        public float HuyetAp { get; set; }

        public int NhipTho { get; set; }

        public float CanNang { get; set; } 

        public float ChieuCao { get; set; }

        public float BMI { get; set; }

        public float SPO2 { get; set; }
    }
   
}

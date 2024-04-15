namespace WebApplication7.ViewModel
{
    public class PhieuKhamBenhViewModel
    {
        // private valiabo
        private string _tenSoYTE;
        private string _tenBenhVien;
        private List<string> _ghiChu;


        // get set 
        public List<string> GhiChu
        {
            get { return _ghiChu; }
            set
            {
                _ghiChu = value;
            }
        }
        public string TenBenhVien
        {
            get { return _tenBenhVien; }
            set
            {
                _tenBenhVien = value.ToUpper();
            }
        }
        public string TenSoYTE {
            get { return _tenSoYTE; }
            set
            {
                _tenSoYTE = value.ToUpper();
            } }
    }
}

using Microsoft.AspNetCore.Mvc;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;
using System.Text;
using WebApplication7.ViewModel;

namespace WebApplication7.Helpers
{
    public class ConvertWordToPDF
    {
        private readonly Document _doc;
        public ConvertWordToPDF()
        {
            _doc = new Document();
        }
        public byte[] ReplaceTextInWord(PhieuKhamBenhViewModel data, string path)
        {
            StringBuilder noiDungGhiChu = new StringBuilder();
            data.GhiChu.ForEach(title =>
            {
                noiDungGhiChu.Append($"- {title} \n");
            });
            _doc.LoadFromFile(path);
            // đang hơi chưa biết nên đã hash code
            _doc.Replace("${TEN_SO_Y_TE}", data.TenSoYTE, false, true);
            _doc.Replace("${TEN_BENH_VIEN}", data.TenBenhVien, false, true);
        
            _doc.Replace("${NOI_DUNG_GHI_CHU}", noiDungGhiChu.ToString(), false, true);

            AddTableToDoc("${TABLE_PLACEHOLDER}");

            MemoryStream stream = new MemoryStream();
            _doc.SaveToStream(stream, FileFormat.Docx2016);
            byte[] result = stream.ToArray();

            _doc.Close();

            return result;
        }
        public (byte[], string) ConvertDocumentToPdf(byte[] fileData)
        {
            // Create a MemoryStream from the fileData
            MemoryStream ms = new MemoryStream(fileData);
            _doc.LoadFromStream(ms, FileFormat.Docx2016);
            MemoryStream stream = new MemoryStream();
            _doc.SaveToStream(stream, FileFormat.PDF);
            stream.Position = 0; // Reset the stream position to the beginning.

            _doc.Dispose();

            string fileName = "ConvertedDocument.pdf";
            string directoryPath = @"C:\Users\thuan\source\repos\WebApplication7\WebApplication7\PDFResult\"; // Thay thế bằng đường dẫn thư mục mà bạn muốn lưu file vào
            string fullPath = Path.Combine(directoryPath, fileName);

            // Lưu file vào thư mục
            File.WriteAllBytes(fullPath, stream.ToArray());

            return (stream.ToArray(), fileName);
        }
        private void AddTableToDoc( string placeholder)
        {
            String[] Header = { "Date", "Description", "Country", "On Hands", "On Order" };
            String[][] data = {
                                  new String[]{ "08/07/2021","Dive kayak","United States","24","16"},
                                  new String[]{ "08/07/2021","Underwater Diver Vehicle","United States","5","3"},
                                  new String[]{ "08/07/2021","Regulator System","Czech Republic","165","216"},
                                  new String[]{ "08/08/2021","Second Stage Regulator","United States","98","88"},
                                  new String[]{ "08/08/2021","Personal Dive Sonar","United States","46","45"},
                                  new String[]{ "08/09/2021","Compass Console Mount","United States","211","300"},
                                  new String[]{ "08/09/2021","Regulator System","United Kingdom","166","100"},
                                  new String[]{ "08/10/2021","Alternate Inflation Regulator","United Kingdom","47","43"},
                              };
            TextSelection selection = _doc.FindString(placeholder, false, true);
            TextRange range = selection.GetAsOneRange();

            Table table = range.OwnerParagraph.OwnerTextBody.AddTable(true);
            //Add a table
        
            table.ResetCells(data.Length + 1, Header.Length);

            //Set the first row as table header
            TableRow FRow = table.Rows[0];
            FRow.IsHeader = true;

            //Set the height and color of the first row
            FRow.Height = 23;
            FRow.RowFormat.BackColor = Color.LightSeaGreen;
            for (int i = 0; i < Header.Length; i++)
            {
                //Set alignment for cells 
                Paragraph p = FRow.Cells[i].AddParagraph();
                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //Set data format
                TextRange TR = p.AppendText(Header[i]);
                TR.CharacterFormat.FontName = "Calibri";
                TR.CharacterFormat.FontSize = 12;
                TR.CharacterFormat.Bold = true;
            }

            //Add data to the rest of rows and set cell format
            for (int r = 0; r < data.Length; r++)
            {
                TableRow DataRow = table.Rows[r + 1];
                DataRow.Height = 20;
                for (int c = 0; c < data[r].Length; c++)
                {
                    DataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    Paragraph p2 = DataRow.Cells[c].AddParagraph();
                    TextRange TR2 = p2.AppendText(data[r][c]);
                    p2.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    //Set data format
                    TR2.CharacterFormat.FontName = "Calibri";
                    TR2.CharacterFormat.FontSize = 11;
                }
            }

            range.OwnerParagraph.ChildObjects.Remove(range);
        }
    }
}

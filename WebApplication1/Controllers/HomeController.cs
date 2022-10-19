using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            XWPFDocument doc = new XWPFDocument();



     
            XWPFParagraph para = doc.CreateParagraph();
            XWPFRun run = para.CreateRun();
            //=======================
            run = para.CreateRun();
            run.FontSize = 10;   
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Hôm nay, vào hồi");        

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetText("{{Thời gian giao xe thực tế (giờ/ngày/tháng/năm)}}");
            run.SetFontFamily("Arial", FontCharRange.None);
            run.IsBold = true;
            run.IsItalic = true;

            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("A. Tình trạng xe khi nhận lại xe" ); 
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.IsBold = true;

            run = para.CreateRun();
            run.FontSize = 11;   
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("(đánh dấu tích để lựa chọn)");
            run.IsItalic = true;
            //=======================        
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("\t-Tình trạng nội thất, ngoại thất và máy móc xe:");
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.FontSize = 10;
            run.IsBold=true;
            run.IsItalic = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t-Giống tình trạng bạn đầu (nội thất, ngoại thất, máy móc, giấy tờ, đồ dự phòng)");

            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t- Khác tình trạng ban đầu, các hư hỏng và mất mát sau:");

            para = doc.CreateParagraph();
            run = para.CreateRun();    
            run.SetText("\t\t\t........................................................................................................");
     
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("\t\t\tChi phí khắc phục (tạm tính).......................đ");
            run.IsBold = true;
            run = para.CreateRun();
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman",FontCharRange.None);
            run.SetText("\t-Số công tơ mét (Km):");
            run.IsItalic = true;
            run.IsBold = true;
            run = para.CreateRun();
            
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Số km nhận xe thực tế}} ");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\tTổng số km đã đi:");
            run = para.CreateRun();
            
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Số km thực tế hành trình}} ");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t-Nằm trong giới hạn km:");

            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t-Vượt số giới hạn km, số km vượt: ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Số km phụ trội }} ");
            run.IsBold = true;

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("Số tiền phụ trội km: ");

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Phí phụ trội km}}");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\tĐồng hồ xăng/dầu: ");
            run.IsBold = true;
           
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("(vạch xăng):");
            run.IsItalic = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\tPhụ phí xăng dầu ");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t-Thời gian phụ trội so với hợp đồng: ");
            run.IsItalic = true;
            run.IsBold = true;
            run = para.CreateRun();

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Thời gian trả xe quá }}");
            run.IsBold = true;

            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("giờ,");

            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("phụ phí phát sinh ");
            run.IsBold = true;
                                      
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Phí giao muộn}}");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t-Vé cầu đường phát sinh chưa thanh toán: ");          
            run.IsBold = true;
            run = para.CreateRun();

            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Phí phụ trội ETC}}");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t-Các lỗi ghi nhận được trong quá trình thuê:");
            run.IsBold = true;
            run = para.CreateRun();
            run.IsItalic = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t-Chưa phát hiện lỗi gì");
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t-Phát hiện lỗi:");
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t\t-Vượt tốc độ:");
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t\t-Vào đường cấm: ...........................................................................");
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t\t-Vượt đèn đỏ: ...........................................................................");
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.FontSize = 10;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText("\t\t\t\t-Các lỗi khác (nếu có):");
            //=======================
            run = para.CreateRun();
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.SetText(".....................");
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("Tổng chi phí phát sinh so với hợp đồng: ");
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.IsBold = true;
            
            run = para.CreateRun();
            run.SetFontFamily("Arial", FontCharRange.None);
            run.SetText("{{Tổng phát sinh thêm}}");
            run.FontSize = 10;
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("Bên A đã hoàn trả cho bên B một số giấy tờ và tài sản như sau:");
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("\t\t\t-Toàn bộ giấy tờ và tài sản tại thời điểm giao nhận");
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("\t\t\t-Thiếu hoặc chưa hoàn trả các giấy tờ và tài sản sau: ");
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("\t\t\t.....................................................................................................");

            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("\t\t\t.....................................................................................................");
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("Cam kết của Khách thuê và Chủ xe:");
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.IsBold = true;
            //=======================
            para = doc.CreateParagraph();
            run = para.CreateRun();
            run.SetText("Trong trường hợp phát sinh các khoản phạt nguội, bằng chứng do camera giám sát của Cục CSGT - Bộ Công An ghi nhận được trong thời gian Bên B sử dụng xe ô tô thuê của Bên A. Bên A có trách nhiệm cung cấp các bằng chứng liên quan cho bên B ngay khi nhận được thông tin. Bên B cam kết chịu hoàn toàn trách nhiệm và bồi thường toàn bộ các chi phí liên quan cho Bên A");
            run.FontSize = 11;
            run.SetFontFamily("Times New Roman", FontCharRange.None);
            run.IsItalic = true;
            //=======================
            #region
            XWPFTable table = doc.CreateTable(1, 4);
            table.Width = 4000;
            table.GetRow(0).GetCell(0).SetText("Bên A");
            table.GetRow(0).GetCell(3).SetText("Bên B");
            //=======================
            #endregion
            //GHI FILE
            MemoryStream ms = new MemoryStream();
            doc.Write(ms);
            var bytes = ms.ToArray();
            ms.Close();
            return File(bytes, "application/msword", "st.docx");
        }
    }
}
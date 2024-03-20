using System;
//Microsoft Excel 16 object in references-> COM tab
using Excel = Microsoft.Office.Interop.Excel;

namespace How_To_ReadOrWrite_ExcelFile_csharp_2024_Gacon_ruduesMSK_
{
    internal class Program
    {
        /*
                       ┌ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─    
                                                                                                                                │   
                       │                                                                                                            
                                                               Sơ Lượt những đều cần lưu ý !                                    │   
                       │                                                                                                         
                                                                                                                                │   
                       │                                                                                                            
                                      đây là phiên bản thứ nhất, tus mới tìm hiểu về chúng cách đây vài phút !                  │   
                       │                                                                                                            
                                      nên nếu có sự cố ngoài ý muốn làm project phát sinh lỗi                                   │   
                       │                                                                                                            
                                      hoặc tôi hiểu sai 1 số thành phần hoặc đoạn code nào đó xin các bạn thông cảm !           │   
                       │                                                                                                            
                                      phiên bản sau sẽ có phần write ! nếu tôi có thời gian làm nó :))                          │   
                       │                                                                                                            
                                                                                                                                │                                                                                                                                   │   
                       │                                                                                                            
                                                                                                                                │   
                       │                                                                                                            
                                                                                                                                │   
                       │                                                                   - GaCon(2024)🐔 - RuduesMSK -            
                                                                                                                                │   
                       │                                                                                                            
                        ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ┘   
         */
        static void Main(string[] args)
        {
            // trước khi code! hãy thực hiện các bước sau:
            // 1. hãy cài thư viện của Microsoft
            //    visual studio(2022): 1.1 vào Project
            //                         1.2 chọn Manage Nuget package...
            //                         1.3 Install: "Microsoft.Office.Interop.Excel" 


            // todo: change your path
            const string FileName = @"D:\Desktop\demoExcel_ReadOrWrite.xlsx";

            // Read excel with Microsoft.Office.Interop.Word
            Read(FileName);
            Console.ReadKey();
        }

        public static void Read(string FileName)
        {
            //
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();

            // open excel:
            /* ╔═════════════════════════════════════════════ Giải Thích Các Thành Phần ════════════════════════════════════════════╗
               ║                                                                                                                    ║
               ║  ┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────┐  ║
               ║  │                                                                                                              │  ║
               ║  │   + mình chỉ chú thích một số thành phần(tham số) ở dưới ! vì chúng quá nhiều nên mình chỉ nêu lên 1 vài     │  ║
               ║  │     đối tượng cụ thể bạn có thể chi tiết những thứ khác tại trang này:                                       │  ║
               ║  │                                                                                                              │  ║
               ║  │                                                                                                              │  ║
               ║  │                                                                                                              │  ║
               ║  │  https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.workbooks.open?view=excel-pia   │  ║
               ║  │                                                                                                              │  ║
               ║  │                                                                                                              │  ║
               ║  │                                                                                                              │  ║
               ║  └──────────────────────────────────────────────────────────────────────────────────────────────────────────────┘  ║
               ╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝
            */
            Excel.Workbook workBook = excel.Application.Workbooks.Open(
                FileName, // File name
                missing,  
                true,     // ReadOnly
                missing,  
                missing,  // Password
                missing,
                missing,
                missing,
                missing,
                true,     // Edit Table
                missing,
                missing,
                missing,
                missing,
                missing
                );

            // sheet 1:
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Sheets[1];

            // location:
            Excel.Range cell1 = (Excel.Range)worksheet.Cells[1, 1];
            Excel.Range cell2 = (Excel.Range)worksheet.Cells[2, 2];
            Excel.Range cell3 = (Excel.Range)worksheet.Cells[3, 3];
            Excel.Range cell4 = (Excel.Range)worksheet.Cells[4, 4];
            Excel.Range cell5 = (Excel.Range)worksheet.Cells[5, 5];
            
            // show values:
            Console.WriteLine(
                cell1.Value+" " +
                cell2.Value+" " +
                cell3.Value+" " +
                cell4.Value+" " +
                cell5.Value);

            // close excel
            excel.Application.Workbooks.Close();
        }


    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelWithNPOI
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = string.Empty;
            string optionName = string.Empty;
            string inputValue = string.Empty;
            string outputValue = string.Empty;
            string folderPath = "C:\\TEST\\PCLog\\" + DateTime.Now.ToString("yyyy") + "\\" + DateTime.Now.ToString("MM");
            DirectoryInfo di = new DirectoryInfo(folderPath);
            //argument 지정방법 : project -> properties-> debug tab -> fill the text command linn arguments

            if (args.Length > 1)
            {
                try
                {
                    fileName = args[0] + ".xls";
                    optionName = args[1];
                    inputValue = args[2];
                    outputValue = args[3];

                    //fileName = "newFile.xls";
                    //optionName = "12";
                    //inputValue = "34";
                    //outputValue = "56";

                   
                    //테스트 용
                    Console.WriteLine("Parameter : {0}          {1}          {2}          {3}", args[0], args[1], args[2], args[3]);

                    //디렉토리 생성
                    if (di.Exists == false)
                    {
                        di.Create();
                    }
                    else
                    {
                        Console.WriteLine("Failed make directory.");
                    }
                    //저장할 엑셀 파일 경로 설정 
                    FileInfo fileInfo = new FileInfo(folderPath + "\\" + fileName);
                    
                    if(fileInfo.Exists == false)
                    {//파일이 없다면

                        //파일 생성 
                        createExcelFile(folderPath, fileName);
                        //insert data
                        insertExcelData(folderPath, fileName, args);

                    }
                    else
                    {//파일이 있다면
                        insertExcelData(folderPath, fileName, args);
                    }
                                                                                                                       
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : {0}", ex.ToString());
                }
            }
            else
            {
                Console.WriteLine("less Parameter then threw"); 
            }

            
        }

        public static void createExcelFile(string folderPath, string fileName)
        {
            try
            {
                HSSFWorkbook wb = new HSSFWorkbook();
                HSSFSheet sheet = (HSSFSheet)wb.CreateSheet("219Option");
                HSSFRow row = (HSSFRow)sheet.CreateRow(0);
                HSSFCreationHelper createHelper = (HSSFCreationHelper)wb.GetCreationHelper();

                row.CreateCell(0).SetCellValue(createHelper.CreateRichTextString("Option Name"));
                row.CreateCell(1).SetCellValue(createHelper.CreateRichTextString("Input Value"));
                row.CreateCell(2).SetCellValue(createHelper.CreateRichTextString("Output Value"));

                using (FileStream file = new FileStream(folderPath + "\\" + fileName, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(file);
                    file.Close();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Occur error createExcelFile function : " + ex.ToString());
            }
          
        }

        public static void insertExcelData(string folderPath, string fileName, string[] args)
        {
            try
            {
                HSSFWorkbook wb = new HSSFWorkbook();
                HSSFCreationHelper createHelper = (HSSFCreationHelper)wb.GetCreationHelper();

                using (FileStream file = new FileStream(folderPath + "\\" + fileName, FileMode.Open, FileAccess.Read))
                {
                    wb = new HSSFWorkbook(file);
                }

                ISheet sheet = wb.GetSheet("219Option");
                IRow row = sheet.CreateRow(sheet.LastRowNum + 1);

                row.CreateCell(0).SetCellValue(createHelper.CreateRichTextString(args[1]));
                row.CreateCell(1).SetCellValue(createHelper.CreateRichTextString(args[2]));
                row.CreateCell(2).SetCellValue(createHelper.CreateRichTextString(args[3]));

                using (FileStream file = new FileStream(folderPath + "\\" + fileName, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(file);
                    file.Close();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Occur error insertExcelData function : " + ex.ToString());
            }
           
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.SharePoint.Client.Utilities;
using Excel = Microsoft.Office.Interop.Excel;


namespace UploadFile
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter your password.");
            Credentials crd = new Credentials();

            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ExampleGratia"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(crd.userName, crd.password);


                //GetFile(clientContext);
                //AddFiles(clientContext);
                //ReadExcelData(clientContext, "SharePointUploadList.xlsx");
                GetExcelFile(clientContext);
                Console.Read();
            }
        }

        public static void GetFile(ClientContext cxt)
        {

            List UploadToList = cxt.Web.Lists.GetByTitle("LokeshPractice");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View><Query></Query></View>";
            //camlQuery.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='Name'/><Value Type='Text'>SharePointUploadList</Value></Eq></Where></Query></View>";
            ListItemCollection FilesinLib = UploadToList.GetItems(camlQuery);
            cxt.Load(FilesinLib);
            cxt.ExecuteQuery();


            foreach (ListItem file in FilesinLib)
            {
                Console.WriteLine(file.File.Author);
                File ExcelFile = file.File;
                cxt.Load(ExcelFile);
                cxt.ExecuteQuery();
                string FileUrl = ExcelFile.ServerRelativeUrl;

                //Console.WriteLine(file.FieldValues["Title"].ToString());
                Console.WriteLine(file.FieldValues["FileLeafRef"]);
            }



        }

        private static void ReadExcelData(ClientContext clientContext, string fileName)
        {

            string strErrorMsg = string.Empty;
            const string lstDocName = "LokeshPractice";
            try
            {
                DataTable dataTable = new DataTable("ExcelDataTable");
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();

                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            //WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            //Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }

                        for (int datarow = 0; datarow < 3; datarow++)
                        {
                            DataRow r = dataTable.Rows[datarow];
                            AddFiles(clientContext, r);

                        }

                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "0000");
            }
        }





        public static void AddFiles(ClientContext cxt, DataRow row)
        {

            int datacolumn = 0;

            string[] farr = row[datacolumn].ToString().Split('/');
            string FilepathString = row[datacolumn].ToString();
            string FileNameForURL = farr[farr.Length - 1];



            System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilepathString);

            long filesize = fileInfo.Length;
            //var pathstring = @"D:/SPAssessment/FilesToUpload/SharePointUploadList.xlsx";
            if (filesize < 15000)
            {

                List l = cxt.Web.Lists.GetByTitle("LokeshPractice");



                FileCreationInformation fileToUpload = new FileCreationInformation();
                fileToUpload.Content = System.IO.File.ReadAllBytes(FilepathString);
                fileToUpload.Overwrite = true;
                fileToUpload.Url = "LokeshPractice/" + FileNameForURL;


                //fileToUpload.Content.GetLength(filesize);


                // fileToUpload.Url = "LokeshPractice/SharePointUploadList.xlsx";

                //folder.Folders.GetByUrl("LokeshPractice").Folders.GetByUrl("created Folder");

                //  var list = cxt.Web.Lists.GetByTitle("LokeshPractice");
                File uploadfile = l.RootFolder.Files.Add(fileToUpload);

                //File fil = folder.Files.Add(fileToUpload);

                farr = row["Status"].ToString().Split(',');
                ListItem fileitem = uploadfile.ListItemAllFields;
                fileitem["Title"] = "File generated using Code";
                fileitem["Multiselectcheck"] = farr;
                fileitem["File_x0020_Type"] = fileInfo.Extension;
                fileitem["CreatedBy"] = row["CreatedBy"];
                fileitem.Update();
                // cxt.Load(item);

                //uploadfile.Update();
                cxt.ExecuteQuery();
                //Console.ReadLine();
            }
            else
            {
                Console.WriteLine(FileNameForURL + " file size exceed");
            }

        }



        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
            return value;
        }



        public static void GetExcelFile(ClientContext cxt)
        {


            var list = cxt.Web.Lists.GetByTitle("LokeshPractice");
            var listItem = list.GetItemById(14);
            cxt.Load(list);
            cxt.Load(listItem, i => i.File);
            cxt.ExecuteQuery();

            var fileRef = listItem.File.ServerRelativeUrl;
            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(cxt, fileRef);
            var fileName = System.IO.Path.Combine(@"D:\SPAssessment", (string)listItem.File.Name);
            using (var fileStream = System.IO.File.Create(fileName))
            {
                fileInfo.Stream.CopyTo(fileStream);
            }

        }

        public static void ReadData()
        {           

            //Read the excel document using Microsoft Office package
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            int sprowCnt = 0; // row count
            int spcolumnCnt = 0; // column count

            xlApp = new Excel.ApplicationClass();
            xlWorkBook = xlApp.Workbooks.Open(@"<strong>C:\splessons.xlsx</strong>", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (sprowCnt = 1; sprowCnt <= range.Rows.Count; sprowCnt++)
            {
                for (spcolumnCnt = 1; spcolumnCnt <= range.Columns.Count; spcolumnCnt++)
                {
                    Console.WriteLine(" Coulmn Number: " + spcolumnCnt + "--> " + (range.Cells[sprowCnt, spcolumnCnt] as Excel.Range).Value2);
                }
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Console.ReadKey();

           


        }
    }
}

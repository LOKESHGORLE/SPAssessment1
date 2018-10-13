using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using Microsoft.SharePoint.Client.Utilities;
//using System.Web.UI.WebControls.FileUpload;

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


                // GetFile(clientContext);
                //AddFiles(clientContext);
                ReadExcelData(clientContext, "SharePointUploadList.xlsx");

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
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
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
                        
                        for(int datarow=0; datarow < 3; datarow++)
                        {
                            DataRow r = dataTable.Rows[datarow];
                            for(int datacolumn = 0; datacolumn < 1; datacolumn++)
                            {
                                string[] farr= r[datacolumn].ToString().Split('/');
                                string FilePath= r[datacolumn].ToString();
                                string FileName=  farr[farr.Length - 1];
                                AddFiles(clientContext, FilePath, FileName);
                            }
                        }

                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message+"0000");
            }
        }





        public static void AddFiles(ClientContext cxt,string FilepathString,string FileNameForURL)
        {

            
            var pathstring =  FilepathString;
            //var pathstring = @"D:/SPAssessment/FilesToUpload/SharePointUploadList.xlsx";
            List l = cxt.Web.Lists.GetByTitle("LokeshPractice");

            FileCreationInformation fileToUpload = new FileCreationInformation();
            fileToUpload.Content = System.IO.File.ReadAllBytes(pathstring);
            fileToUpload.Overwrite = true;
            fileToUpload.Url = "LokeshPractice/"+ FileNameForURL;
            int filesize = 0;
            fileToUpload.Content.GetLength(filesize);
            // fileToUpload.Url = "LokeshPractice/SharePointUploadList.xlsx";

            //folder.Folders.GetByUrl("LokeshPractice").Folders.GetByUrl("created Folder");

            //  var list = cxt.Web.Lists.GetByTitle("LokeshPractice");
            File uploadfile = l.RootFolder.Files.Add(fileToUpload);

            //File fil = folder.Files.Add(fileToUpload);


            uploadfile.Update();
            cxt.ExecuteQuery();
            //Console.ReadLine();


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
    }
}

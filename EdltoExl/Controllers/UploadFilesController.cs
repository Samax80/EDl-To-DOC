using EdltoExl.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Reflection;

namespace EdltoExl.Controllers
{
    public class UploadFilesController : Controller
    {
        // GET: UploadFiles
        public ActionResult Index()
        {
            ViewBag.ItemList = "";
            return View();
        }

        // POST
        [HttpPost]
        public ActionResult UploadFileXml(HttpPostedFileBase file)
        {

            string fileExtenion = Path.GetExtension(file.FileName);//.edl

            if (file != null && file.ContentLength > 0 && fileExtenion == ".edl")
                try
                {

                    var OriginalEdlFIleName = Path.GetFileName(file.FileName);

                    string serverFolderpath = Server.MapPath("~/FilesXmlUploaded/");


                    string currentEdlfilePathAndName = file.FileName;



                    var temporaryEdlfile = Path.Combine(serverFolderpath, OriginalEdlFIleName);


                    string selectedEdlFileName = Path.GetFileName(file.FileName).Remove(Path.GetFileName(file.FileName).IndexOf('.'));

                    string folderPathOfEdlFile = currentEdlfilePathAndName.Remove(currentEdlfilePathAndName.Length - Path.GetFileName(file.FileName).Length);

                    file.SaveAs(temporaryEdlfile);


                    string[] lines = System.IO.File.ReadAllLines(temporaryEdlfile);

                    IEnumerable<string> selectLines = lines.Where(line => line.Contains(".wav"));

                    List<InfoMusic> MusicList_name = new List<InfoMusic>();
                    List<InfoMusic> MusicList_time = new List<InfoMusic>();

                    int indexof_ = 0;
                    int indexofSecond_ = 0;
                    int indexOfpoint = 0;

                    foreach (var item in selectLines)
                    {


                        InfoMusic Musicobject = new InfoMusic();

                        Musicobject.Music_FullName = item.Substring(18);

                        indexof_ = Musicobject.Music_FullName.IndexOf('_');

                        Musicobject.Music_Cd = Musicobject.Music_FullName.Remove(indexof_).Remove(0, 3);

                        Musicobject.Music_Cue = Musicobject.Music_FullName.Remove(0, indexof_ + 1).Remove(Musicobject.Music_FullName.Remove(0, indexof_ + 1).IndexOf('_'));

                        indexofSecond_ = Musicobject.Music_FullName.Remove(0, indexof_ + 1).IndexOf('_');

                        indexOfpoint = Musicobject.Music_FullName.Remove(0, indexof_ + 2).Remove(0, indexofSecond_).IndexOf('.');

                        Musicobject.Music_Name = Musicobject.Music_FullName.Remove(0, indexof_ + 2).Remove(0, indexofSecond_).Remove(indexOfpoint).Replace("-", " ");

                        MusicList_name.Add(Musicobject);

                    }

                    IEnumerable<string> selectLinestime = lines.Where(line => line.Contains("000"));
                    foreach (var itemb in selectLinestime)
                    {
                        InfoMusic Musicobject = new InfoMusic();


                        string wavMusicInfoLine = itemb;

                        int totalLenght = (wavMusicInfoLine.Length);

                        wavMusicInfoLine = wavMusicInfoLine.Remove(0, (totalLenght - 24));

                        string timeInpre = wavMusicInfoLine.Remove(9);


                        string timeOutpre = wavMusicInfoLine.Remove(0, (wavMusicInfoLine.Length - 12));

                        timeOutpre = timeOutpre.Remove(9);

                        Musicobject.Music_Tc_InTime = DateTime.Parse(timeInpre).TimeOfDay;
                        Musicobject.Music_Tc_OutTime = DateTime.Parse(timeOutpre).TimeOfDay;
                        Musicobject.Music_Tc_Duration = Musicobject.Music_Tc_OutTime - Musicobject.Music_Tc_InTime;


                        MusicList_time.Add(Musicobject);

                    }


                    int cellTimeNumber = 10;
                    int cellNameNumber = 10;

                    int tcIncellD = 4; //"d";
                    int tcoutcellE = 5; //"e";
                    int tcdurcellF = 6;//"f";
                    int usecellG = 7;//"g";
                    int cdCellH = 8;//"h";
                    int cueCellI = 9;//"i";
                    int nameCellJ = 10;//"j";
                    int fullnamecellK = 11; //"k";
                    int publicshercellL = 12;//"l";

                    string fileName = ("Music_Cue_Template");

                    string excelExtension = (".xlsx");




                    var templateFileLocation = (serverFolderpath + fileName + excelExtension);


                    //locating template 
                    var TemplateFile = new FileInfo(templateFileLocation);

                    //new code dic 12


                    //Creating the new xlsx file to work with

                    ExcelPackage excelPackage = new ExcelPackage(TemplateFile, TemplateFile);//use the templte file and  ceate a new(copy ) edited file !!
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];


                    //end new code dic 12


                    var datePrepared = DateTime.Now.ToString("MM/dd/yyyy");
                    worksheet.Cells[6, 12].Value = datePrepared;

                    var productionTitle = selectedEdlFileName;

                    worksheet.Cells[2,4].Value = productionTitle;


                    foreach (var musictime in MusicList_time)

                    {
                        worksheet.Cells[cellTimeNumber, tcIncellD].Value = musictime.Music_Tc_InTime.ToString();
                        worksheet.Cells[cellTimeNumber, tcoutcellE].Value = musictime.Music_Tc_OutTime.ToString();
                        worksheet.Cells[cellTimeNumber, tcdurcellF].Value = musictime.Music_Tc_Duration.ToString();
                        worksheet.Cells[cellTimeNumber, usecellG].Value = "BI";

                        //excel.WriteToCell(cellTimeNumber, tcIncellD, musictime.Music_Tc_InTime.ToString());

                        //excel.WriteToCell(cellTimeNumber, tcoutcellE, musictime.Music_Tc_OutTime.ToString());

                        //excel.WriteToCell(cellTimeNumber, tcdurcellF, musictime.Music_Tc_Duration.ToString());

                        //excel.WriteToCell(cellTimeNumber, usecellG, "BI");

                        cellTimeNumber++;

                    }


                    foreach (var musicname in MusicList_name)
                    {
                        worksheet.Cells[cellNameNumber, cdCellH].Value = musicname.Music_Cd;
                        worksheet.Cells[cellNameNumber, cueCellI].Value = musicname.Music_Cue;
                        worksheet.Cells[cellNameNumber, nameCellJ].Value = musicname.Music_Name.ToString();
                        worksheet.Cells[cellNameNumber, fullnamecellK].Value = musicname.Music_FullName.ToString();
                        worksheet.Cells[cellNameNumber, publicshercellL].Value = "Audio Network Limited (PRS)";

                        //excel.WriteToCell(cellNameNumber, cdCellH, musicname.Music_Cd);

                        //excel.WriteToCell(cellNameNumber, cueCellI, musicname.Music_Cue);

                        //excel.WriteToCell(cellNameNumber, nameCellJ, musicname.Music_Name.ToString());

                        //excel.WriteToCell(cellNameNumber, fullnamecellK, musicname.Music_FullName.ToString());

                        //excel.WriteToCell(cellNameNumber, publicshercellL, "Audio Network Limited (PRS)");

                        cellNameNumber++;
                    }

                    //save xlsx file at server folder
                    var pathToSaveXlsxFile = serverFolderpath + selectedEdlFileName + excelExtension;

                    //save xlsx file at user folder
                    //  var pathToSaveXlsxFile = folderPathOfEdlFile + selectedEdlFileName + excelExtension;


                    //excel.SaveAs(pathToSaveXlsxFile);
                    //excel.close();
                    FileInfo xlsxFileToUser = new FileInfo(pathToSaveXlsxFile);
                    excelPackage.SaveAs(xlsxFileToUser);




                    //TO DOWNLOAD FROM THE WEB Automatically(to be improved)


                    System.IO.FileInfo filetoDownload = new System.IO.FileInfo(pathToSaveXlsxFile);//it looks for the temporary file xlsx created at  server folder  the to be downloaded!!

                    Response.Clear();
                    Response.ClearContent();
                    Response.ClearHeaders();
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + selectedEdlFileName + excelExtension);
                    Response.AddHeader("Content-Length", filetoDownload.Length.ToString());
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.WriteFile(filetoDownload.FullName);
                    Response.Flush();
                    Response.Close();

                    //delete temporary files
                    System.IO.File.Delete(temporaryEdlfile);
                    System.IO.File.Delete(pathToSaveXlsxFile);


                    ViewBag.Message = "File saved successfully to " + folderPathOfEdlFile;


                }
                catch (Exception ex)
                {
                    ViewBag.Message = "ERROR:" + ex.Message.ToString();
                }
            else
            {
                ViewBag.Message = "You have not specified a .edl file.";
            }
            return View("Index");
        }
    }



}
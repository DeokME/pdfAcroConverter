using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;

namespace pdfAcroConverter
{
    class Program
    {
        static void Main(string[] args)
        {

            if (args.Length > 2)
            {

                string originalFile = args[0];
                string outputPath = args[1];
                string type = args[2];

                if (File.Exists(originalFile))//Check File exxists
                {
                    string fileName = Path.GetFileNameWithoutExtension(originalFile);
                    string fileExtension = Path.GetExtension(originalFile);

                    Acrobat.AcroPDDoc pdfDoc = new Acrobat.AcroPDDoc();

                    pdfDoc.Open(@originalFile);

                    object jsObj = pdfDoc.GetJSObject();
                    var jsObjType = jsObj.GetType();

                    string target = outputPath + "/" + fileName + "." + type;
                    string exportFormat = null;
                    switch (type)
                    {
                        case "eps":
                            exportFormat = "com.adobe.acrobat.eps";
                            break;
                        case "html":
                        case "htm":
                            exportFormat = "com.adobe.acrobat.html";
                            break;
                        case "jpg":
                        case "jpeg":
                        case "jpe":
                            exportFormat = "com.adobe.acrobat.jpeg";
                            break;
                        case "doc":
                            exportFormat = "com.adobe.acrobat.doc";
                            break;
                        case "docx":
                            exportFormat = "com.adobe.acrobat.docx";
                            break;
                        case "png":
                            exportFormat = "com.adobe.acrobat.png";
                            break;
                        case "ps":
                            exportFormat = "com.adobe.acrobat.ps";
                            break;
                        case "rtf":
                            exportFormat = "com.adobe.acrobat.rtf";
                            break;
                        case "xlsx":
                            exportFormat = "com.adobe.acrobat.xlsx";
                            break;
                       /* Doesn't work well
                        case "xls":
                            exportFormat = "com.adobe.acrobat.spreadsheet";
                            break;
                        */
                        case "txt":
                            exportFormat = "com.adobe.acrobat.accesstext";
                            break;
                        case "tiff":
                        case "tif":
                            exportFormat = "com.adobe.acrobat.tiff";
                            break;
                        case "xml":
                            exportFormat = "com.adobe.acrobat.xml-1-00";
                            break;

                    }
                    
                    jsObjType.InvokeMember("saveAs",
                                                BindingFlags.InvokeMethod |
                                                BindingFlags.Public |
                                                BindingFlags.Instance,
                                                null, jsObj, new object[] { @target, exportFormat });

                    Console.WriteLine("Convert complete");
                }
                else
                {
                    Console.WriteLine("PDF doesn't exists");
                }


            }
            else
            {
                Console.WriteLine("Not enough args");

            }

        }
    }
}

using ExportImplementation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using RazorEngine;
using RazorEngine.Templating;
using System.Xml;

namespace ODSConversionTest
{
    public class CustomExportODS<T> : ExportODS<T> where T : class
    {
  
        internal byte[] GenerateODS(string textSheet)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(Templates.ods, 0, Templates.ods.Length);
                using (var za = new ZipArchive(ms, ZipArchiveMode.Update))
                {
                    za.GetEntry("content.xml").Delete();
                    var c = za.CreateEntry("content.xml");
                    using (var s = c.Open())
                    {
                        using (var writer = new StreamWriter(s))
                        {
                            writer.Write(textSheet);
                        }
                    }
                }
                return ms.ToArray();
            }
        }


        private byte[] CreateODS(string textSheet)
        {
            //textSheet = @"D:\ODSConversionTest\ODSConversionTest\bin\Debug\data1.xml";
            //textSheet = @"C:\Users\joseph.lai\Desktop\sample.xml";
            using (var ms = new MemoryStream())
            {
                ms.Write(Templates.ods, 0, Templates.ods.Length);
                using (var za = new ZipArchive(ms, ZipArchiveMode.Update))
                {
                    za.GetEntry("content.xml").Delete();
                    var c = za.CreateEntry("content.xml");
                    using (var s = c.Open())
                    {
                        using (var writer = new StreamWriter(s))
                        {
                            writer.Write(textSheet);
                        }
                    }
                }
                return ms.ToArray();
            }
        }
    }
}

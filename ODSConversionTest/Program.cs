using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;
using ExportImplementation;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.Threading;
using System.Reflection;
using System.Reflection.Emit;
using System.Collections;
using System.Xml;
using OfficeOpenXml;

namespace ODSConversionTest
{
    class Program
    {
        static int propNum = 0;
        static List<string> keyList = new List<string>();
        static List<Object> ODSResult = new List<object>();
        static int totalNum = default(int);
        static string sheetFinalName;

        static void Main(string[] args)
        {
            string outputExcelFile = "output.xlsx";
            RemoveExcelLastRow();
            GetExcelToJson(outputExcelFile);
            Type custDataType = BuildDynamicTypeWithProperties();
            dynamic jsonResult = LoadJson(sheetFinalName + ".json");
            var str = JsonToDefaultODS(jsonResult, custDataType);
            var strXml = PerformXML(str);
            var export = new CustomExportODS<Object>();
            var data = export.GenerateODS(strXml);
            File.WriteAllBytes("result.ods", data);
            File.Delete(sheetFinalName + ".json");
            File.Delete(outputExcelFile);

            Console.ReadKey();
        }

        private static void RemoveExcelLastRow()
        {
            using (FileStream fs = new FileStream(@"C:\Users\joseph.lai\Desktop\excelFolder\sampleODS.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //載入Excel檔案
                using (ExcelPackage ep = new ExcelPackage(fs))
                {
                    var ws = ep.Workbook.Worksheets.FirstOrDefault();
                    int endRow = ws.Dimension.End.Row;
                    int colCount = ws.Dimension.End.Column;
                    for (int i = 1; i <= colCount; i++)
                    {
                        ws.DeleteRow(endRow, i, true);
                    }
                    ep.SaveAs(new FileInfo("output.xlsx"));
                }
            }
        }

        private static string PerformXML(string str)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(str);
            XmlNodeList nodes = doc.GetElementsByTagName("table:table-row");
            for (int i = 0; i < nodes.Count - 1; i++)
            {
                XmlNode item = nodes[i];
                for (int j = 1; j <= keyList.Count; j++)
                {
                    XmlElement element = doc.CreateElement("table-cell");
                    element.SetAttribute("style-name", "Standard");
                    element.SetAttribute("value-type", "string");
                    item.AppendChild(element);

                    XmlElement tmpElm = doc.CreateElement("p");
                    string value = ODSResult[i].GetType().GetProperty("CustomerName" + j).GetValue(ODSResult[i]) as string;
                    tmpElm.InnerText = value;
                    element.AppendChild(tmpElm);
                }
            }

            StringWriter sw = new StringWriter();
            XmlTextWriter tx = new XmlTextWriter(sw);
            doc.WriteTo(tx);
            string finalStr = sw.ToString();
            finalStr = finalStr.Replace("table-cell", "table:table-cell");
            finalStr = finalStr.Replace("style-name", "table:style-name");
            finalStr = finalStr.Replace("value-type", "office:value-type");
            finalStr = finalStr.Replace("<p>", "<text:p>");
            finalStr = finalStr.Replace("</p>", "</text:p>");
            finalStr = finalStr.Replace("table:table:style-name", "table:style-name");
            return finalStr;
        }

        private static string JsonToDefaultODS(dynamic jsonResult, Type custDataType)
        {
            Type classType = BuildDynamicTypeWithProperties();
            dynamic jsonList2 = jsonResult[sheetFinalName.ToUpper()];
            for (int i = 1; i <= totalNum; i++)
            {
                var item = jsonList2["" + i];
                var custData = Activator.CreateInstance(classType);
                for (int j = 1; j <= keyList.Count; j++)
                {
                    var key = keyList[j - 1];
                    string currValue = item[key + ""];
                    classType.InvokeMember("CustomerName" + j, BindingFlags.SetProperty,
                                          null, custData, new object[] { currValue });
                }
                ODSResult.Add(custData);
            }
            var export = new CustomExportODS<Object>();
            return export.ExportResultStringPart(ODSResult);
        }

        public static dynamic LoadJson(string filePath)
        {
            using (StreamReader r = new StreamReader(filePath))
            {
                string json = r.ReadToEnd();
                return Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(json);
            }
        }

        public static Type BuildDynamicTypeWithProperties()
        {
            AppDomain myDomain = Thread.GetDomain();
            AssemblyName myAsmName = new AssemblyName();
            myAsmName.Name = "MyDynamicAssembly";

            // To generate a persistable assembly, specify AssemblyBuilderAccess.RunAndSave.
            AssemblyBuilder myAsmBuilder = myDomain.DefineDynamicAssembly(myAsmName,
                                                            AssemblyBuilderAccess.RunAndSave);
            // Generate a persistable single-module assembly.
            ModuleBuilder myModBuilder =
                myAsmBuilder.DefineDynamicModule(myAsmName.Name, myAsmName.Name + ".dll");

            TypeBuilder myTypeBuilder = myModBuilder.DefineType("ODSData",
                                                            TypeAttributes.Public);

            for (int i = 1; i <= propNum; i++)
            {
                SetCustomClassProperty(myTypeBuilder, "customerName" + i, "CustomerName" + i);
            }
            Type retval = myTypeBuilder.CreateType();

            return retval;
        }

        private static void SetCustomClassProperty(TypeBuilder myTypeBuilder, string fieldName, string propertyName)
        {
            FieldBuilder customerNameBldr = myTypeBuilder.DefineField(fieldName,
                                                            typeof(string),
                                                            FieldAttributes.Private);
            PropertyBuilder custNamePropBldr = myTypeBuilder.DefineProperty(propertyName,
                                                             System.Reflection.PropertyAttributes.HasDefault,
                                                             typeof(string),
                                                             null);

            MethodAttributes getSetAttr =
                MethodAttributes.Public | MethodAttributes.SpecialName |
                    MethodAttributes.HideBySig;

            MethodBuilder custNameGetPropMthdBldr =
                myTypeBuilder.DefineMethod("get_" + propertyName,
                                           getSetAttr,
                                           typeof(string),
                                           Type.EmptyTypes);

            ILGenerator custNameGetIL = custNameGetPropMthdBldr.GetILGenerator();

            custNameGetIL.Emit(OpCodes.Ldarg_0);
            custNameGetIL.Emit(OpCodes.Ldfld, customerNameBldr);
            custNameGetIL.Emit(OpCodes.Ret);

            MethodBuilder custNameSetPropMthdBldr =
                myTypeBuilder.DefineMethod("set_" + propertyName,
                                           getSetAttr,
                                           null,
                                           new Type[] { typeof(string) });

            ILGenerator custNameSetIL = custNameSetPropMthdBldr.GetILGenerator();

            custNameSetIL.Emit(OpCodes.Ldarg_0);
            custNameSetIL.Emit(OpCodes.Ldarg_1);
            custNameSetIL.Emit(OpCodes.Stfld, customerNameBldr);
            custNameSetIL.Emit(OpCodes.Ret);

            custNamePropBldr.SetGetMethod(custNameGetPropMthdBldr);
            custNamePropBldr.SetSetMethod(custNameSetPropMthdBldr);
        }


        private static void GetExcelToJson(string fileName)
        {
            if (fileName.EndsWith("xlsx") || fileName.EndsWith("xls"))
            {
                if (File.Exists(fileName))
                {

                    FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);

                    IWorkbook workbook = new XSSFWorkbook(file);
                    WriteJson(workbook);
                }
            }

        }

        private static void WriteJson(IWorkbook workbook)
        {
            //服务端表头位置
            var server_title = 1;
            //服务端数据开始行数
            var server_num = 4;

            //获取excel的第一个sheet

            string sheet_name = workbook.GetSheetName(0);
            sheetFinalName = sheet_name;

            ISheet sheet = workbook.GetSheetAt(0);

            try
            {

                string txtPath = sheet_name.ToLower() + ".json";
                FileStream aFile = new FileStream(txtPath, FileMode.OpenOrCreate);
                System.Text.Encoding encode = System.Text.Encoding.GetEncoding("utf-8"); 
                StreamWriter sw = new StreamWriter(aFile,encode);
                sw.Write("{\r\n");
                sw.Write("   \"" + sheet_name.Trim().ToUpper() + "\":{\r\n");


                //获取sheet的第二行，服务端用的表头
                IRow titleRow = sheet.GetRow(server_num);

                foreach (var item in titleRow.Cells)
                {
                    if (!string.IsNullOrEmpty(item.ToString()))
                    {
                        ++propNum;
                        keyList.Add(item.ToString());
                    }
                }
                //一行最后一个方格的编号 即总的列数
                int cellCount = titleRow.LastCellNum;

                //最后一行
                int rowCount = sheet.LastRowNum;

                totalNum = sheet.LastRowNum - server_num + 1;
                //遍历行
                for (int i = server_num; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        break;
                    }

                    string str_write = "        \"" + (i - server_num + 1) + "\":{\r\n";
                    sw.Write(str_write);


                    string str_write2 = "";
                    //遍历该行的列
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {

                        if (titleRow.GetCell(j) != null && titleRow.GetCell(j).ToString().Length != 0)
                        {
                            string value = "";

                            if (row.GetCell(j) != null)
                            {
                                value = row.GetCell(j).ToString().Trim();
                            }

                            string title = titleRow.GetCell(j).ToString().Trim();

                            str_write2 += "          \"" + title + "\":\"" + value + "\",\r\n";

                        }

                    }

                    int end_str2 = str_write2.LastIndexOf(",");
                    if (end_str2 != -1)
                    {
                        str_write2 = str_write2.Remove(end_str2, 1);
                        sw.Write(str_write2);
                    }




                    if (i == rowCount)
                    {
                        sw.Write("        }\r\n");
                    }
                    else
                    {
                        sw.Write("        },\r\n");
                    }


                }




                sw.Write("   }\r\n");
                sw.Write("}\r\n");
                sw.Close();
            }
            catch (IOException ex)
            {
                return;
            }

        }
    }
}

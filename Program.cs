using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace BSTReportPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            string dirpathIN = @ConfigurationManager.AppSettings.Get("INdir");
            string dirpathOUT = @ConfigurationManager.AppSettings.Get("OUTdir");
            var dirIN = new DirectoryInfo(dirpathIN); // папка с файлами
            var dirOUT = new DirectoryInfo(dirpathOUT);
            DataTable CSVtable = new DataTable();
            string inFile = "";
            string outFile = "";

            foreach (FileInfo file in dirIN.GetFiles())
            {
                inFile = Path.GetFileName(file.FullName);
                Console.WriteLine(inFile);
                CSVtable = CSVUtility.CSVUtility.GetDataTableFromCSVFile(file.FullName); //Получаем DataTable из CSV файла
                outFile = dirpathOUT + "\\" + inFile;
                
                ReportR08(CSVtable, outFile);

                              
                CSVtable = null;
                //GC.Collect(1, GCCollectionMode.Forced);
            }

        }

        static void FixFiles(string inDir, string outDir)
        {
            var dirIN = new DirectoryInfo(@inDir); // папка с входящими файлами 
            var dirOUT = new DirectoryInfo(@outDir); // папка с исходящими файлами             
            string fileName = "";

            foreach (FileInfo file in dirIN.GetFiles())
            {
                fileName = Path.GetFileName(file.FullName);
                Console.WriteLine(fileName);
                //FixReport(@file.FullName, @outDir + @"\" + inFile);
            }
        }

        static void ReportR08 (DataTable inTable, string fileName)
        {
            DataTable reestr = new DataTable();
            DataColumn column;
            DataRow reestrRow;

            #region Задаем структуру таблицы reestr
            //1. AccountOperator (ИНН оператора ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountOperator";
            column.AllowDBNull = true;
            column.DefaultValue = null;
            reestr.Columns.Add(column);

            //2. StreetCode (Код улицы КЛАДР)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "StreetCode";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //3. BuildingNumber (Номер дома)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "BuildingNumber";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //4. PremisesNumber (Номер помещения)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PremisesNumber";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //5. PremisesPart (Часть помещения)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PremisesPart";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //6. AccountNum (Номер ЛС (ФЛС))
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountNum";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //7. BeginDate (Дата открытия лицевого счета)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "BeginDate";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //8. CloseDate (Дата закрытия лицевого счета)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "CloseDate";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //9. Family (Фамилия абонента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Family";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //10. Name (Имя абонента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Name";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //11. Lastname (Отчество абонента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Lastname";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //12. BirthDate (Дата рождения абонента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "BirthDate";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //13. Sex (Пол абонента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Sex";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //14. INN (ИНН юридического лица)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "INN";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //15. KPP (КПП юридического лица)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "KPP";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //16. CompanyName (Наименование юридического лица)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "CompanyName";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //17. TypePremises (Тип помещения)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "TypePremises";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //18. FormOfOwnership (Форма собственности помещения)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "FormOfOwnership";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //19. TotalArea (Общая площадь помещения)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "TotalArea";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //20. OwnPart (Доля в собственности абонента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "OwnPart";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //21. Comment (Комментарии)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Comment";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //22. ExtNumber (Номер во внешней системе)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ExtNumber";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //23. GISAcount (Единый ЛС в ГИС ЖКХ)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "GISAcount";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //24. GISService (Идентификатор ЖКУ в ГИС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "GISService";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //25. GISAccGUID (Идентификатор ЛС в ГИС ЖКХ)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "GISAccGUID";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //26. AddrDelivDescr (Описание адреса доставки из дополнительных параметров ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AddrDevilDescr";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //27. Deliver (Способ доставки)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Deliver";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //28. CloseReason (Причина закрытия ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "CloseReason";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            #endregion

            ////Вывод структуры таблицы
            //foreach (DataColumn col in inTable.Columns)
            //    Console.Write(col.ColumnName);

            Console.WriteLine("Заполняем реестр");
            foreach (DataRow row in inTable.Rows)
            {
                reestrRow = reestr.NewRow();
                reestrRow["StreetCode"] = row["StreetCode"];
                reestrRow["BuildingNumber"] = row["BuildingNumber"];
                reestrRow["PremisesNumber"] = row["PremisesNumber"];
                reestrRow["AccountNum"] = row["AccountNum"];
                reestrRow["BeginDate"] = row["BeginDate"];
                reestrRow["CloseDate"] = row["CloseDate"];
                reestrRow["Family"] = row["FIO"];
                reestrRow["INN"] = row["INN"];
                reestrRow["KPP"] = row["KPP"];
                reestrRow["CompanyName"] = row["CompanyName"];
                reestrRow["TypePremises"] = row["TypePremises"];
                reestrRow["FormOfOwnership"] = row["FormOfOwnership"];
                reestrRow["TotalArea"] = row["TotalArea"];
                reestrRow["Comment"] = row["Comment"];
                reestrRow["Deliver"] = row["Deliver"];
                reestrRow["CloseReason"] = row["CloseReason"];
                reestrRow["AddrDelivDescr"] = row["Adress"];

                reestr.Rows.Add(reestrRow);
            }

            CSVUtility.CSVUtility.ToCSV(reestr, fileName);

            reestr.Dispose();
        }

    }
}

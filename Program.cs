using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;



namespace BSTReportPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            string dirpathIN = @ConfigurationManager.AppSettings.Get("INdir");
            string dirpathOUT = @ConfigurationManager.AppSettings.Get("OUTdir");
            string repType = ConfigurationManager.AppSettings.Get("REPORT_TYPE");

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
                
                switch (repType)
                {
                    case "0":
                        CSVUtility.CSVUtility.ToCSV(CSVtable, outFile);
                        break;
                    case "08":
                        ReportR08(CSVtable, outFile);
                        break;
                    case "16":
                        ReportR16(CSVtable, outFile);
                        break;
                    case "16P":
                        ReportR16P(CSVtable, outFile);
                        break;
                    default:
                        CSVUtility.CSVUtility.ToCSV(CSVtable, outFile);
                        break;
                }                
                              
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

        static void ReportR16 (DataTable inTable, string fileName)
        {
            DataTable reestr = new DataTable();
            DataColumn column;
            DataRow reestrRow;
            string header = "#RTYPE=R16\n"
                            +"\n"
                            +"#AccountOperator;AccountNum;ServiceCode;ProviderCode;ChargeYear;ChargeMonth;SaldoIn;ChargeVolume;"
                            +"Tarif;ChargeSum;RecalSum;PaySum;SaldoOut;SaldoFineIn;FineSum;PayFineSum;CorrectFineSum;SaldoFineOut;"
                            +"LastPayDate;PayAgent;PrivChargeSum;PrivRecalSum;PrivCategory;PrivPaySum";

            #region Задаем структуру таблицы reestr
            //1. AccountOperator (ИНН оператора ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountOperator";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //2. Account Num (Номер ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountNum";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //3. ServiceCode (Услуга)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ServiceCode";
            column.AllowDBNull = false;
            column.DefaultValue = "22";
            reestr.Columns.Add(column);

            //4. ProviderCode (ИНН поставщика услуг)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ProviderCode";
            column.AllowDBNull = false;
            column.DefaultValue = "5190996259";
            reestr.Columns.Add(column);

            //5. ChargeYear (Год отчетного периода)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ChargeYear";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //6. ChargeMonth (Отчетный месяц)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ChargeMonth";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //7. SaldoIn (Остаток задолженности по взносам на начало отчетного месяца)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SaldoIn";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //8. ChargeVolume (Площадь помещения)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ChargeVolume";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //9. Tarif (Тариф по взносам в фонд капитального ремонта)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Tarif";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //10. ChargeSum (Сумма начисления в отчетном месяце)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ChargeSum";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //11. RecalcSum (Сумма перерасчета в отчетном месяц)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "RecalSum";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //12. PaySum (Оплата по взносам в фонд капитального ремонта)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PaySum";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //13. SaldoOut (Остаток задолженности (только по начислениям) на конец отчетного периода)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SaldoOut";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //14. SaldoFineIn (Остаток задолженности по пени на начало месяца)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SaldoFineIn";
            column.AllowDBNull = false;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //15. FineSum (Сумма пени, начисленная в отчетном месяце)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "FineSum";
            column.AllowDBNull = false;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //16. PayFineSum (Оплата пени в отчетном месяце)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PayFineSum";
            column.AllowDBNull = false;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //17. CorrectFineSum (Корректировка пени на конец месяца)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "CorrectFineSum";
            column.AllowDBNull = false;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //18. SaldoFineOut (Остаток задолженности по пени (только начисления) на конец месяца)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SaldoFineOut";
            column.AllowDBNull = false;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //19. LastPayDate (Дата последней оплаты)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "LastPayDate";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //20. PayAgent (Код платежного агента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PayAgent";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //21. PrivChargeSum (Сумма начисления льготы в отчетном месяце)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PrivChargeSum";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //22. PrivRecalSum (Сумма перерасчета льготы в отчетном месяце)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PrivRecalSum";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //23. PrivCategory (Код категории льготника)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PrivCategory";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //24. PrivPaySum (Оплата пени в отчетном месяце)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PrivPaySum";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            #endregion 

            Console.WriteLine("Заполняем реестр R16");
            foreach (DataRow row in inTable.Rows)
            {
                reestrRow = reestr.NewRow();
                reestrRow["AccountNum"] = row["AccountNum"];
                reestrRow["ChargeYear"] = row["ChargeYear"];
                reestrRow["ChargeMonth"] = row["ChargeMonth"];
                reestrRow["SaldoIn"] = FixSum (row["SaldoIn"].ToString());
                reestrRow["ChargeVolume"] = FixSum(row["ChargeVolume"].ToString());
                reestrRow["Tarif"] = FixSum(row["Tarif"].ToString());
                reestrRow["ChargeSum"] = FixSum(row["ChargeSum"].ToString());
                reestrRow["RecalSum"] = FixSum(row["RecalSum"].ToString());
                reestrRow["SaldoOut"] = FixSum(row["SaldoOut"].ToString());
                reestrRow["LastPayDate"] = row["LastPayDate"].ToString();

                reestr.Rows.Add(reestrRow);
            }
            CSVUtility.CSVUtility.ToCSV(reestr, fileName, header);

            reestr.Dispose();
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
            column.DefaultValue = "0";
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
            column.ColumnName = "AddrDelivDescr";
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

            Console.WriteLine("Заполняем реестр R08");
            foreach (DataRow row in inTable.Rows)
            {
                reestrRow = reestr.NewRow();
                reestrRow["StreetCode"] = row["StreetCode"];
                reestrRow["BuildingNumber"] = Regex.Replace(row["BuildingNumber"].ToString(), @"^\d+\w{1}(?:.*)", m => m.Value.ToString().ToUpper());
                reestrRow["PremisesNumber"] = row["PremisesNumber"];
                reestrRow["AccountNum"] = row["AccountNum"];
                reestrRow["BeginDate"] = row["BeginDate"];
                reestrRow["CloseDate"] = row["CloseDate"];
                if (row["FIO"].ToString().Length > 150)
                {
                    string[] FIO = SplitFIO(row["FIO"].ToString());
                    reestrRow["Family"] = FIO[0];
                    reestrRow["Name"] = FIO[1];
                }
                else
                    reestrRow["Family"] = row["FIO"];
                reestrRow["INN"] = row["INN"];
                reestrRow["KPP"] = row["KPP"];
                reestrRow["CompanyName"] = row["CompanyName"];
                if (row["TypePremises"].ToString() == "500301")
                    reestrRow["TypePremises"] = "2";
                else
                    reestrRow["TypePremises"] = row["TypePremises"];
                reestrRow["FormOfOwnership"] = row["FormOfOwnership"];
                reestrRow["TotalArea"] = row["TotalArea"];
                reestrRow["Comment"] = row["Comment"];
                reestrRow["Deliver"] = row["Deliver"];
                reestrRow["CloseReason"] = row["CloseReason"];
                reestrRow["AddrDelivDescr"] = row["Address"];

                reestr.Rows.Add(reestrRow);
            }

            CSVUtility.CSVUtility.ToCSV(reestr, fileName);

            reestr.Dispose();
        }

        static void ReportR16P (DataTable inTable, string fileName)
        {
            DataTable reestr = new DataTable();
            DataColumn column;
            DataRow reestrRow;

            string header = "#RTYPE=R16P\n"
                            + "\n"
                            + "#AccountOperator;AccountNum;ServiceCode;ProviderCode;PaySum;PayFineSum;LastPayDate;PayAgent;PayID;SpecAccount;Comment";


            #region Задаем структуру таблицы reestr
            //1. AccountOperator (ИНН оператора ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountOperator";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);
            
            //2. AccountNum (Номер ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountNum";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //3. ServiceCode (Услуга)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ServiceCode";
            column.AllowDBNull = false;
            column.DefaultValue = "22";
            reestr.Columns.Add(column);

            //4. ProviderCode (ИНН поставщика услуги)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ProviderCode";
            column.AllowDBNull = false;
            column.DefaultValue = "5190996259";
            reestr.Columns.Add(column);

            //5. PaySum (Сумма по взносам в фонд капитального ремонта)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PaySum";
            column.AllowDBNull = false;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //6. PayFineSum (Сумма платежа по пени)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PayFineSum";
            column.AllowDBNull = true;
            column.DefaultValue = "0.00";
            reestr.Columns.Add(column);

            //7. LastPayDate (Номер ЛС)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "LastPayDate";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //8. PayAgent (Код платежного агента)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PayAgent";
            column.AllowDBNull = false;
            column.DefaultValue = "MR1010";
            reestr.Columns.Add(column);

            //9. Код платежа (PayID)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "PayID";
            column.AllowDBNull = false;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            //10. Комментарии (Comment)
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Comment";
            column.AllowDBNull = true;
            column.DefaultValue = "";
            reestr.Columns.Add(column);

            #endregion

            Console.WriteLine("Заполняем реестр R16P");
            foreach (DataRow row in inTable.Rows)
            {
                reestrRow = reestr.NewRow();
                reestrRow["AccountNum"] = row["AccountNum"];
                reestrRow["PaySum"] = FixSum(row["PaySum"].ToString());
                reestrRow["LastPayDate"] = row["PayDate"];
                reestrRow["PayID"] = row["PayDocId"];
                reestrRow["Comment"] = row["ByReason"];

                reestr.Rows.Add(reestrRow);
            }
            CSVUtility.CSVUtility.ToCSV(reestr, fileName, header);

            reestr.Dispose();
        }

        static string[] SplitFIO (string fio)
        {

            //Console.WriteLine(row["AccountNum"].ToString() + "  " + row["FIO"].ToString().Length);
            var words = fio.ToString().Split(new Char[] { ' ' });
            int maxLengthString = 150;
            int wordIndex = 0;          
            string[] splitFIO = new string[3];
            string spaceLetter = " ";
            StringBuilder currentLine = new StringBuilder();
            foreach (string word in words)
            {
                if (currentLine.Length + word.Length + 1 > maxLengthString)// Определяем не привысила ли текущая строка максимальную длину
                {

                    splitFIO[wordIndex] = currentLine.ToString();
                    currentLine.Remove(0, currentLine.Length);
                    wordIndex++;
                    maxLengthString = 20;
                }
                else
                {
                    currentLine.Append(word);
                    currentLine.Append(spaceLetter);                                     
                }
            }
            //wordIndex++;
            //splitFIO[wordIndex] = currentLine.ToString();
            return splitFIO;

        }

        static string FixSum (string sum)
        {
            string fixSum;
            fixSum = Regex.Replace(sum, @"(?:^(,|\.){1})(\d+)", "0.$2");           
            fixSum = Regex.Replace(fixSum, ",", ".");
            
            return fixSum;
        }
    }
}

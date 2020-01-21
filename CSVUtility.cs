using Microsoft.VisualBasic.FileIO;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;


namespace CSVUtility
{
    public static class CSVUtility
    {
        public static void ToCSV(this DataTable dtDataTable, string strFilePath, string header = null)
        //Сохраняем DataTable в CSV файл
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.Unicode);
            

            //Паттерн для поиска разделителя в полях таблицы
            string pattern = ";+";

            //Выводим заголовок на основании названия полей в DataTable
            if (header == null)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    sw.Write(dtDataTable.Columns[i]);
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
            }
            
            //Выводим заголовок указанный в header
            else
                sw.Write(@header);


            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if(!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        //if (value.Contains(';'))
                        //{
                        //value = String.Format("\\{0}\\", value);
                        //value = ANSItoUTF(value);
                        value = Regex.Replace(value, @"\n+", " ");
                        value = Regex.Replace(value, pattern, ":");
                        sw.Write(value);
                        //}
                        //else
                        //{
                        //    sw.Write(dr[i].ToString());
                        //}
                    }
                    else if (Convert.IsDBNull(dr[i]))
                    {
                        sw.Write("");
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        public static DataTable GetDataTableFromCSVFile(string csv_file_path)
        //Копируем данные из CSV в DataTable
        
        {
            
            Console.WriteLine(csv_file_path);
            string tmpFile = CSVreformat(csv_file_path);
            //string[] lines = System.IO.File.ReadAllLines(csv_file_path);

            //// Display the file contents by using a foreach loop.
            //System.Console.WriteLine("Contents of PAY_DOC.CSV = ");
            //foreach (string line in lines)
            //{
            //    // Use a tab to indent each line of the file.
            //    Console.WriteLine("\t" + line);
            //}

            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(tmpFile, Encoding.GetEncoding(1251)))
                {
                    csvReader.SetDelimiters(new string[] { "|" }); //Устанавливаем символ-разделитель
                    csvReader.HasFieldsEnclosedInQuotes = false;
                    string[] colFields = csvReader.ReadFields();
                    Console.WriteLine("Количество столбцов: {0}", colFields.Length);
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        //Console.WriteLine("Поле: {0}", column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Обработка null значений
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "NULL")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                    csvReader.Close();
                    File.Delete(tmpFile); //Удаляем временный файл
                }
            }
            
            catch (Exception ex)
            {
                Console.WriteLine("Exception: {0}", ex);
                return null;
            }
            return csvData;
        }


        public static void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData, string tn, string cs)
        {
            using (SqlConnection dbConnection = new SqlConnection(cs))
            {
                dbConnection.Open();
                using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                {
                    s.DestinationTableName = tn;
                    s.EnableStreaming = true;
                    s.BatchSize = 10000;
                    s.BulkCopyTimeout = 0;
                    s.NotifyAfter = 100;
                    s.SqlRowsCopied += delegate (object sender, SqlRowsCopiedEventArgs e)
                    {
                        Console.WriteLine(e.RowsCopied.ToString("#,##0") + " rows copied.");
                    };
                    foreach (var column in csvFileData.Columns)
                    {
                        s.ColumnMappings.Add(column.ToString(), column.ToString());

                        Console.WriteLine();
                    }
                    s.WriteToServer(csvFileData);
                }
                dbConnection.Close();
            }
        }

        public static string ANSItoUTF (string stringANSI)
        //Перекодирует строку из ANSItoUTF
        {
            // Create two different encodings.
            Encoding ansi = Encoding.GetEncoding(1251);
            Encoding unicode = Encoding.UTF8;

            // Perform the conversion from one encoding to the other.
            byte[] ansiBytes = ansi.GetBytes(stringANSI); //Encoding.Convert(ansi, unicode, unicodeBytes);

            // Convert the string into a byte array.
            byte[] unicodeBytes = Encoding.Convert(ansi, unicode, ansiBytes);            

            // Convert the new byte[] into a char[] and then into a string.
            char[] unicodeChars = new char[unicode.GetCharCount(unicodeBytes, 0, unicodeBytes.Length)];
            unicode.GetChars(unicodeBytes, 0, unicodeBytes.Length, unicodeChars, 0);
            string unicodeString = new string(unicodeChars);
            return unicodeString;            
        }

        public static string CSVreformat (string inCSV)
        //Приводит CSV файл к привычному виду
        {
            try
            {                
                using (StreamReader sr = new StreamReader(inCSV, Encoding.GetEncoding(1251)) )
                {
                    string line;
                    string tempFile = Path.GetTempFileName(); 
                    using (StreamWriter sw = new StreamWriter(tempFile, false ,Encoding.UTF8) )
                    {
                        line = sr.ReadLine();
                        line = Regex.Replace(line, "~", "");
                        line = Regex.Replace(line, @"\n+", " ");                      
                        sw.WriteLine(line);
                        while ((line = sr.ReadLine()) != null)
                        {
                            //if (line.StartsWith("eof"))
                            //    break;
                            if (Regex.IsMatch(line, @"^(?:eof\b)"))
                                break;

                            line = ANSItoUTF(line);
                            
                            if (line.StartsWith("~"))
                            {
                                line = Regex.Replace(line, "~", "");
                                line = Regex.Replace(line, @"\n+", " ");
                                sw.WriteLine();
                                sw.Write(line);
                            }
                            else
                            {
                                line = Regex.Replace(line, @"\n+", " ");
                                sw.Write(line);
                            }
                        }
                        sw.Close();
                        sr.Close();
                        //File.Replace(tempFile, inCSV, null);
                    }
                    return tempFile;
                }
                
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }


    }  
 



using System;
using ExcelDataReader;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Linq;

namespace ExcelLib
{
    //Класс открывает документ и считывает содержимое
    public class Excel: IDisposable
    {
        public delegate void CallBackHandler(bool result);
        private CallBackHandler resultListener;

        private Thread classThread;

        private List<List<object>> sheets = new List<List<object>>();

        /// <summary>
        /// Содержит считанные страницы документа
        /// </summary>
        public List<List<object>> Sheets {
            get
            {
                return sheets;
            }
        }

        //Асинхронный метод, передающий результат в вызывающий класс
        private bool CallBack(bool result)
        {
            return result;
        }

        /// <summary>
        /// Получение данных из Excel файла
        /// </summary>
        /// <param name="fileName">имя файла</param>
        /// <param name="resultListener">Функция обратного вызова</param>
        public void GetData(string fileName, CallBackHandler resultListener)
        {
            this.resultListener = resultListener;

            classThread = new Thread(new ParameterizedThreadStart(_getData));
            classThread.Start(fileName);
        }

        // Метод получения данных, вызываемый в отдельном потоке
        private void _getData(object fileName)
        {
            bool result = false;

            try
            {
                var file = new FileInfo((string)fileName);
                using (var stream = new FileStream((string)fileName, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader = null;
                    if (file.Extension == ".xls")
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (file.Extension == ".xlsx")
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }

                    if (reader != null)
                    {
                        int sheet = 0;
                        do
                        {
                            sheets.Add(new List<object>() { "Sheet" + sheet.ToString() });

                            while (reader.Read())
                            {
                                object[] tmp = new object[reader.FieldCount];

                                for (int i = 0; i < reader.FieldCount; i++)
                                    tmp[i] = reader.GetValue(i);

                                tmp = tmp.Where(item => item != null).ToArray();

                                if (tmp.Length > 0)
                                {
                                    sheets.Add(tmp.ToList());

                                    if (!result)
                                        result = true;
                                }
                            }

                            sheet++;
                        }
                        while (reader.NextResult());

                        reader.Dispose();
                        reader = null;

                        stream.Dispose();
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }

            resultListener?.Invoke(result);
        }

        private Array ResizeArray(Array arr, int[] newSizes)
        {
            if (newSizes.Length != arr.Rank)
                throw new ArgumentException("arr must have the same number of dimensions " +
                                            "as there are elements in newSizes", "newSizes");

            var temp = Array.CreateInstance(arr.GetType().GetElementType(), newSizes);
            int length = arr.Length <= temp.Length ? arr.Length : temp.Length;
            Array.ConstrainedCopy(arr, 0, temp, 0, length);

            return temp;
        }

        //Деструктор
        public void Dispose()
        {
            classThread.Abort();
            classThread = null;

            resultListener = null;

            sheets = null;

            GC.GetTotalMemory(true);
        }
    }
}

/*
    The DB-Matcher is an easy-to-use console application written in C# and based on the .NET framework. 
    The DB-Matcher can merge two databases in Excel format (*.xlsx, *.xls). 
    It follows the following algorithms in order of importance: Levenshtein distance, Hamming distance, Jaccard index. 
    The DB-Matcher takes you by the hand at all times and guides you through the data matching process    

    Copyright (C) 2024  Carl Öttinger (Carl Oettinger)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as published
    by the Free Software Foundation, either version 3 of the License, or
    any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.

    You should have received a copy of the GNU Affero General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

    You can contact me in the following ways:
        EMail: oettinger.carl@web.de, big-programming@web.de
 */


using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;


namespace RandomExcelFiller
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            /*
            string checksumConverted = "";
            Console.Write("Path: ");
            string path = Console.ReadLine();
            using (FileStream fileStream = File.OpenRead(path))
            {
                SHA256Managed sha = new SHA256Managed();
                byte[] checksum = sha.ComputeHash(fileStream);
                checksumConverted = BitConverter.ToString(checksum).Replace("-", String.Empty);
            }
            Console.WriteLine(checksumConverted);*/


            int row = 10000;
            int col = 2;

            Random random = new Random();

            Console.Write("Eingabe des Dateipfades: ");
            string path = Convert.ToString(Console.ReadLine()); //Pfad der Excel-Datei durch Konsoleneinabe

            IWorkbook workbook;

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite)) //Lese-/Schreibzugriff
            {
                workbook = new XSSFWorkbook(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);

            for (int i = 0; i < row; i++)
            {
                IRow irow = sheet.CreateRow(i);

                for (int j = 0; j < col; j++)
                {
                    int cellValueHold = random.Next(1000, 10000);
                    //Console.WriteLine($"Row: {i} | Col: {j} | Value: {cellValueHold}");

                    ICell icell = irow.CreateCell(j);
                    icell.SetCellValue(cellValueHold);

                    using (FileStream fstream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                    {
                        workbook.Write(fstream);
                    }
                }
                Console.SetCursorPosition(0, Console.WindowHeight - 1);
                Console.Write(new string(' ', Console.WindowWidth));
                Console.SetCursorPosition(0, Console.WindowHeight - 1);
                int percentage = Convert.ToInt32((Convert.ToDouble(i + 1) / (row + 1)) * 100);
                double tprogress = (i * (Console.WindowWidth - 6) / row);
                Console.Write($"[{new string('#', Convert.ToInt32(tprogress))}{new string(' ', (Console.WindowWidth - 6) - Convert.ToInt32(tprogress))}]{percentage}%");
            }
        }
    }
}

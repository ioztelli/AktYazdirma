using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocXEditor;
using Excel;

namespace Akt_Yazdirma
{
    class Program
    {
        static void Main(string[] args)
        {
           DataSet result= Oku("Akt_Excel_Data.xls");
           Console.WriteLine("Excel Okunuyor...");
           foreach (DataRow rows in result.Tables[0].Rows)
           {
               if (rows[0] != DBNull.Value)
               {
                   string line = rows[0].ToString();
                   double len = Convert.ToDouble(rows[1]);
                   int dia = Convert.ToInt16(rows[2]);
                   int en_trench = Convert.ToInt16(rows[3]);
                   string en2_kum = rows[4].ToString();
                   int kum = Convert.ToInt16(rows[5]);

                   Console.WriteLine(line);
                   DocYazdir("Boru_Hatti_Akt.docx", line, dia, len, en_trench, kum,en2_kum);

               }
           }
            //DocYazdir("Boru_Hatti_Akt.docx","1-2",200,20.454,70,30);
            Console.ReadLine();
            
            
        }

        private static void DocYazdir(string mainfile,string line,int dia,double len,double en,double kum,string ken)
        {
            //Console.WriteLine("--Akt Yazdirma--");
            ReplacementList replacementList = new ReplacementList();
            replacementList.Add(new ReplaceItem("%Line%",line));
            replacementList.Add(new ReplaceItem("%Dia%", "Ø"+dia.ToString()));
            replacementList.Add(new ReplaceItem("%L%",len.ToString()));
            replacementList.Add(new ReplaceItem("%En%", en.ToString())); //En Cm olacak
            replacementList.Add(new ReplaceItem("%Kum%",kum.ToString()));
            replacementList.Add(new ReplaceItem("%K%", ken));            //(((en-Convert.ToDouble(dia))/10)/2).ToString()));

            DocXEditor.DocXEditor dox=new DocXEditor.DocXEditor(mainfile);
            dox.ReplaceList = replacementList;
            if (!dox.ReplaceContent(line+".docx", true))
            {
                Console.WriteLine("ERROR: " + dox.LastError);
            }
            else
            {
                //Console.WriteLine("No errors");
            }
        }
        public static DataSet Oku(string dosyaAdi)
        {

            FileStream stream = File.Open(@dosyaAdi, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            return result;



        }
    }
}

//using EPPlus_Library;
using EPPlus_Library;


using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace FastTesting
{
    class Program
    {
        private const string TestFolderPath = @"E:\TestFolder\";
        static void Main(string[] args)
        {


            Logger logger = Logger.getInstance;
            
            logger.setLogPathandFile(TestFolderPath, "Error.log");
            List<Backlog> mainTable = GetBacklog();
            Backlog[] t = new Backlog[mainTable.Count];
            mainTable.CopyTo(t);

            mainTable[0].commentarios = "Change";

            List<Backlog> ossFileItems = GetOSSFileItems();



            foreach(Backlog ossFile in ossFileItems)
            {

                Backlog match = mainTable.Where(
                    x => x.hpOrderNumber == ossFile.hpOrderNumber &&
                    x.itemNumber == ossFile.itemNumber &&
                    x.productNumber == ossFile.productNumber).FirstOrDefault();

                if(match != null)
                {
                   

                    ossFile.commentarios = match.commentarios;
                }
            };


            
          
          
            //NpoiExcelCreator npoiExcelCreator = new NpoiExcelCreator();
            //npoiExcelCreator.createSheet("TestSheet");
            //npoiExcelCreator.setSheet(0);
            //npoiExcelCreator.createRowsInstance(table.Count());
            //npoiExcelCreator.WriteList_To_Excel(0,0,0,table[0].Count()-1, table,1);
            //npoiExcelCreator.saveFile(TestFolderPath + "Max2.xlsx");
            
            //EPPlusCreator ePPlus = new EPPlusCreator();
            //ePPlus.CreateSheet("TestSheet");
            //ePPlus.SetSheet("TestSheet");
            //ePPlus.WriteInfo(table);
            //ePPlus.SaveFile(TestFolderPath + "EPPlus.xlsx",false);
           
        }

        static List<Backlog> GetBacklog()
        {
            List<Backlog> backlog = new List<Backlog>();
            backlog.Add(new Backlog { hpOrderNumber = "1", itemNumber = "1", productNumber = "1", commentarios = "" });
            backlog.Add(new Backlog { hpOrderNumber = "2", itemNumber = "2", productNumber = "2", commentarios = "2" });
            backlog.Add(new Backlog { hpOrderNumber = "3", itemNumber = "3", productNumber = "3", commentarios = "3" });
            backlog.Add(new Backlog { hpOrderNumber = "4", itemNumber = "4", productNumber = "4", commentarios = "4" });
            backlog.Add(new Backlog { hpOrderNumber = "5", itemNumber = "5", productNumber = "5", commentarios = "5" });
            return backlog;
        }

        static List<Backlog> GetOSSFileItems()
        {
            List<Backlog> ossFileItems = new List<Backlog>();
            ossFileItems.Add(new Backlog { hpOrderNumber = "0", itemNumber = "0", productNumber = "0", commentarios = "" });
            ossFileItems.Add(new Backlog { hpOrderNumber = "2", itemNumber = "2", productNumber = "2" , commentarios = "" });
            ossFileItems.Add(new Backlog { hpOrderNumber = "3", itemNumber = "3", productNumber = "3" , commentarios = "" });
            ossFileItems.Add(new Backlog { hpOrderNumber = "4", itemNumber = "4", productNumber = "4" , commentarios = "" });
            ossFileItems.Add(new Backlog { hpOrderNumber = "5", itemNumber = "5", productNumber = "5" , commentarios = "" });
            return ossFileItems;
        }
    }
}

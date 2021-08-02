using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GerarAmostragem
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Executando...");
            Console.WriteLine("----------------------------------");
            Console.WriteLine("Lendo documento...");

            WorkBook workbook = WorkBook.Load("d:\\documento.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();

            var dataSet = workbook.ToDataSet();
            var qtdLinhas = dataSet.Tables[0].Rows.Count;

            var porcentagem = Math.Round(qtdLinhas * 0.10);

            Random randNum = new Random();

            List<int> ListaNumeroRandom = new List<int>();

            for (int i = 0; i < porcentagem; i++)
            {
                var numero = randNum.Next(qtdLinhas);
                if (!ListaNumeroRandom.Contains(numero))
                {
                    ListaNumeroRandom.Add(numero);
                }
                else {
                    i--;
                }
            }

            int contador = 1;
            foreach (var item in ListaNumeroRandom)
            {
                sheet[$"X{item}"].Value = "Selecionado";
                Console.WriteLine($"{contador}/{ListaNumeroRandom.Count()}");
                contador++;
            }

            workbook.Save();
            Console.WriteLine("----------------------------------");
            Console.WriteLine(">>>> PRONTO!!! <<<<");
        }
    }
}

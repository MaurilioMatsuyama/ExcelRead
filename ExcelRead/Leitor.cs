using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using Npgsql;
using System.Linq;

namespace ExcelRead
{
    public class Leitor
    {
        //public static void Main(string[] args)
        //{
        //    string fileName = "C:\\Município.xlsx";
        //    var arquivo = new XLWorkbook(fileName);
        //    var sheet = arquivo.Worksheet(1);

        //    List<Registro> lista = new List<Registro>();
        //    List<Registro> listaerro = new List<Registro>();

        //    var rows = sheet.Rows();

        //    foreach (var item in rows)
        //    {
        //        lista.Add(new Registro()
        //        {
        //            Codigo = item.Cell(1).Value.ToString(),
        //            Descricao = item.Cell(2).Value.ToString().Substring(0, item.Cell(2).Value.ToString().IndexOf("(")).Trim()
        //        });
        //    }

        //    //Conexão
        //    using (NpgsqlConnection conn = new NpgsqlConnection("Server=192.168.0.21;Port=5432;Database=cep;User Id=postgres;Password=Bredas;"))
        //    {
        //        NpgsqlTransaction transaction = null;
        //        try
        //        {
        //            conn.Open();
        //            transaction = conn.BeginTransaction();

        //            foreach (var item in lista)
        //            {
        //                //Query
        //                NpgsqlCommand cmd = new NpgsqlCommand("UPDATE cidade SET ibge = '" + item.Codigo + "' WHERE nome = '" + item.Descricao.Replace("'", "''") + "'", conn, transaction);
        //                Console.WriteLine(item.Codigo + " - " + item.Descricao);
        //                int result = cmd.ExecuteNonQuery();
        //                if (result == 0)
        //                    listaerro.Add(item);
        //            }



        //            transaction.Commit();

        //            Console.Clear();

        //            foreach (var item in listaerro)
        //            {
        //                Console.WriteLine("ERRO: " + item.Codigo + " - " + item.Descricao);
        //            }
        //        }
        //        catch (Exception ex)
        //        {

        //            transaction.Rollback();
        //            throw;
        //        }
        //    }           

        //    foreach(var item in lista)
        //    {
        //        Console.WriteLine(item.Codigo);
        //        Console.WriteLine(item.Descricao);

        //        Console.ReadKey();
        //    }      
        //}

        public static void Main(string[] args)
        {
            string fileName = "C:\\País.xlsx";
            var arquivo = new XLWorkbook(fileName);
            var sheet = arquivo.Worksheet(1);

            List<Registro> lista = new List<Registro>();
            List<Registro> listaerro = new List<Registro>();

            var rows = sheet.Rows();

            foreach (var item in rows)
            {
                lista.Add(new Registro()
                {
                    Codigo = item.Cell(1).Value.ToString(),
                    Descricao = item.Cell(2).Value.ToString()
                });
            }

            //Conexão
            using (NpgsqlConnection conn = new NpgsqlConnection("Server=192.168.0.21;Port=5432;Database=cep;User Id=postgres;Password=Bredas;"))
            {
                NpgsqlTransaction transaction = null;
                try
                {
                    conn.Open();
                    transaction = conn.BeginTransaction();

                    List<int> paisesduplicados = new List<int>();

                    foreach (var item in lista)
                    {
                        //Query
                        if(paisesduplicados.Contains(int.Parse(item.Codigo)))
                        {
                            Console.WriteLine("duplicado: " + item.Codigo + " Descricao: " + item.Descricao);
                        }
                        NpgsqlCommand cmd = new NpgsqlCommand("INSERT INTO pais VALUES (nextval('pais_id_seq'), " + item.Codigo + ", '" + item.Descricao.Replace("'", "''") + "') ", conn, transaction);
                        Console.WriteLine(item.Codigo + " - " + item.Descricao);
                        int result = cmd.ExecuteNonQuery();
                        if (result == 0)
                            listaerro.Add(item);

                        paisesduplicados.Add(int.Parse(item.Codigo));
                    }

                    transaction.Commit();

                    Console.Clear();

                    foreach (var item in listaerro)
                    {
                        Console.WriteLine("ERRO: " + item.Codigo + " - " + item.Descricao);
                    }
                }
                catch (Exception ex)
                {

                    transaction.Rollback();
                    throw;
                }
            }

            foreach (var item in lista)
            {
                Console.WriteLine(item.Codigo);
                Console.WriteLine(item.Descricao);

                Console.ReadKey();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Bolinha.Modelos;

namespace Bolinha.DAL
{
    //*********************************************** OK *******************************************************************************
    public class ConexaoBD
    {
        //******************************************************************************************************************************
        // private SqlConnection con = new SqlConnection(@"Data Source=printserver;Initial Catalog=bolinhas;                                                                                  user=sa;password=D@teujeito");
        private SqlConnection con = new SqlConnection(@"Data Source=192.168.1.254;Initial Catalog=dbBolinhas;                                                                                  user=jordao;password=jordao!0192");

        // private SqlConnection con = new SqlConnection(@"Data Source=192.168.1.40;Initial Catalog=bolinhas2;user=sa;password=jordao");
        // private SqlConnection con = new SqlConnection(@"Data Source=monitoramento;Initial Catalog=bolinhas2;user=sa;password=jordao");

        //******************************************************************************************************************************
        //Procedure para Conexão ao Banco
        public void Abreconexao()
        {
            try
            {
                con.Open();
            }
            catch
            {
                //Tirei System.Windows.Forms.MessageBox.Show("4 - A conexão não foi estabelecida, procure o suporte!!!");
            }
            finally
            {
                //con.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure Para Retornar O Maior Valor Cadastrado
        public static string StringDeConexao
        {
            get
            {
              // return @"Data Source=printserver;Initial Catalog=bolinhas;                                                                                                                  user=sa;password=D@teujeito";
              return @"Data Source=192.168.1.254;Initial Catalog=dbBolinhas;                                                                                                                  user=jordao;password=jordao!0192";

               // return @"Data Source=192.168.1.40;Initial Catalog=bolinhas2;user=sa;password=jordao";
               // return @"Data Source=monitoramento;Initial Catalog=bolinhas2;user=sa;password=jordao";
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.SqlClient;

namespace Bolinha.DAL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             SqlConnection cn = new SqlConnection();
             try
             {
                 cn.ConnectionString = ConexaoBD.StringDeConexao;
                 cn.Open();
                 MessageBox.Show("Conectado!!!!");
             }
             catch
             {
                 MessageBox.Show("Erro!!!");
             }
        }
    }
}

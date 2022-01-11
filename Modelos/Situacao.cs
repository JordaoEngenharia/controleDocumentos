using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bolinha.Modelos
{
    //***************************************************** OK *************************************************************************
    public class Situacao
    {
        private int CodSit;
        private string DescricaoSit;

        public Situacao() { }
        public Situacao(int CodSit, string DescricaoSit)
        {
            this.CodSit = CodSit;
            this.DescricaoSit = DescricaoSit;
        }
        public int Codigo
        {
            get { return CodSit; }
            set { CodSit = value; }
        }

        public string Descricao
        {
            get { return DescricaoSit; }
            set { DescricaoSit = value; }
        }
    }

    public static class Global
    {
        private static int m_CodigoEnquete = 0;
        public static int CodigoEnquete
        {
            get { return m_CodigoEnquete; }
            set { m_CodigoEnquete = value; }
        }
        private static string _Status = "";
        public static string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }
    }
}

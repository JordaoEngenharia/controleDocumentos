using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bolinha.Modelos
{
    //************************************************* OK *****************************************************************************
    public class Mensagens
    {
        protected int codigoMsg;
        protected string assunto, conteudo;

        public Mensagens() { }

        public Mensagens(int codigoMsg, string assunto, string conteudo)
        {
            this.codigoMsg = codigoMsg;
            this.assunto = assunto;
            this.conteudo = conteudo;
        }
        public int CodigoMsg
        {
            get { return codigoMsg; }
            set { codigoMsg = value; }
        }
        public string Assunto
        {
            get { return assunto; }
            set { assunto = value; }
        }
        public string Conteudo
        {
            get { return conteudo; }
            set { conteudo = value; }
        }
    }
}

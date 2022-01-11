using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bolinha.Modelos
{
    //*********************************************************** OK *******************************************************************
    public class Usuarios
    {
        protected int codigoUsr;
        protected string nome, status, setor, telefone, ramal, loginusr, pendencia, observacao;
        protected DateTime nascimento, formatura, admissao, folgap;
        protected DateTime entseg, entter, entqua, entqui, entsex, saiseg, saiter, saiqua, saiqui, saisex;
        protected byte[] foto, fotao;

        public Usuarios() { }

        public Usuarios(int codigoUsr, string nome, string status, DateTime nascimento, DateTime formatura, DateTime admissao, DateTime folgap, string setor, byte[] foto, byte[] fotao, string pendencia, string telefone, string ramal, string loginusr, string observacao, DateTime entseg, DateTime entter, DateTime entqua, DateTime entqui, DateTime entsex, DateTime saiseg, DateTime saiter, DateTime saiqua, DateTime saiqui, DateTime saisex)
        {
            this.codigoUsr = codigoUsr;
            this.nome = nome;
            this.status = status;
            this.nascimento = nascimento;
            this.formatura = formatura;
            this.admissao = admissao;
            this.folgap = folgap;
            this.setor = setor;
            this.pendencia = pendencia;
            this.telefone = telefone;
            this.ramal = ramal;
            this.loginusr = loginusr;
            this.observacao = observacao;
            this.foto = foto;
            this.fotao = fotao;
            this.entseg = entseg;
            this.entter = entter;
            this.entqua = entqua;
            this.entqui = entqui;
            this.entsex = entsex;
            this.saiseg = saiseg;
            this.saiter = saiter;
            this.saiqua = saiqua;
            this.saiqui = saiqui;
            this.saisex = saisex;
        }
        public int CodigoUsr
        {
            get { return codigoUsr; }
            set { codigoUsr = value; }
        }
        public string Nome
        {
            get { return nome; }
            set { nome = value; }
        }
        public string Status
        {
            get { return status; }
            set { status = value; }
        }
        public DateTime Nascimento
        {
            get { return nascimento; }
            set { nascimento = value; }
        }
        public DateTime Formatura
        {
            get { return formatura; }
            set { formatura = value; }
        }
        public DateTime Admissao
        {
            get { return admissao; }
            set { admissao = value; }
        }
        public DateTime FolgaP
        {
            get { return folgap; }
            set { folgap = value; }
        }
        public string Setor
        {
            get { return setor; }
            set { setor = value; }
        }
        public byte[] Foto
        {
            get { return foto; }
            set { foto = value; }
        }
        public byte[] Fotao
        {
            get { return fotao; }
            set { fotao = value; }
        }
        public string Pendencia
        {
            get { return pendencia; }
            set { pendencia = value; }
        }
        public string Telefone
        {
            get { return telefone; }
            set { telefone = value; }
        }
        public string Ramal
        {
            get { return ramal; }
            set { ramal = value; }
        }
        public string LoginUsr
        {
            get { return loginusr; }
            set { loginusr = value; }
        }
        public string Observacao
        {
            get { return observacao; }
            set { observacao = value; }
        }
        public DateTime EntSeg
        {
            get { return entseg; }
            set { entseg = value; }
        }

        public DateTime EntTer
        {
            get { return entter; }
            set { entter = value; }
        }

        public DateTime EntQua
        {
            get { return entqua; }
            set { entqua = value; }
        }

        public DateTime EntQui
        {
            get { return entqui; }
            set { entqui = value; }
        }

        public DateTime EntSex
        {
            get { return entsex; }
            set { entsex = value; }
        }


        public DateTime SaiSeg
        {
            get { return saiseg; }
            set { saiseg = value; }
        }

        public DateTime SaiTer
        {
            get { return saiter; }
            set { saiter = value; }
        }

        public DateTime SaiQua
        {
            get { return saiqua; }
            set { saiqua = value; }
        }

        public DateTime SaiQui
        {
            get { return saiqui; }
            set { saiqui = value; }
        }

        public DateTime SaiSex
        {
            get { return saisex; }
            set { saisex = value; }
        }
    }
}

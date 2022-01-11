using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bolinha.Modelos;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace Bolinha.DAL
{
    public class Acesso
    {
        //******************************************************************************************************************************
        //Procedure para cadastrar um Usuário
        public void CadastraUsuario(Usuarios u)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;

            try
            {
                SqlCommand CmdIns = new SqlCommand();
                CmdIns.Connection = cn;
                CmdIns.CommandType = CommandType.StoredProcedure;
                CmdIns.CommandText = "InserirUsuarios";

                //   Inserindo Parâmetros
                SqlParameter pcodigoUsr = new SqlParameter("@CodigoUsr", SqlDbType.Int);
                pcodigoUsr.Value = u.CodigoUsr;
                CmdIns.Parameters.Add(pcodigoUsr);

                SqlParameter pnomeUsr = new SqlParameter("@NomeUsr", SqlDbType.NVarChar, 50);
                pnomeUsr.Value = u.Nome;
                CmdIns.Parameters.Add(pnomeUsr);

                SqlParameter pstatus = new SqlParameter("@Status", SqlDbType.NVarChar, 30);
                pstatus.Value = u.Status;
                CmdIns.Parameters.Add(pstatus);

                SqlParameter pfoto = new SqlParameter("@foto", SqlDbType.Image);
                pfoto.Value = u.Foto;
                CmdIns.Parameters.Add(pfoto);

                SqlParameter pfotao = new SqlParameter("@fotao", SqlDbType.Image);
                pfotao.Value = u.Fotao;
                CmdIns.Parameters.Add(pfotao);

                SqlParameter pnascimento = new SqlParameter("@nascimento", SqlDbType.DateTime);
                pnascimento.Value = u.Nascimento;
                CmdIns.Parameters.Add(pnascimento);

                SqlParameter padmissao = new SqlParameter("@admissao", SqlDbType.DateTime);
                padmissao.Value = u.Admissao;
                CmdIns.Parameters.Add(padmissao);

                SqlParameter pfolgap = new SqlParameter("@folgap", SqlDbType.DateTime);
                pfolgap.Value = u.FolgaP;
                CmdIns.Parameters.Add(pfolgap);

                SqlParameter pformatura = new SqlParameter("@formatura", SqlDbType.DateTime);
                pformatura.Value = u.Formatura;
                CmdIns.Parameters.Add(pformatura);

                SqlParameter psetor = new SqlParameter("@setor", SqlDbType.NVarChar, 30);
                psetor.Value = u.Setor;
                CmdIns.Parameters.Add(psetor);

                SqlParameter ppendencia = new SqlParameter("@pendencia", SqlDbType.NVarChar, 50);
                ppendencia.Value = u.Pendencia;
                CmdIns.Parameters.Add(ppendencia);

                SqlParameter ptelefone = new SqlParameter("@telefone", SqlDbType.NVarChar, 20);
                ptelefone.Value = u.Telefone;
                CmdIns.Parameters.Add(ptelefone);

                SqlParameter pramal = new SqlParameter("@ramal", SqlDbType.NVarChar, 10);
                pramal.Value = u.Ramal;
                CmdIns.Parameters.Add(pramal);

                SqlParameter ploginusr = new SqlParameter("@loginusr", SqlDbType.NVarChar, 50);
                ploginusr.Value = u.LoginUsr;
                CmdIns.Parameters.Add(ploginusr);

                SqlParameter pobservacao = new SqlParameter("@observacao", SqlDbType.NVarChar, 250);
                pobservacao.Value = u.Observacao;
                CmdIns.Parameters.Add(pobservacao);

                SqlParameter pEntSeg = new SqlParameter("@EntSeg", SqlDbType.DateTime);
                pEntSeg.Value = u.EntSeg;
                CmdIns.Parameters.Add(pEntSeg);

                SqlParameter pEntTer = new SqlParameter("@EntTer", SqlDbType.DateTime);
                pEntTer.Value = u.EntTer;
                CmdIns.Parameters.Add(pEntTer);

                SqlParameter pEntQua = new SqlParameter("@EntQua", SqlDbType.DateTime);
                pEntQua.Value = u.EntQua;
                CmdIns.Parameters.Add(pEntQua);

                SqlParameter pEntQui = new SqlParameter("@EntQui", SqlDbType.DateTime);
                pEntQui.Value = u.EntQui;
                CmdIns.Parameters.Add(pEntQui);

                SqlParameter pEntSex = new SqlParameter("@EntSex", SqlDbType.DateTime);
                pEntSex.Value = u.EntSex;
                CmdIns.Parameters.Add(pEntSex);

                SqlParameter pSaiSeg = new SqlParameter("@SaiSeg", SqlDbType.DateTime);
                pSaiSeg.Value = u.SaiSeg;
                CmdIns.Parameters.Add(pSaiSeg);

                SqlParameter pSaiTer = new SqlParameter("@SaiTer", SqlDbType.DateTime);
                pSaiTer.Value = u.SaiTer;
                CmdIns.Parameters.Add(pSaiTer);

                SqlParameter pSaiQua = new SqlParameter("@SaiQua", SqlDbType.DateTime);
                pSaiQua.Value = u.SaiQua;
                CmdIns.Parameters.Add(pSaiQua);

                SqlParameter pSaiQui = new SqlParameter("@SaiQui", SqlDbType.DateTime);
                pSaiQui.Value = u.SaiQui;
                CmdIns.Parameters.Add(pSaiQui);

                SqlParameter pSaiSex = new SqlParameter("@SaiSex", SqlDbType.DateTime);
                pSaiSex.Value = u.SaiSex;
                CmdIns.Parameters.Add(pSaiSex);

                // Executando a procedure
                cn.Open();
                CmdIns.ExecuteNonQuery();
                System.Windows.Forms.MessageBox.Show("Usuário Inserido com Sucesso!!!");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                cn.Close();
            }
        }
        //******************************************************************************************************************************
        //Procedure para cadastrar uma Mensagem
        public void CadastraMensagem(Mensagens m)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;

            try
            {
                SqlCommand CmdIns = new SqlCommand();
                CmdIns.Connection = cn;
                CmdIns.CommandType = CommandType.StoredProcedure;
                CmdIns.CommandText = "InserirMensagens";

                //Inserindo Parâmetros
                SqlParameter pcodigoMsg = new SqlParameter("@codigoMsg", SqlDbType.Int);
                pcodigoMsg.Direction = ParameterDirection.Output;
                CmdIns.Parameters.Add(pcodigoMsg);

                SqlParameter pAssunto = new SqlParameter("@Assunto", SqlDbType.NVarChar, 50);
                pAssunto.Value = m.Assunto;
                CmdIns.Parameters.Add(pAssunto);

                SqlParameter pConteudo = new SqlParameter("@Conteudo", SqlDbType.NVarChar, 300);
                pConteudo.Value = m.Conteudo;
                CmdIns.Parameters.Add(pConteudo);

                //Executando a procedure
                cn.Open();
                CmdIns.ExecuteNonQuery();
                System.Windows.Forms.MessageBox.Show("Mensagem Inserida com Sucesso!!!");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                cn.Close();
            }
        }
        //******************************************************************************************************************************
        //Procedure para Atualizar Usuários

        public void AtualizaUsuario(Usuarios u)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;

            try
            {
                SqlCommand CmdIns = new SqlCommand();
                CmdIns.Connection = cn;
                CmdIns.CommandType = CommandType.StoredProcedure;
                CmdIns.CommandText = "AtualizaUsuarios";

                //Inserindo Parâmetros
                SqlParameter pcodigoUsr = new SqlParameter("@CodigoUsr", SqlDbType.Int);
                pcodigoUsr.Value = u.CodigoUsr;
                CmdIns.Parameters.Add(pcodigoUsr);

                SqlParameter pnomeUsr = new SqlParameter("@NomeUsr", SqlDbType.NVarChar, 50);
                pnomeUsr.Value = u.Nome;
                CmdIns.Parameters.Add(pnomeUsr);

                SqlParameter pstatus = new SqlParameter("@Status", SqlDbType.NVarChar, 30);
                pstatus.Value = u.Status;
                CmdIns.Parameters.Add(pstatus);

                SqlParameter pnascimento = new SqlParameter("@Nascimento", SqlDbType.Date);
                pnascimento.Value = u.Nascimento;
                CmdIns.Parameters.Add(pnascimento);

                SqlParameter pformatura = new SqlParameter("@Formatura", SqlDbType.Date);
                pformatura.Value = u.Formatura;
                CmdIns.Parameters.Add(pformatura);

                SqlParameter padmissao = new SqlParameter("@Admissao", SqlDbType.Date);
                padmissao.Value = u.Admissao;
                CmdIns.Parameters.Add(padmissao);

                SqlParameter pfolgap = new SqlParameter("@FolgaP", SqlDbType.Date);
                pfolgap.Value = u.FolgaP;
                CmdIns.Parameters.Add(pfolgap);

                SqlParameter psetor = new SqlParameter("@Setor", SqlDbType.NVarChar, 30);
                psetor.Value = u.Setor;
                CmdIns.Parameters.Add(psetor);

                SqlParameter ppendencia = new SqlParameter("@Pendencia", SqlDbType.NVarChar, 50);
                ppendencia.Value = u.Pendencia;
                CmdIns.Parameters.Add(ppendencia);

                SqlParameter ptelefone = new SqlParameter("@Telefone", SqlDbType.NVarChar, 20);
                ptelefone.Value = u.Telefone;
                CmdIns.Parameters.Add(ptelefone);

                SqlParameter pramal = new SqlParameter("@Ramal", SqlDbType.NVarChar, 10);
                pramal.Value = u.Ramal;
                CmdIns.Parameters.Add(pramal);

                SqlParameter ploginusr = new SqlParameter("@LoginUsr", SqlDbType.NVarChar, 50);
                ploginusr.Value = u.LoginUsr;
                CmdIns.Parameters.Add(ploginusr);

                SqlParameter pobservacao = new SqlParameter("@Observacao", SqlDbType.NVarChar, 250);
                pobservacao.Value = u.Observacao;
                CmdIns.Parameters.Add(pobservacao);

                SqlParameter pfoto = new SqlParameter("@Foto", SqlDbType.Image);
                pfoto.Value = u.Foto;
                CmdIns.Parameters.Add(pfoto);

                SqlParameter pfotao = new SqlParameter("@Fotao", SqlDbType.Image);
                pfotao.Value = u.Fotao;
                CmdIns.Parameters.Add(pfotao);

                SqlParameter pEntSeg = new SqlParameter("@EntSeg", SqlDbType.DateTime);
                pEntSeg.Value = u.EntSeg;
                CmdIns.Parameters.Add(pEntSeg);

                SqlParameter pEntTer = new SqlParameter("@EntTer", SqlDbType.DateTime);
                pEntTer.Value = u.EntTer;
                CmdIns.Parameters.Add(pEntTer);

                SqlParameter pEntQua = new SqlParameter("@EntQua", SqlDbType.DateTime);
                pEntQua.Value = u.EntQua;
                CmdIns.Parameters.Add(pEntQua);

                SqlParameter pEntQui = new SqlParameter("@EntQui", SqlDbType.DateTime);
                pEntQui.Value = u.EntQui;
                CmdIns.Parameters.Add(pEntQui);

                SqlParameter pEntSex = new SqlParameter("@EntSex", SqlDbType.DateTime);
                pEntSex.Value = u.EntSex;
                CmdIns.Parameters.Add(pEntSex);

                SqlParameter pSaiSeg = new SqlParameter("@SaiSeg", SqlDbType.DateTime);
                pSaiSeg.Value = u.SaiSeg;
                CmdIns.Parameters.Add(pSaiSeg);

                SqlParameter pSaiTer = new SqlParameter("@SaiTer", SqlDbType.DateTime);
                pSaiTer.Value = u.SaiTer;
                CmdIns.Parameters.Add(pSaiTer);

                SqlParameter pSaiQua = new SqlParameter("@SaiQua", SqlDbType.DateTime);
                pSaiQua.Value = u.SaiQua;
                CmdIns.Parameters.Add(pSaiQua);

                SqlParameter pSaiQui = new SqlParameter("@SaiQui", SqlDbType.DateTime);
                pSaiQui.Value = u.SaiQui;
                CmdIns.Parameters.Add(pSaiQui);

                SqlParameter pSaiSex = new SqlParameter("@SaiSex", SqlDbType.DateTime);
                pSaiSex.Value = u.SaiSex;
                CmdIns.Parameters.Add(pSaiSex);

                //Executando a procedure
                cn.Open();
                CmdIns.ExecuteNonQuery();
                System.Windows.Forms.MessageBox.Show("Usuário Alterado com Sucesso!!!");

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                cn.Close();
            }

        }

        //******************************************************************************************************************************
        // Procedure para Atualizar Mensagens

        public void AtualizaMensagens(Mensagens m)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;

            try
            {
                SqlCommand CmdIns = new SqlCommand();
                CmdIns.Connection = cn;
                CmdIns.CommandType = CommandType.StoredProcedure;
                CmdIns.CommandText = "AtualizaMensagens";

                //Inserindo Parâmetros
                SqlParameter pcodigoMsg = new SqlParameter("@CodigoMsg", SqlDbType.Int);
                pcodigoMsg.Value = m.CodigoMsg;
                CmdIns.Parameters.Add(pcodigoMsg);

                SqlParameter pAssunto = new SqlParameter("@Assunto", SqlDbType.NVarChar, 50);
                pAssunto.Value = m.Assunto;
                CmdIns.Parameters.Add(pAssunto);

                SqlParameter pConteudo = new SqlParameter("@Conteudo", SqlDbType.NVarChar, 300);
                pConteudo.Value = m.Conteudo;
                CmdIns.Parameters.Add(pConteudo);

                //Executando a procedure
                cn.Open();
                CmdIns.ExecuteNonQuery();
                System.Windows.Forms.MessageBox.Show("Mensagem Atualizada com Sucesso!!!");

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure para obter todos os Usuarios

        public void getUsuarios()
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;

            cn.Open();
            try
            {
                SqlCommand Cmdget = new SqlCommand();
                Cmdget.Connection = cn;
                Cmdget.CommandType = CommandType.StoredProcedure;
                Cmdget.CommandText = "getUsuarios";

                //Executando a procedure
                cn.Open();
                Cmdget.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        ////Procedure para preencher Data Grid Relatorio
        //public void PopulaGridRel(System.Windows.Forms.DataGridView grid, DateTime D_Ini, DateTime D_Fin, string u)
        //{
        //    SqlConnection cn = new SqlConnection();
        //    cn.ConnectionString = ConexaoBD.StringDeConexao;
        //    cn.Open();

        //    string Sql = "SELECT HoraLog, StatusLog FROM RegLogBolinas where  HoraLog>= " + D_Ini + " and HoraLog<= " + D_Fin + " and UsrLog in " + "(SELECT CodigoUsr FROM Usuario as U where U.NomeUsr like '%" + u + "%)";
            
        //    SqlCommand cmd = new SqlCommand(Sql, cn);

        //    SqlDataAdapter adpt = new SqlDataAdapter(cmd);

        //    DataTable dt = new DataTable();

        //    adpt.Fill(dt);
        //    grid.DataSource = dt;
            
        //}



        //******************************************************************************************************************************
        //Procedure para preencher Data List Usuário
        // Criando em: 19/01/2012 Por: Filé
        
        public void PopulaListUsuario(ComboBox list)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT NomeUsr FROM Usuario order by NomeUsr";    // where NomeUsr like '%" + u + "%'";

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            adpt.Fill(dt);
            
            list.DisplayMember = "NomeUsr";
            list.ValueMember = "NomeUsr";
            list.DataSource = dt;
            cn.Close();
        }


        //******************************************************************************************************************************
        //Procedure para apresentar o ultimousuario logado
        // Criando em: 24/01/2012 Por: Filé

        public int UltimoPopup()
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT TOP 1 CodigoLog FROM RegLogBolinas order by HoraLog DESC";

            SqlCommand cmd = new SqlCommand(Sql, cn);

            resposta = Convert.ToInt32(cmd.ExecuteScalar());

            // MessageBox.Show("Operação Realizada!!!" + tmpCodigoLog);

            cn.Close();

            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para apresentar o status do ultimo usuario logado
        // Criando em: 24/01/2012 Por: Filé

        public String StatusUltimoPopup( int t)
        {
            String resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT TOP 1 StatusLog FROM RegLogBolinas where CodigoLog = " + t;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (String)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para apresentar o codigo do ultimo usuario logado
        // Criando em: 24/01/2012 Por: Filé

        public int UsuarioUltimoPopup(int t)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT TOP 1 UsrLog FROM RegLogBolinas where CodigoLog = " + t;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para apresentar determinado dado de uma determinada tabela
        // Criando em: 24/01/2012 Por: Filé
        // t = Código do usuario

        public string PegaNomeUsr(int t)
        {
            string resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT TOP 1 NomeUsr FROM Usuario where CodigoUsr = " + t;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (string)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para dados int do usuario.
        // Criando em: 27/01/2012 Por: Filé
        // Modificada em: 17/07/2012 Por: Filé

        public int UsuarioDadosInt(string tmp_codigo)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            // CodigoMsg
            string Sql = "SELECT TOP 1 " + tmp_codigo + " FROM Usuario where '" + System.Environment.UserName + "' in (LoginUsr)";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            // resposta = 0;
            // if (NomeLoginUsr() == System.Environment.UserName)
            // {
                resposta = (int)cmd.ExecuteScalar();
            // }
            cn.Close();
            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para pegar dados String do usuario.
        // Criando em: 17/07/2012 Por: Filé
        // Modificada em:  Por:

        public String UsuarioDadosString(string tmp_codigo)
        {
            String resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 " + tmp_codigo + " FROM Usuario where '" + System.Environment.UserName + "' in (LoginUsr)";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (String)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para apresentar a ultima mensagem lida
        // Criando em: 27/01/2012 Por: Filé

        public string MensagemLoginUsr(int tmp_CodMsg)
        {
            string resposta;

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 Conteudo FROM Mensagem where CodigoMsg = " + tmp_CodMsg;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (string)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para verificar a ultima mensagem criada
        // Criando em: 30/01/2012 Por: Filé
        // Modificada em: 21/05/2012 Por: Filé

        public int UltMsgLoginUsr(String tmp_CodMsg, String tmp_Tabela)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            // CodigoMsg, Mensagem
            string Sql = "SELECT TOP 1 " + tmp_CodMsg + " FROM " + tmp_Tabela + " order by " + tmp_CodMsg + " desc";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (int)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para gravar no cadastro de usuario a ultima mensagem lida.
        // Criando em: 30/01/2012 Por: Filé
        public void GravaMensagemLoginUsr(int CodMsg)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Usuario SET CodigoMsg = " + CodMsg + " where '" + System.Environment.UserName + "' in (LoginUsr)";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para gravar no cadastro de usuario a ultima mensagem lida.
        // Criando em: 30/01/2012 Por: Filé
        public void GravaReuniaoIntLoginUsr(string CodMsg)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Usuario SET Status = '" + CodMsg + "' where '" + System.Environment.UserName + "' in (LoginUsr) ";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para verificar se o usuario esta com pendencia
        // Criando em: 19/04/2012 Por: Filé

        public string PendenciaLoginUsr()
        {
            string resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 Pendencia FROM Usuario where pendencia is not null and '" + System.Environment.UserName + "' in (LoginUsr)";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (string)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para verificar o status do usuario
        // Criando em: 19/04/2012 Por: Filé

        public string StatusLoginUsr()
        {
            string resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 Status FROM Usuario where '" + System.Environment.UserName + "' in (LoginUsr)";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (string)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para testar a conexão
        public string TestaConexao()
        {
            string resposta;
            resposta = "EE";
            // SqlConnection cn = new SqlConnection();
            try
            {
                SqlConnection cn = new SqlConnection();
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                cn.Open();
                string Sql = "SELECT TOP 1 NomeUsr from Usuario";
                SqlCommand cmd = new SqlCommand(Sql, cn);
                SqlDataAdapter adpt = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adpt.Fill(dt);
                cn.Close();
                resposta = "OK";
                // MessageBox.Show("1 - " + resposta);
            }

            catch
            {
                // MessageBox.Show("1 - " + resposta);
            }

            return resposta;
        }



        //******************************************************************************************************************************
        //Procedure para preencher Data List Relatorio
        // FUNCIONANDO
        //public void PopulaListRel(ListBox list, string D_Ini, string D_Fin, string u)
        //{
        //    SqlConnection cn = new SqlConnection();
        //    cn.ConnectionString = ConexaoBD.StringDeConexao;
        //    cn.Open();

        //    string Sql = "SELECT HoraLog, StatusLog FROM RegLogBolinas where  HoraLog>= '" + D_Ini + "' and HoraLog<= '" + D_Fin + "' and UsrLog in (SELECT CodigoUsr FROM Usuario where NomeUsr like '%" + u + "%')";

        //    SqlCommand cmd = new SqlCommand(Sql, cn);

        //    SqlDataAdapter adpt = new SqlDataAdapter(cmd);

        //    DataTable dt = new DataTable();
        //    adpt.Fill(dt);
        //    list.DisplayMember = "HoraLog";
        //    list.ValueMember = "HoraLog";
        //    list.DataSource = dt;
        //    cn.Close();
        //}


        //******************************************************************************************************************************
        //Procedure para preencher Data grid

        public void PopulaGrid(System.Windows.Forms.DataGridView grid, string t)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT CodigoUsr, NomeUsr, Status, Nascimento, Setor, foto, fotao, Pendencia, Formatura, EntSeg, EntTer, EntQua, EntQui, EntSex, SaiSeg, SaiTer, SaiQua, SaiQui, SaiSex, Admissao, Observacao, FolgaP, Telefone, Ramal, LoginUsr FROM " + t + " order by Setor,NomeUsr";
            //                    0          1        2       3           4      5     6      7          8          9
            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();

            adpt.Fill(dt);
            grid.DataSource = dt;
        }

        //******************************************************************************************************************************
        //Procedure p/ Pesquisar Usuários

        public void PesquisaUsr(System.Windows.Forms.DataGridView grid, string chave)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT CodigoUsr, NomeUsr, Status, Nascimento, Setor, foto, fotao, Pendencia, Formatura, EntSeg, EntTer, EntQua, EntQui, EntSex, SaiSeg, SaiTer, SaiQua, SaiQui, SaiSex, Admissao, Observacao, FolgaP, Telefone, Ramal, LoginUsr FROM  Usuario where  NomeUsr like '%" + chave + "%' or  Status like '%" + chave + "%' or Setor like'%" + chave + "%'";

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();

            adpt.Fill(dt);
            grid.DataSource = dt;
        }

        //******************************************************************************************************************************
        //Obter Fotão para o picturebox

        public void PegaFotao(System.Windows.Forms.PictureBox foto, string chave)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT fotao FROM  Usuario where CodigoUsr=" + chave;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
            foto.Image = (Image)converter.ConvertFrom(cmd.ExecuteScalar());

        }

        //******************************************************************************************************************************
        // Obter Foto para o picturebox
        // Criando em: 25/01/2012 Por: Filé

        public void PegaFoto(System.Windows.Forms.PictureBox foto, int chave)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT fotao FROM  Usuario where CodigoUsr=" + chave;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            try
            {
                System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
                foto.Image = (Image)converter.ConvertFrom(cmd.ExecuteScalar());
            }
            catch
            { }

        }

        //******************************************************************************************************************************
        //Procedure para validar nome
        
        public void PegaNome(System.Windows.Forms.TextBox texto, string chave)
        {
            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT NomeUsr FROM  Usuario where CodigoUsr=" + chave;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            texto.Text = "Olá " + cmd.ExecuteScalar().ToString().Trim() + "!" + Environment.NewLine + "Selecione uma opção.";
        }

        //******************************************************************************************************************************
        //Procedure Para Apagar Usuários

        public void ApagarUsuario(string cod)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            try
            {
                string Sql = "delete from  Usuario where CodigoUsr=" + cod;
                SqlCommand cmd = new SqlCommand(Sql, cn);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Usuário Excluído!!!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Erros!!!");

            }
            finally
            {
                cn.Close();
            }
        }
        
        //******************************************************************************************************************************
        //Testa Usuário se ele existe
        
        public int TestaUsr(string cod)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "select count(*) from Usuario where CodigoUsr=" + cod;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            return resposta;

        }
        
        //******************************************************************************************************************************
        //Insere Item na tabela Status

        public void CadastraSituacao(Situacao S)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            try
            {
                string Sql = "insert into CadStatus (descricao) values (@desc)";
                SqlCommand cmd = new SqlCommand(Sql, cn);
                SqlParameter c = new SqlParameter("@desc", SqlDbType.NVarChar, 50);
                c.Value = S.Descricao;
                cmd.Parameters.Add(c);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Status Inserido\nCom Sucesso!!!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Erros!!!");
            }
            finally
            {
                cn.Close();
            }
        }
        
        //******************************************************************************************************************************
        //Procedure para popular Listbox Status

        public void PopulaList(System.Windows.Forms.ListBox list, string tabela)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT * FROM " + tabela;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            adpt.Fill(dt);
            list.DisplayMember = "descricao";
            list.ValueMember = "cod";
            list.DataSource = dt;
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para popular combo status
        public void PopulaCombo(System.Windows.Forms.ComboBox list, string tabela)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT * FROM " + tabela;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            adpt.Fill(dt);
            list.DisplayMember = "descricao";
            list.ValueMember = "cod";
            list.DataSource = dt;
            cn.Close();
        }

        //******************************************************************************************************************************
        //Função para testar status

        public int TestaStatus(string S, string t)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "select count(*) from CadStatus where " + t + "='" + S + "'";

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            return resposta;
        }

        //******************************************************************************************************************************
        //Função para testar Mensagem
        
        public int TestaMensagem(int cod)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "select count(*) from Mensagem where CodigoMsg=" + cod;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            return resposta;
        }

        //******************************************************************************************************************************
        //Função para testar Assunto

        public int TestaAssunto(string assunto)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "select count(*) from Mensagem where Assunto='" + assunto + "'";

            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para atualizar Status

        public void AtualizaStatus(Situacao s)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            try
            {
                string Sql = "update CadStatus set descricao=@desc where cod=@cod";
                SqlCommand cmd = new SqlCommand(Sql, cn);

                SqlParameter pcodigo = new SqlParameter("@Cod", SqlDbType.Int);
                pcodigo.Value = s.Codigo;
                cmd.Parameters.Add(pcodigo);

                SqlParameter pdescricao = new SqlParameter("@desc", SqlDbType.NVarChar, 50);
                pdescricao.Value = s.Descricao;
                cmd.Parameters.Add(pdescricao);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Atualizado!!!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Erros!!!");
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure Apaga Status
        
        public void ApagaStatus(Situacao s)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            try
            {
                string Sql = "delete from CadStatus where cod=@cod";
                SqlCommand cmd = new SqlCommand(Sql, cn);

                SqlParameter pcodigo = new SqlParameter("@Cod", SqlDbType.Int);
                pcodigo.Value = s.Codigo;
                cmd.Parameters.Add(pcodigo);

                cmd.ExecuteNonQuery();
                MessageBox.Show("Excluído!!!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Erros!!!");
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Chama Usuário via código
        
        public void GetUsuario(int cod, System.Windows.Forms.TextBox nome, System.Windows.Forms.DateTimePicker nasc, System.Windows.Forms.ComboBox sts, System.Windows.Forms.ComboBox setor, System.Windows.Forms.PictureBox foto)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT * FROM Usuario where CodigoUsr=" + cod;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataReader R = cmd.ExecuteReader();

            while (R.Read())
            {
                nome.Text = R["NomeUsr"].ToString();
                nasc.Value = Convert.ToDateTime(R["Nascimento"].ToString());
                sts.Text = R["Status"].ToString();
                setor.Text = R["Setor"].ToString();
                
                System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
                foto.Image = (Image)converter.ConvertFrom(R["fotao"]);
            }
        }

        //******************************************************************************************************************************
        //Procedure para popular Combo Mensagens

        public void PopulaComboMsg(System.Windows.Forms.ComboBox list, string tabela)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT * FROM " + tabela;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            adpt.Fill(dt);
            list.DisplayMember = "Assunto";
            list.ValueMember = "CodigoMsg";
            list.DataSource = dt;
            cn.Close();
        }

        //******************************************************************************************************************************
        public void BuscaConteudo(int cod, System.Windows.Forms.TextBox texto)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT Conteudo FROM Mensagem where CodigoMsg=" + cod;

            SqlCommand cmd = new SqlCommand(Sql, cn);
            texto.Text = cmd.ExecuteScalar().ToString();
        }

        //******************************************************************************************************************************
        // Procedure para apagar Mensagem

        public void ApagaMensagem(int cod)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            try
            {
                string Sql = "delete from Mensagem where CodigoMsg=" + cod;

                SqlCommand cmd = new SqlCommand(Sql, cn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Operação Realizada!!!");
            }
            catch
            {
                MessageBox.Show("Falha ao Apagar!!!");
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure para popular grid telão
        public void PopulaTelao(System.Windows.Forms.DataGridView telao)
        {
            SqlConnection cn = new SqlConnection();
            try
            {
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                cn.Open();

                string Sql = "SELECT foto,NomeUsr,Status,Setor,Pendencia from Usuario where Setor<>'6 - Inativo'  order by Setor,NomeUsr";

                SqlCommand cmd = new SqlCommand(Sql, cn);

                SqlDataAdapter adpt = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();

                adpt.Fill(dt);
                telao.DataSource = dt;
            }
            catch
            {
                //Tirei System.Windows.Forms.MessageBox.Show("1 - A conexão não foi Estabelecida\n1-Verifique o cabo de rede e tecle \"enter\"\n2-Procure o suporte!!!");                
            }
        }

        //******************************************************************************************************************************
        //Procedure para popular grid telão PopUp
        // Criado em 26-01-2012 por Filé
        // Modificado em: 11-07-2012 por Filé
        public void PopulaTelaoPopup(System.Windows.Forms.DataGridView telao)
        {
            SqlConnection cn = new SqlConnection();
            try
            {
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                cn.Open();

                string Sql = "SELECT foto,NomeUsr,Status,Setor,Pendencia,Telefone,Ramal,CodigoUsr,LoginUsr from Usuario where Setor<>'6 - Inativo' and not Status='Ausente' order by Setor,NomeUsr";

                SqlCommand cmd = new SqlCommand(Sql, cn);

                SqlDataAdapter adpt = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();

                adpt.Fill(dt);
                telao.DataSource = dt;
            }
            catch
            {
                //Tirei System.Windows.Forms.MessageBox.Show("2 - A conexão não foi Estabelecida\n1-Verifique o cabo de rede e tecle \"enter\"\n2-Procure o suporte!!!");
            }
        }


        //******************************************************************************************************************************
        //Altera status
        
        public void AlteraStatus()
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            //string Sql = "update Usuario set Status='Ausente' where Status='Externo'";
            string Sql = "update Usuario set Status='Ausente' where  (NOT (Status IN ('Férias', 'Folga Programada', 'Viagem', 'Ausente')))";

            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteNonQuery();
        }

        //******************************************************************************************************************************
        public void AlteraStatus(int cod, string status)
        {

            try
            {
                SqlConnection cn = new SqlConnection();
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                cn.Open();

                string Sql = "update Usuario set Status='" + status + "' where CodigoUsr=" + cod;

                SqlCommand cmd = new SqlCommand(Sql, cn);
                cmd.ExecuteNonQuery();
            }
            catch
            {

            }
        }


        //******************************************************************************************************************************
        //Procedure para verificar o status do usuario
        // Criando em: 09/05/2012 Por: Filé

        public string VerificaStatus(int cod)
        {
            string resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 Status FROM Usuario where CodigoUsr = " + cod;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (string)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }
        
        
        //******************************************************************************************************************************
        // Pinta Usuários
        // dtgTelao
        public void PintaStatus(System.Windows.Forms.DataGridView Telao)
        {
            int n = Telao.RowCount;

            for (int i = 0; i < n; i++)
            {
                if (Telao.Rows[i].Cells[4].Value.ToString().TrimEnd() != "")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
                    Telao.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Pink;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "1 - Diretoria")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "2 - Administrativo")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Green;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "3 - Engenharia")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Navy;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "4 - TI")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Brown;
                }
                if (Telao.Rows[i].Cells[2].Value.ToString().TrimEnd() == "Ausente")
                {

                    //Telao.Rows[i].Cells[0].Value = MakeGrayscale(Telao.Rows[i].Cells[0].Value);
                    //Telao.Rows[i].Cells[1].Value = MakeGrayscale(Telao.Rows[i].Cells[1].Value);
                    //Telao.Rows[i].Cells[2].Value = MakeGrayscale(Telao.Rows[i].Cells[2].Value);
                    
                    Telao.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Silver;
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Gray;
                }
            }
        }


        //public static Bitmap MakeGrayscale(Bitmap original)
        public Bitmap MakeGrayscale(Bitmap original)
        {
            //fazer um bitmap vazio do mesmo tamanho que o original
            Bitmap newBitmap = new Bitmap(original.Width, original.Height);

            for (int i = 0; i < original.Width; i++)
            {
                for (int j = 0; j < original.Height; j++)
                {
                    //obter o pixel da imagem original
                    Color originalColor = original.GetPixel(i, j);

                    //criar a versão em tons de cinza de um pixel
                    int grayScale = (int)((originalColor.R * .3) + (originalColor.G * .59)
                        + (originalColor.B * .11));

                    //criar o objeto de cor
                    Color newColor = Color.FromArgb(grayScale, grayScale, grayScale);

                    //conjunto de pixels da nova imagem para a versão em tons de cinza
                    newBitmap.SetPixel(i, j, newColor);
                }
            }

            return newBitmap;
        }



        //******************************************************************************************************************************
        // Pinta Usuários
        // dtgTelao
        public void PintaStatusPopUp(System.Windows.Forms.DataGridView Telao)
        {
            int n = Telao.RowCount;

            for (int i = 0; i < n; i++)
            {
                Telao.Rows[i].Cells[0].ToolTipText = "Ramal: " + Telao.Rows[i].Cells[6].Value.ToString() + "\nCelular: " + Telao.Rows[i].Cells[5].Value.ToString();
                Telao.Rows[i].Cells[1].ToolTipText = "Ramal: " + Telao.Rows[i].Cells[6].Value.ToString() + "\nCelular: " + Telao.Rows[i].Cells[5].Value.ToString();
                Telao.Rows[i].Cells[2].ToolTipText = "Ramal: " + Telao.Rows[i].Cells[6].Value.ToString() + "\nCelular: " + Telao.Rows[i].Cells[5].Value.ToString();

                if (Telao.Rows[i].Cells[4].Value.ToString().TrimEnd() != "")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
                    Telao.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Pink;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "1 - Diretoria")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "2 - Administrativo")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Green;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "3 - Engenharia")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Navy;
                }
                if (Telao.Rows[i].Cells[3].Value.ToString().TrimEnd() == "4 - TI")
                {
                    Telao.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Brown;
                }
            }
        }

        //******************************************************************************************************************************
        //Registra Opção no Log

        public void RegistraLog(int cod, DateTime hora, string status)
        {
            SqlConnection cn = new SqlConnection();

            try
            {
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                string Sql = "insert into RegLogBolinas (HoraLog,StatusLog,UsrLog)   values (@hora, @status,@cod)";
                cn.Open();

                SqlCommand cmd = new SqlCommand(Sql, cn);

                SqlParameter pdata = new SqlParameter("@hora", SqlDbType.DateTime);
                pdata.Value = hora;
                cmd.Parameters.Add(pdata);

                SqlParameter pstatus = new SqlParameter("@status", SqlDbType.NVarChar, 50);
                pstatus.Value = status;
                cmd.Parameters.Add(pstatus);

                SqlParameter pcodigo = new SqlParameter("@Cod", SqlDbType.Int);
                pcodigo.Value = cod;
                cmd.Parameters.Add(pcodigo);

                cmd.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Chame o Suporte!!!");
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure para preenche Grid de Aniversariantes

        public void GetAniversario(System.Windows.Forms.DataGridView grid, int mes)
        {
            SqlConnection cn = new SqlConnection();
            try
            {
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                cn.Open();

                string Sql = "select foto,NomeUsr,Nascimento from Usuario where Setor<>'6 - Inativo' and  Month(Nascimento)=" + mes + "  Order by Day(Nascimento), NomeUsr";

                SqlCommand cmd = new SqlCommand(Sql, cn);

                SqlDataAdapter adpt = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();

                adpt.Fill(dt);
                grid.DataSource = dt;
            }
            catch
            {

            }
        }

        //******************************************************************************************************************************
        //Procedure para trocar mensagem de exibição

        public void PublicaMensagem(string texto)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "update exibir set conteudo=@texto";

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlParameter pTexto = new SqlParameter("@texto", SqlDbType.NVarChar, 300);
            pTexto.Value = texto;
            cmd.Parameters.Add(pTexto);

            cmd.ExecuteNonQuery();

        }
        
        //******************************************************************************************************************************
        // Procedure para exibir a mensagem

        public void ExibeMensagem(System.Windows.Forms.TextBox texto)
        {
            try
            {
                SqlConnection cn = new SqlConnection();
                cn.ConnectionString = ConexaoBD.StringDeConexao;
                cn.Open();

                string Sql = "select conteudo from exibir";

                SqlCommand cmd = new SqlCommand(Sql, cn);

                texto.Text = cmd.ExecuteScalar().ToString();
            }
            catch
            { 
            }
        }

        //******************************************************************************************************************************
        public void HoraTrabalhada(System.Windows.Forms.TextBox texto, string dia, string cod)
        {
            SqlConnection cn = new SqlConnection();
            //try
            //{

            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql =


" declare @ret varchar(7) declare @Diaria int declare @Almoco int declare @total1 int" +
" set @Diaria = datediff(MINUTE,(select HoraLog from RegLogBolinas where UsrLog=" + cod + " and DAY(HoraLog)=" + dia + " and StatusLog='Cheguei'), (select HoraLog from RegLogBolinas where UsrLog=1234 and DAY(HoraLog)=" + dia + " and StatusLog='Fui Para Casa'))" +
" set @Almoco = datediff(MINUTE,(select HoraLog from RegLogBolinas where UsrLog=" + cod + " and DAY(HoraLog)=" + dia + " and StatusLog='Vou Almoçar'),(select HoraLog from RegLogBolinas where UsrLog=" + cod + " and DAY(HoraLog)=" + dia + " and StatusLog='Voltei do Almoço'))" +
" if @Almoco is null" +
" begin" +
"	 set @total1 = @Diaria - 60" +
" end" +
" else if @Diaria is null" +
" begin" +
"		set @total1 = 0" +
" end" +
" else" +
" begin" +
"   set  @total1 = @Diaria - @Almoco" +
" end" +
" if @total1 <> 0" +
" begin" +
"	set @ret = dbo.ConverteMin(@total1)" +
" end" +
" else" +
" begin" +
"	set @ret = 0" +
" end" +
" select  @ret";

            SqlCommand cmd = new SqlCommand(Sql, cn);

            texto.Text = cmd.ExecuteScalar().ToString();

            //}
            //catch
            //{ }
            //finally { cn.Close(); }
        }

        //******************************************************************************************************************************
        public void PegaHora(System.Windows.Forms.TextBox texto, int cod)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;

            try
            {
                SqlCommand CmdIns = new SqlCommand();
                CmdIns.Connection = cn;
                CmdIns.CommandType = CommandType.StoredProcedure;
                CmdIns.CommandText = "dbo.ConverteMin";

                //Inserindo Parâmetros
                SqlParameter pcodigoMsg = new SqlParameter("@MINUTOS", SqlDbType.Int);
                pcodigoMsg.Value = cod;
                CmdIns.Parameters.Add(pcodigoMsg);

                //Executando a procedure
                cn.Open();
                texto.Text = CmdIns.ExecuteNonQuery().ToString();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure para incluir uma Enquete nova.
        // Criando em: 17/05/2012 Por: Filé
        public void IncluiEnquete(string tmp_Assunto, string tmp_Resposta1, string tmp_Resposta2, string tmp_Resposta3)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "INSERT into Enquete ( Assunto, Resposta1, Resposta2, Resposta3, Data, Total1, Total2, Total3 ) values ( '" + tmp_Assunto + "','" + tmp_Resposta1 + "','" + tmp_Resposta2 + "','" + tmp_Resposta3 + "','" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "', 0, 0, 0 )";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para popular enquete
        // Criando em: 17/05/2012 Por: Filé
        public void PopulaEnquete(System.Windows.Forms.ComboBox list, string tabela)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT * FROM " + tabela;

            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            adpt.Fill(dt);
            list.DisplayMember = "descricao";
            list.ValueMember = "cod";
            list.DataSource = dt;
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para preencher Data grid Enquete
        // Criando em: 17/05/2012 Por: Filé
        public void PopulaGridEnquete(System.Windows.Forms.DataGridView grid, string t)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();

            string Sql = "SELECT ID,Data,Assunto,Resposta1,Resposta2,Resposta3,Total1,Total2,Total3 FROM " + t;
            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();

            adpt.Fill(dt);
            grid.DataSource = dt;
        }

        //******************************************************************************************************************************
        //Procedure para atualizar uma Enquete.
        // Criando em: 17/05/2012 Por: Filé
        public void AtualizaEnquete(string tmp_ID, string tmp_Assunto, string tmp_Resposta1, string tmp_Resposta2, string tmp_Resposta3)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Enquete SET Assunto = '" + tmp_Assunto + "', Resposta1 = '" + tmp_Resposta1 + "', Resposta2 = '" + tmp_Resposta2 + "', Resposta3 = '"+ tmp_Resposta3 + "' where ID = '" + tmp_ID + "'";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Testa Enquete se ele existe
        // Criando em: 18/05/2012 Por: Filé
        public int TestaEnquete(string tmp_ID)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "select count(*) from Enquete where ID = '" + tmp_ID + "'";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            return resposta;

        }

        //******************************************************************************************************************************
        //Procedure Para Apagar Enquete
        // Criando em: 18/05/2012 Por: Filé
        public void ApagarEnquete(string tmp_ID)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            try
            {
                string Sql = "delete from  Enquete where ID = '" + tmp_ID + "'";
                SqlCommand cmd = new SqlCommand(Sql, cn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Enquete Excluída!!!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Erros!!!");

            }
            finally
            {
                cn.Close();
            }
        }

        //******************************************************************************************************************************
        //Procedure para gravar no cadastro de usuario a ultima enquete respondida.
        // Criando em: 21/05/2012 Por: Filé

        public void GravaEnqueteLoginUsr(int CodMsg)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Usuario SET CodigoEnquete = '" + CodMsg + "' where '" + System.Environment.UserName + "' in (LoginUsr)";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para verificar a ultima mensagem criada
        // Criando em: 21/01/2012 Por: Filé

        public String PegaEnqueteUsr(String tmp_Campo, String tmp_CodMsg)
        {
            String resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            // CodigoMsg, Mensagem
            string Sql = "SELECT TOP 1 " + tmp_Campo + " FROM Enquete WHERE ID = '" + tmp_CodMsg + "'";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (String)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para atualizar uma Enquete.
        // Criando em: 22/05/2012 Por: Filé
        public void AtualizaEnqueteUsr(string tmp_ID, string tmp_Campo)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Enquete SET " + tmp_Campo + " = " + tmp_Campo + " + 1 where ID = '" + tmp_ID + "'";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para preencher Data grid com os usuários que não responderam a Enquete
        // Criando em: 37/05/2012 Por: Filé
        public void PopulaGridEnquetePendente(System.Windows.Forms.DataGridView grid, string t)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT NomeUsr FROM Usuario where  CodigoEnquete < '" + t + "' order by NomeUsr";
            SqlCommand cmd = new SqlCommand(Sql, cn);

            SqlDataAdapter adpt = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();

            adpt.Fill(dt);
            grid.DataSource = dt;
        }

        //******************************************************************************************************************************
        //Procedure para pegar a solicitação de chat por outro usuário.
        // Criando em: 12/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public int ChatPara(int u)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 IDChat FROM Chat WHERE '" + u + "' in (ParaUsr) and ParaOpen='CL' order by IDChat DESC";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            //resposta = (int)cmd.ExecuteScalar();
            //resposta = Convert.ToInt16(cmd.ExecuteScalar());
            resposta = Convert.ToInt32(cmd.ExecuteScalar());
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para pegar determinada coluna int da tabela chat de um chat específico.
        // Criando em: 12/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public int ChatDadosInt(String d, int c)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 " + d + " FROM Chat where IDChat = " + c ;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (int)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }

        //******************************************************************************************************************************
        //Procedure para pegar determinada coluna string da tabela chat de um chat específico.
        // Criando em: 13/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public String ChatDadosString(String d, int c)
        {
            String resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 " + d + " FROM Chat where IDChat = " + c;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (String)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para gravar o status do chat.
        // Criando em: 12/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public void GravaStatusChat(String s, int c)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Chat SET ParaOpen = '" + s + "' where IDChat = " + c;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }

        //******************************************************************************************************************************
        //Procedure para criar o chat.
        // Criando em: 12/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public void IncluiChat(int d, int p, String t, String n)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "INSERT into Chat ( Data, ParaOpen, DeUsr, ParaUsr, Obse ) values ( '" + t + "', 'CL', " + d + "," + p + ", '" + n + "' )";
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }


        //******************************************************************************************************************************
        //Procedure para gravar a mensagem escrita do chat.
        // Criando em: 12/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public void GravaMsgChat(String s, int c)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "UPDATE Chat SET Obse = '" + s + "' + Obse where IDChat = " + c;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            cmd.ExecuteScalar();
            cn.Close();
        }


        //******************************************************************************************************************************
        //Procedure para pegar determinada coluna string da tabela chat de um chat específico.
        // Criando em: 12/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public String ChatMsg(int c)
        {
            String resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 Obse FROM Chat where IDChat = " + c;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (String)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }


        //******************************************************************************************************************************
        //Procedure para pegar o numero do novo Chat criado.
        // Criando em: 13/07/2012 Por: Filé
        // Modificada em:   /  /     Por:

        public int PegaNovoChat(int d, int p, String t)
        {
            int resposta;
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConexaoBD.StringDeConexao;
            cn.Open();
            string Sql = "SELECT TOP 1 IDChat FROM Chat where Data = '" + t + "' and DeUsr = " + d + " and ParaUsr = " + p;
            SqlCommand cmd = new SqlCommand(Sql, cn);
            resposta = (int)cmd.ExecuteScalar();
            cn.Close();
            return resposta;
        }
        //******************************************************************************************************************************
    }
}

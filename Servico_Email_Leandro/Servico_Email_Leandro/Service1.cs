using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Servico_Email_Leandro
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            ThreadStart start = new ThreadStart(vertificarEmailPendente);
            Thread thread = new Thread(start);

            thread.Start();
        }

        protected override void OnStop()
        {
        }

        public void vertificarEmailPendente()
        {
            while (true)
            {
                Thread.Sleep(5000);

                SqlConnection sqlconnection = new SqlConnection("");
                SqlCommand cmd = new SqlCommand("");

                sqlconnection.Open();

                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    enviarEmail(reader["EmailOrigem"].ToString(),
                        reader["EmailDestino"].ToString(),
                        reader["NomeOrige"].ToString(),
                        reader["NomeDestino"].ToString(),
                        reader["Assunto"].ToString(),
                        reader["Mensagem"].ToString());

                    atualizarEmail(Convert.ToInt32(reader["Id"].ToString()));
                }

                reader.Close();
                sqlconnection.Close();
            }
        }

        private void enviarEmail(string emailOrigem,
                                 string emailDestino,
                                 string nomeOrigem,
                                 string nomeDestino,
                                 string assunto,
                                 string mensagemEmail)
        {
            MailAddress origem = new MailAddress(emailOrigem, nomeOrigem);
            MailAddress destino = new MailAddress(emailDestino, nomeDestino);

            MailMessage mensagem = new MailMessage(origem, destino);

            mensagem.Subject = assunto;
            mensagem.Body = assunto;

            SmtpClient smtp = new SmtpClient();
            smtp.Host = "leandro.visualc@gmail.com";
            smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;

            smtp.Credentials = new NetworkCredential(origem.Address, "mundial2015");
        }

        private void atualizarEmail(int id)
        {
            SqlConnection sqlConnection = new SqlConnection();
            SqlCommand cmdUpdate = new SqlCommand();

            cmdUpdate.Parameters.Add(new SqlParameter("@Id", id));

            sqlConnection.Open();
            cmdUpdate.ExecuteNonQuery();
            sqlConnection.Close();
        }

    }
}

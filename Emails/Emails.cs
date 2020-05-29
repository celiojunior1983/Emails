using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace Emails
{
    public static class Emails
    {
        /// <summary>
        /// Envia um email simples com Assunto e Mensagem para os Dados Salvos no AppConfig
        /// </summary>
        /// <param name="assunto">Assunto do Email</param>
        /// <param name="mesagem">Mensagem do Email</param>
        /// <param name="NomeExibicao">Nome que será exibido no email enviado</param>
        /// <returns></returns>
        public static bool EnviarEmail(string assunto, string mesagem, string NomeExibicao)
        {
            //Cria o objeto que envia o e-mail 
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["smtpServer"].ToString());

            //Cria o endereço de email do remetente 
            MailAddress de = new MailAddress(ConfigurationManager.AppSettings["smtpUser"].ToString(), NomeExibicao);

            //Cria o endereço de email do destinatário -->
            MailAddress para = new MailAddress(ConfigurationManager.AppSettings["EmailFrom"].ToString());

            //MailAddress copia = new MailAddress("INFORME UM EMAIL");

            MailMessage mensagem = new MailMessage(de, para);

            mensagem.IsBodyHtml = true;

            //Assunto do email 
            mensagem.Subject = assunto;

            //Conteúdo do email
            mensagem.Body = mesagem;


            try
            {
                client.EnableSsl = true;

                client.Port = 25;
                //Envia o email 
                client.Send(mensagem);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Envia um email individual
        /// </summary>
        /// <param name="assunto">Titulo do email</param>
        /// <param name="mesagem">Conteudo da mensagem</param>
        /// <param name="destinatorio">Email do destinatario</param>
        /// <param name="NomeExibicao">Nome que será exibido no email enviado</param>
        /// <returns></returns>
        public static bool EnviarEmail(string assunto, string mesagem, string destinatorio, string NomeExibicao)
        {
            //Cria o objeto que envia o e-mail 
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["smtpServer"].ToString());

            //Cria o endereço de email do remetente 
            MailAddress de = new MailAddress(ConfigurationManager.AppSettings["smtpUser"].ToString(), NomeExibicao);

            //Cria o endereço de email do destinatário -->
            MailAddress para = new MailAddress(destinatorio);

            //MailAddress copia = new MailAddress("INFORME UM EMAIL");

            MailMessage mensagem = new MailMessage(de, para);

            mensagem.IsBodyHtml = true;

            //Assunto do email 
            mensagem.Subject = assunto;

            //Conteúdo do email
            mensagem.Body = mesagem;


            try
            {
                client.EnableSsl = true;

                client.Port = 587;
                //Envia o email 
                client.Send(mensagem);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Envia emails para uma lista
        /// </summary>
        /// <param name="assunto">Titulo do email</param>
        /// <param name="mesagem">Conteudo da mensagem</param>
        /// <param name="destinatorio">DataTable com os emails que serão enviados contendo apenas uma coluna com o nome 'email'</param>
        /// <param name="NomeExibicao">Nome que será exibido no email enviado</param>
        /// <returns></returns>
        public static bool EnviarEmail(string assunto, string mesagem, DataTable destinatorio, string NomeExibicao)
        {
            try
            {
                for (int i = 0; i < destinatorio.Rows.Count - 1; i++)
                {
                    //Cria o objeto que envia o e-mail 
                    SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["smtpServer"].ToString());

                    //Cria o endereço de email do remetente 
                    MailAddress de = new MailAddress(ConfigurationManager.AppSettings["smtpUser"].ToString(), NomeExibicao);

                    //Cria o endereço de email do destinatário -->
                    MailAddress para = new MailAddress(destinatorio.Rows[i].ItemArray[2].ToString());

                    //MailAddress copia = new MailAddress("INFORME UM EMAIL");

                    MailMessage mensagem = new MailMessage(de, para);

                    mensagem.IsBodyHtml = true;

                    //Assunto do email 
                    mensagem.Subject = assunto;

                    //Conteúdo do email
                    mensagem.Body = mesagem;

                    try
                    {
                        client.EnableSsl = true;

                        client.Port = 25;
                        //Envia o email 
                        client.Send(mensagem);

                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Envia um email individual com autenticação
        /// </summary>
        /// <param name="assunto">Titulo do email</param>
        /// <param name="mesagem">Conteudo da mensagem</param>
        /// <param name="destinatorio">Email do destinatario</param>
        /// <param name="NomeExibicao">Nome que será exibido no email enviado</param>
        /// <param name="usuario">usuário do email </param>
        /// <param name="senha"> senha do email cripto </param>
        /// <returns></returns>
        public static bool EnviarEmail(string assunto, string mesagem, string destinatorio, string NomeExibicao, string usuario, string senha)
        {
            //Cria o objeto que envia o e-mail 
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["smtpServer"].ToString());

            client.Credentials = new NetworkCredential(usuario, senha);
            //Cria o endereço de email do remetente 
            MailAddress de = new MailAddress(ConfigurationManager.AppSettings["smtpUser"].ToString(), NomeExibicao);

            //Cria o endereço de email do destinatário -->
            MailAddress para = new MailAddress(destinatorio);

            //MailAddress copia = new MailAddress("INFORME UM EMAIL");

            MailMessage mensagem = new MailMessage(de, para);

            mensagem.IsBodyHtml = true;

            //Assunto do email 
            mensagem.Subject = assunto;

            //Conteúdo do email
            mensagem.Body = mesagem;

            try
            {
                client.EnableSsl = true;

                client.Port = 587;
                //Envia o email 
                client.Send(mensagem);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Envia um email individual com autenticação
        /// </summary>
        /// <param name="assunto">Titulo do email</param>
        /// <param name="mesagem">Conteudo da mensagem</param>
        /// <param name="destinatorio">Email do destinatario</param>
        /// <param name="NomeExibicao">Nome que será exibido no email enviado</param>
        /// <param name="usuario">usuário do email </param>
        /// <param name="senha"> senha do email cripto </param>
        /// <param name="anexos">Anexar arquivos no email</param>
        /// <returns>Verdadeiro, quando consegue enviar</returns>
        public static bool EnviarEmail(string assunto, string mesagem, string destinatorio, string NomeExibicao, string usuario, string senha, ArrayList Anexo)//string Destinatario, string Remetente, string Assunto, string enviaMensagem, ArrayList anexos)
        {

            //Cria o objeto que envia o e-mail 
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["smtpServer"].ToString());

            client.Credentials = new NetworkCredential(usuario, senha);
            //Cria o endereço de email do remetente 
            MailAddress de = new MailAddress(ConfigurationManager.AppSettings["smtpUser"].ToString(), NomeExibicao);

            //Cria o endereço de email do destinatário -->
            MailAddress para = new MailAddress(destinatorio);

            //MailAddress copia = new MailAddress("INFORME UM EMAIL");

            MailMessage mensagem = new MailMessage(de, para);

            mensagem.IsBodyHtml = true;

            //Assunto do email 
            mensagem.Subject = assunto;

            //Conteúdo do email
            mensagem.Body = mesagem;

            foreach (string anexo in Anexo)
            {
                Attachment anexado = new Attachment(anexo, MediaTypeNames.Application.Octet);
                mensagem.Attachments.Add(anexado);
            }

            try
            {
                client.EnableSsl = true;

                client.Port = 587;
                //Envia o email 
                client.Send(mensagem);

                return true;
            }
            catch
            {
                return false;
            }
        }

    }
}

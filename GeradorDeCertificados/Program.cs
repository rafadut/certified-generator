using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace GeradorDeCertificados
{
    class Program
    {
        public class ParticipanteModel
        {
            private string nomeDoParticipante;
            private string nomeDoCursoOuForum;
            private short cargaHoraria;

            public string NomeDoParticipante
            {
                get
                {
                    return nomeDoParticipante;
                }

                set
                {
                    nomeDoParticipante = value;
                }
            }

            public string NomeDoCursoOuForum
            {
                get
                {
                    return nomeDoCursoOuForum;
                }

                set
                {
                    nomeDoCursoOuForum = value;
                }
            }

            public short CargaHoraria
            {
                get
                {
                    return cargaHoraria;
                }

                set
                {
                    cargaHoraria = value;
                }
            }
        }

        #region Constantes

        public const string CAMINHO_DIRETORIO = @"C:\GeradorDeCertificados\";
        public const string CAMINHO_ARQUIVOS_GERADOS = @"gerados\{0}.bmp";
        public const string NOME_ARQUIVO_EXCEL = "lista_participantes.xlsx";
        public const string NOME_ARQUIVO_IMAGEM = "modelo_certificado.bmp";

        #endregion

        static void Main(string[] args)
        {
            //Leitura da planilha
            var planilha = new LinqToExcel.ExcelQueryFactory(CAMINHO_DIRETORIO + NOME_ARQUIVO_EXCEL);

            var query =
                from row in planilha.Worksheet("Participantes")
                let item = new
                {
                    Nome = row["Nome"].Cast<string>(),
                    Curso = row["Curso"].Cast<string>(),
                    Horas = row["Horas"].Cast<short>(),
                    Minutos = row["Minutos"].Cast<short>(),
                    Email = row["Email"].Cast<string>()
                }
                where item.Nome != ""
                select item;

            var participantes = query.ToList();
            int quantidadeParticipantes = participantes.Count();
            string nomeDoArquivoGerado = string.Empty;
            Console.Write(Environment.NewLine);

            try
            {
                for (int i = 0; i < quantidadeParticipantes; i++)
                {
                    //Montagem do texto
                    StringBuilder sb = new StringBuilder();
                    sb.Append("Certificamos que Sr. (a) ");
                    sb.Append(participantes[i].Nome);
                    sb.Append(Environment.NewLine);
                    sb.Append("participou do ");
                    sb.Append(participantes[i].Curso);
                    sb.Append(Environment.NewLine);
                    sb.Append("com carga horária de ");
                    sb.Append(participantes[i].Horas);
                    sb.Append(" horas");
                    if (participantes[i].Minutos != 0)
                    {
                        sb.Append(" e ");
                        sb.Append(participantes[i].Minutos);
                        sb.Append(" minutos");
                    }
                    sb.Append(".");

                    string texto = sb.ToString();

                    //Leitura da imagem
                    string imagemModelo = CAMINHO_DIRETORIO + NOME_ARQUIVO_IMAGEM;
                    Bitmap imagem = (Bitmap)Image.FromFile(imagemModelo);
                    PointF posicao = new PointF(615f, 335f);

                    //Inserção do texto na imagem
                    using (Graphics graphics = Graphics.FromImage(imagem))
                    {
                        using (Font arialFont = new Font("Arial", 18))
                        {
                            graphics.DrawString(texto, arialFont, Brushes.Black, posicao);
                        }
                    }

                    //Gravação da nova imagem na pasta
                    nomeDoArquivoGerado = string.Format(
                        CAMINHO_DIRETORIO + CAMINHO_ARQUIVOS_GERADOS,
                        participantes[i].Nome);

                    using (MemoryStream memory = new MemoryStream())
                    {
                        using (FileStream fs = new FileStream(nomeDoArquivoGerado, FileMode.Create, FileAccess.ReadWrite))
                        {
                            imagem.Save(memory, ImageFormat.Jpeg);
                            byte[] bytes = memory.ToArray();
                            fs.Write(bytes, 0, bytes.Length);
                        }
                    }

                    Console.WriteLine("O certificado de "
                        + participantes[i].Nome
                        + " foi gerado.");
                }

                Console.Write(Environment.NewLine);
                Console.Write("Deseja enviar os certificados por email? (S/N) ");
                bool enviarEmails = Console.ReadLine().ToUpper().Equals("S");
                Console.Write(Environment.NewLine);

                if (enviarEmails)
                {
                    string enderecoEmailRemetente;
                    string chave;

                    //Definição do email do remetente
                    chave = string.Empty;
                    Console.Write("Digite o endereço de email do remetente: ");
                    enderecoEmailRemetente = Console.ReadLine();
                    Console.Write(string.Format("Digite a senha do email {0}: ", enderecoEmailRemetente));
                    ConsoleKeyInfo key;

                    do
                    {
                        key = Console.ReadKey(true);

                        if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                        {
                            chave += key.KeyChar;
                            Console.Write("*");
                        }
                        else
                        {
                            if (key.Key == ConsoleKey.Backspace && chave.Length > 0)
                            {
                                chave = chave.Substring(0, (chave.Length - 1));
                                Console.Write("\b \b");
                            }
                        }
                    }

                    while (key.Key != ConsoleKey.Enter);

                    Console.Write(Environment.NewLine);
                    Console.Write(Environment.NewLine);

                    for (int i = 0; i < quantidadeParticipantes; i++)
                    {
                        //Envio de email
                        MailMessage msg = new MailMessage();

                        msg.From = new MailAddress(enderecoEmailRemetente);
                        msg.To.Add(participantes[i].Email);
                        msg.Subject = "Certificado de conclusão do " + participantes[i].Curso;
                        msg.Body = "Olá, " + participantes[i].Nome + "!"
                            + Environment.NewLine
                            + Environment.NewLine
                            + "Segue em anexo o seu certificado de conclusão do " + participantes[i].Curso + "."
                            + Environment.NewLine
                            + Environment.NewLine
                            + "Obrigada!";
                        msg.Attachments.Add(new Attachment(nomeDoArquivoGerado));
                        SmtpClient client = new SmtpClient();
                        client.UseDefaultCredentials = false;
                        if (enderecoEmailRemetente.Contains("gmail"))
                        {
                            client.Host = "smtp.gmail.com";
                        }
                        else
                        {
                            client.Host = "smtp.live.com";
                        }
                        client.Port = 587;
                        client.EnableSsl = true;
                        client.Credentials = new NetworkCredential(enderecoEmailRemetente, chave);
                        client.Send(msg);

                        Console.WriteLine("O email para "
                            + participantes[i].Nome
                            + " foi enviado.");
                    }

                    Console.Write(Environment.NewLine);
                }

                Console.Write("Processamento concluído com sucesso.");
            }
            catch (Exception)
            {
                Console.Write(Environment.NewLine);
                Console.Write("Ocorreu um erro.");
            }
            finally
            {
                Console.ReadLine();
                Environment.Exit(0);
            }
        }
    }
}

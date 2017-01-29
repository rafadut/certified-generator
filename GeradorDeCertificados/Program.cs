using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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
                    Minutos = row["Minutos"].Cast<short>()
                }
                where item.Nome != "" 
                select item;

            var participantes = query.ToList();
            int quantidadeParticipantes = participantes.Count();

            for (int i = 0; i < quantidadeParticipantes; i++)
            {
                //Montagem do texto
                StringBuilder sb = new StringBuilder();
                sb.Append("Certificamos que Sr. (a) ");
                sb.Append(participantes[i].Nome);
                sb.Append("\n participou do ");
                sb.Append(participantes[i].Curso);
                sb.Append("\n com carga horária de ");
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
                string nomeDoArquivoGerado = string.Format(
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
            }
        }
    }
}

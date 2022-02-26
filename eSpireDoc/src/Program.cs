using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace src
{
    class Program
    {
        static void Main(string[] args)
        {
            string txt_arquivo_modelo = "Modelo.docx";
            string txt_arquivo = "{" + Guid.NewGuid().ToString().ToUpper() + "}_" + txt_arquivo_modelo;

            File.Copy(SistemaRetornaCaminhoModelo() + txt_arquivo_modelo, SistemaRetornaCaminhoGravacaoArquivo() + txt_arquivo);

            var document = new Document();
            document.LoadFromFile(SistemaRetornaCaminhoGravacaoArquivo() + txt_arquivo, Spire.Doc.FileFormat.Auto);

            document.Replace("@@nome@@", "Wanderson Caldas", true, true);
            document.Replace("@@email@@", "wcaldasti@gmail.com", true, true);
            document.Replace("@@telefone@@", "(99)99999-9999", true, true);
            document.Replace("@@endereço@@", "Avenida X, rua y", true, true);

            MemoryStream strBinario = new MemoryStream();

            document.SaveToStream(strBinario, FileFormat.Docx);            

            File.WriteAllBytes(SistemaRetornaCaminhoGravacaoArquivo() + "{" + Guid.NewGuid().ToString().ToUpper() + "}_" + txt_arquivo_modelo, strBinario.ToArray());
            File.Delete(SistemaRetornaCaminhoGravacaoArquivo() + txt_arquivo);
            
            strBinario.Close();
            document.Close();
        }

        private static string SistemaRetornaCaminhoModelo()
        {
            return @"C:\inetpub\wwwroot\github\SpireDoc\eSpireDoc\src\Modelo\";
        }

        private static string SistemaRetornaCaminhoGravacaoArquivo()
        {
            return @"C:\inetpub\wwwroot\github\SpireDoc\eSpireDoc\src\Modelo\";
        }
    }
}

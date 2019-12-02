using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FileUploadTest.Model
{
    public class ImportRowData
    {
        public string Acao { get; set; }
        public string ID_Externo { get; set; }
        public string Nome { get; set; } 
        public string DataNascimento { get; set; }
        public string Genero { get; set; }

        public string Login { get; set; }
        public string Senha { get; set; }
        public string CPF { get; set; }
        public string RG { get; set; }
        public string Email { get; set; }
        public string Celular { get; set; }

        public string ID_Externo_Responsavel { get; set; }
        public string Papel { get; set; }
        public string GroupCode { get; set; }
        public string GroupName { get; set; }
        public string GroupDescription { get; set; }
        public string TAG { get; set; }

    }
}

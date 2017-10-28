using System;
using NetOffice.ExcelApi;
using System.Text.RegularExpressions;


namespace sistema_de_concessionaria
{
    class Program
    {
        static void Main(string[] args)
        {
            //CriarExcel();
            //LerExcel();
            cadastrar_cliente();
            //cliente cliente1 = new cliente();
            //cliente1.validar_cep();

        }

        //funcao para criar excel
        static void CriarExcel(string nome, string documento, string email, string endereco, string cep, string data_cadastro)
        {
            Application ex = new Application();
            ex.Workbooks.Add();
            // verifica se celula esta vazia
            int contador = 0;
            do
            {
                contador++;
            } while (ex.Cells[contador, 1].Value != null);
            ex.Cells[contador, 1].Value = nome;
            ex.Cells[contador, 2].Value = documento;
            ex.Cells[contador, 3].Value = email;
            ex.Cells[contador, 4].Value = endereco;
            ex.Cells[contador, 5].Value = cep;
            ex.Cells[contador, 6].Value = data_cadastro;
            ex.ActiveWorkbook.SaveAs(@"C:\Users\bruno afonso\Desktop\sistema de concessionaria\clientes.xls", true);
            ex.Quit();
        }

      

        //metodo para criar o arquivo excel e escreve as informações do objeto
        static void LerExcel() {
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\bruno afonso\Desktop\sistema de concessionaria\clientes.xls");
            for(var i = 1; i <= 2; i++) {
                for(var e = 1; e <= 5; e++) {
                    string texto = ex.Cells[i,e].Value.ToString();
                    System.Console.WriteLine(texto);
                }
            }
        }

        //metodo cadastrar o cliente, instanciando classe 
        static void cadastrar_cliente() {
            cliente cliente1 = new cliente("", "", "", "", "", "");
            string Nome = cliente1.validar_nome();
            //chamar o metodo de acordo com o tipo de documento
            string Documento = cliente1.validar_documento();
            string Email = cliente1.validar_email();
            string Endereco = cliente1.validar_endereco();
            string Cep = cliente1.validar_cep();
            string Data_cadastro = cliente1.data_do_cadastro();
            //System.Console.WriteLine(cliente1.email);
            cliente1 = new cliente(Nome, Documento, Email, Endereco, Cep, Data_cadastro);
            CriarExcel(cliente1.Nome, cliente1.Documento, cliente1.Email, cliente1.Endereco, cliente1.Cep, cliente1.Data_cadastro);
        }
        static void salvar_dados_cliente() {

        }
    }
}
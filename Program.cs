using System;
using System.IO;
using NetOffice.ExcelApi;
using System.Text.RegularExpressions;


namespace sistema_de_concessionaria
{
    class Program
    {
        static void Main(string[] args)
        {
            inicializar();
            //cliente cliente1 = new cliente();
            //cliente1.validar_documento();
        }

        static void inicializar() {
            System.Console.WriteLine("------ Concessionaria ------");
            System.Console.WriteLine("Digite o que deseja fazer:");
            System.Console.WriteLine("1 - Cadastro de cliente");
            System.Console.WriteLine("2 - Cadastro de carros");
            System.Console.WriteLine("3 - Realizar venda de carro");
            System.Console.WriteLine("4 - listar carros vendidos no dia");
            System.Console.WriteLine("5 - Sair");
            string opcao = Console.ReadLine();
            while(opcao != "1" && opcao != "2" && opcao != "3" && opcao != "4" && opcao != "5") {
                 System.Console.WriteLine("------ Concessionaria ------");
                System.Console.WriteLine("Digite o que deseja fazer:");
                System.Console.WriteLine("1 - Cadastro de cliente");
                System.Console.WriteLine("2 - Cadastro de carros");
                System.Console.WriteLine("3 - Realizar venda de carro");
                System.Console.WriteLine("4 - listar carros vendidos no dia");
                System.Console.WriteLine("5 - Sair");
                opcao = Console.ReadLine();
            }
            switch(opcao) {
                case "1":
                    cadastrar_cliente();
                break;
                case "2":
                    //cadastrar_produto();
                break;
                case "3":
                    //vender();
                break;
                case "4":
                    
                break;
                case "5":
                  System.Environment.Exit(1);  
                break;
            }
        }
        //funcao para criar excel
        static void CriarExcel(string nome, string documento, string email, string endereco, string cep, string data_cadastro)
        {
            StreamWriter arq = new StreamWriter("clientes.csv",true);
            arq.WriteLine();
            arq.Write("{0};", nome);
            arq.Write("{0};", documento);
            arq.Write("{0};", email);
            arq.Write("{0};", endereco);
            arq.Write("{0};", cep);
            arq.Write("{0};", data_cadastro);
            arq.Close();
        }

        //metodo para criar o arquivo excel e escreve as informações do objeto
        
        static void verificar_cad(string Documento) {
            string[] linhas = System.IO.File.ReadAllLines(@"C:\Users\bruno afonso\Desktop\sistema de concessionaria\clientes.csv");
            string[] colunas;
            foreach (var linha in linhas) {
                colunas = linha.Split(";");
                if(Documento == colunas[1]) {
                    System.Console.WriteLine("Cliente ja existe, favor verificar os dados e tentar novamente.");
                    cadastrar_cliente();
                    break;
                }
            }
        }

        //metodo cadastrar o cliente, instanciando classe 
        static void cadastrar_cliente() {
            cliente cliente1 = new cliente("", "", "", "", "", "");
            string Nome = cliente1.validar_nome();
            //chamar o metodo de acordo com o tipo de documento
            string Documento = cliente1.validar_documento();
            verificar_cad(Documento);
            string Email = cliente1.validar_email();
            string Endereco = cliente1.validar_endereco();
            string Cep = cliente1.validar_cep();
            string Data_cadastro = cliente1.data_do_cadastro();
            //System.Console.WriteLine(cliente1.email);
            cliente1 = new cliente(Nome, Documento, Email, Endereco, Cep, Data_cadastro);
            CriarExcel(cliente1.Nome, cliente1.Documento, cliente1.Email, cliente1.Endereco, cliente1.Cep, cliente1.Data_cadastro);        
            System.Console.WriteLine("Gostaria de cadastrar um novo cliente ?");
            String cad = Console.ReadLine();
            while(cad != "sim" && cad != "SIM" && cad != "nao" && cad != "NAO") {
                System.Console.WriteLine("Gostaria de cadastrar um novo cliente ?");
                cad = Console.ReadLine();
            }
            if (cad == "SIM"  || cad == "sim") {
                cadastrar_cliente();
            } else {
                inicializar();
            }
        }
    }
}
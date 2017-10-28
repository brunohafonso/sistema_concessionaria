
using System;
using System.Text.RegularExpressions;
class cliente
{
    public string Nome{get; set;}
    public string Documento{get; set;}

    public string Email{get; set;}
    public string Endereco{get; set;}
    public string Cep{get; set;}

    public string Data_cadastro{get; set;}

    public cliente(string nome, string documento, string email, string endereco, string cep, string data_cadastro) {
        Nome = nome;
        Documento = documento;
        Email = email;
        Endereco = endereco;
        Cep = cep;
        Data_cadastro = data_cadastro;
    } 
    
    public string validar_nome() {
        System.Console.WriteLine("Digite o nome do cliente");
        string Nome = Console.ReadLine();
        while(Nome.Length < 1) {
            System.Console.WriteLine("Digite o nome do cliente");
            Nome = Console.ReadLine();
        }
        return Nome;
    }
    public string validar_documento() {
        System.Console.WriteLine("Digite o CPF ou CNPJ do cliente:");
        string documento = Console.ReadLine();
        while(documento.Length != 11 && documento.Length != 14) {
            System.Console.WriteLine("Documento invalido, digite o CPF ou CNPJ do cliente:");
            documento = Console.ReadLine();
        }
        if(documento.Length == 11) {
            return validar_cpf(documento);
        } else {
            return validar_cnpj(documento);
        }
    }
    public string validar_cnpj(string documento) {  
            do {
               try {
                    Convert.ToInt64(documento);
               } catch {
                   System.Console.WriteLine("Digite o numero do CNPJ do cliente novamente, CNPJ invalido.");
                   documento = Console.ReadLine();
               }
            } while(documento.Length > 14 || documento.Length < 14);
            int[] multiplicador1 = new int[]{5,4,3,2,9,8,7,6,5,4,3,2};
            int[] multiplicador2 = new int[]{6,5,4,3,2,9,8,7,6,5,4,3,2};
            string[] cnpj_char = new string[documento.Length];
            int[] cnpj_char_conv = new int[documento.Length];
            int[] mult_dig1 = new int[12];
            int[] mult_dig2 = new int[13];
            int sum_final = 0;
            int dig1 = 0;
            int dig2 = 0;
            for(var i = 0; i < multiplicador1.Length + 2; i++) {
                cnpj_char[i] = documento.Substring(i, 1);
                cnpj_char_conv[i] = Convert.ToInt32(cnpj_char[i]);
                //System.Console.WriteLine("{0}", cnpj_char_conv[i]); 
            }

            for(var i = 0; i < 12; i++) {
                mult_dig1[i] = cnpj_char_conv[i] * multiplicador1[i];
                sum_final = sum_final + mult_dig1[i];
            }
            dig1 = sum_final%11;
            if (dig1 < 2) {
                    dig1 = 0;
            } else {
                dig1 = 11 - dig1;
            }
            //System.Console.WriteLine(dig1);
            sum_final = 0;
            for(var i = 0; i < 13; i++) {
                mult_dig2[i] = cnpj_char_conv[i] * multiplicador2[i];
                sum_final = sum_final + mult_dig2[i];
            }
            dig2 = sum_final%11;
            if (dig2 < 2) {
                    dig2 = 0;
            } else {
                dig2 = 11 - dig2;
            }
            //System.Console.WriteLine(dig2);
            if (dig1 == cnpj_char_conv[12] && dig2 == cnpj_char_conv[13]) {
                    //System.Console.WriteLine("CPF valido");
                    string cnpj_mask = Convert.ToString(cnpj_char_conv[0]) + Convert.ToString(cnpj_char_conv[1]) + "." + Convert.ToString(cnpj_char_conv[2]) + Convert.ToString(cnpj_char_conv[3]) + Convert.ToString(cnpj_char_conv[4])+ "." + Convert.ToString(cnpj_char_conv[5]) + Convert.ToString(cnpj_char_conv[6]) + Convert.ToString(cnpj_char_conv[7]) + "/" + Convert.ToString(cnpj_char_conv[8]) + Convert.ToString(cnpj_char_conv[9]) + Convert.ToString(cnpj_char_conv[10]) + Convert.ToString(cnpj_char_conv[11]) + "." + Convert.ToString(cnpj_char_conv[12]) + Convert.ToString(cnpj_char_conv[13]); 
                    //System.Console.WriteLine(cnpj_mask);
                    return cnpj_mask;
            } else {
                //System.Console.WriteLine("CNPJ invalido");
                return "CNPJ invalido";
            }
        }
    public string validar_cpf(string documento) {
            do {
               try {
                    Convert.ToInt64(documento);
               } catch {
                   System.Console.WriteLine("Digite o numero do CPF do cliente novamente, CPF invalido.");
                   documento = Console.ReadLine();
               }
            } while(documento.Length > 11 || documento.Length < 11); 

                string[] cpf_char = new string[documento.Length];
                int[] cpf_char_conv = new int[documento.Length];
                for(var i = 0; i < documento.Length; i++) {
                cpf_char[i] = documento.Substring(i,1);
                cpf_char_conv[i] = Int32.Parse(cpf_char[i]);
                }
                int[] mult_dig1 = new int[9];
                int mult_sum = 0;
                int e = 10;
                int dig1;
                int dig2;
                for(var i = 0; i < 9; i++) {
                    mult_dig1[i] = cpf_char_conv[i] * e;
                    mult_sum = mult_sum + mult_dig1[i];
                    e--;
                }
                dig1 = mult_sum % 11;
                if (dig1 < 2) {
                    dig1 = 0;
                } else {
                    dig1 = 11 - dig1;
                }

                int[] mult_dig2 = new int[10];
                mult_sum = 0;
                e = 11;
                for(var i = 0; i < 10; i++) {
                    mult_dig2[i] = cpf_char_conv[i] * e;
                    mult_sum = mult_sum + mult_dig2[i];
                    e--;
                }
                dig2 = mult_sum % 11;
                if (dig2 < 2) {
                    dig2 = 0;
                } else {
                    dig2 = 11 - dig2;
                }
                if (dig1 == cpf_char_conv[9] && dig2 == cpf_char_conv[10]) {
                    //System.Console.WriteLine("CPF valido");
                    string cpf_mask = Convert.ToString(cpf_char_conv[0]) + Convert.ToString(cpf_char_conv[1]) + Convert.ToString(cpf_char_conv[2]) + "." + Convert.ToString(cpf_char_conv[3]) + Convert.ToString(cpf_char_conv[4]) + Convert.ToString(cpf_char_conv[5]) + "." + Convert.ToString(cpf_char_conv[6]) + Convert.ToString(cpf_char_conv[7]) + Convert.ToString(cpf_char_conv[8]) + "-" + Convert.ToString(cpf_char_conv[9]) + Convert.ToString(cpf_char_conv[10]);  
                    //System.Console.WriteLine(cpf_mask);
                    return cpf_mask;
                } else {
                    //System.Console.WriteLine("CPF invalido");
                    return "CPF invalido";
                }
    }
    public string validar_email() {
        System.Console.WriteLine("digite seu email");
        string email = Console.ReadLine();
        // mascara a ser aplicada
        Regex regex = new Regex(@"(\w+){1}(@)(\w+){2}(.)(\w+){2}");
        //texto a ser vaidado
        Match match = regex.Match(email);
        // loop para pedir dado caso incorreto
        while(!match.Success) {
            System.Console.WriteLine("digite seu email");
            email = Console.ReadLine();
            match = regex.Match(email);
        }
        //retornar valor
        //System.Console.WriteLine(match.Value);
        return match.Value; 
    }
    
    public string validar_endereco() {
        string[] endereco = new string[3];
        System.Console.WriteLine("Digite o endereco do cliente sem numero e sem complemento:");
        endereco[0] = Console.ReadLine();
        while(endereco[0].Length < 5) {
            System.Console.WriteLine("Endereco invalido, digite o endereco do cliente sem numero e sem complemento:");
            endereco[0] = Console.ReadLine();
        }
        System.Console.WriteLine("Digite o numero do endereco do cliente");
        endereco[1] = Console.ReadLine();
        do {
            try {
                Convert.ToInt32(endereco[1]);
            } catch {
                System.Console.WriteLine("Numero invalido, digite o numero do endereco do cliente");
                endereco[1] = Console.ReadLine();
            }
        } while(endereco[1].Length < 1);
        System.Console.WriteLine("O endereco do cliente possui complemento ?");
        System.Console.WriteLine("Digite 1 para sim ou Digite 2 para não.");
        string opcao = Console.ReadLine();
        while(opcao != "1" && opcao != "2") {
            System.Console.WriteLine("O endereco do cliente possui complemento ?");
            System.Console.WriteLine("opcao digitada invalida, digite 1 para sim ou Digite 2 para não.");
            opcao = Console.ReadLine();
        }
        if(opcao == "1") {
            System.Console.WriteLine("Digite o complemento do enreco do cliente:");
            endereco[2] = Console.ReadLine();
            while(endereco[2].Length < 1 ) {
                System.Console.WriteLine("Digite o complemento do enreco do cliente:");
                endereco[2] = Console.ReadLine();
            }
        } 
        return endereco[0] + " " + endereco[1] + " " + endereco[2];
    }
    public string validar_cep() {
        System.Console.WriteLine("digite seu cep com traço: EX.: 12345-123");
        string cep = Console.ReadLine();
        // mascara a ser verificada numeros nas chaves = quantidade de digitos
        Regex regex = new Regex(@"(\d+){5}(-)(\d+){3}");
        // verifica de o dado digitado confere com a mascara
        Match match = regex.Match(cep);
        while(!match.Success) {
            System.Console.WriteLine("Cep invalido, digite novamente seu cep com traço: EX.: 12345-123");
            cep = Console.ReadLine();
            match = regex.Match(cep);
        }
        //System.Console.WriteLine(match.Value);
        return match.Value;
    }
    public string data_do_cadastro() {
        return Convert.ToString(DateTime.Now);
    }
}


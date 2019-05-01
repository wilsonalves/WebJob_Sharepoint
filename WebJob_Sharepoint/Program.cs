using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebJob_Sharepoint
{
    class Program
    {
        static void Main(string[] args)
        {

            AuthenticationManager authenticationManager = new AuthenticationManager();

            string url = ConfigurationManager.AppSettings["url"];
            string usuario = ConfigurationManager.AppSettings["usuario"];
            string senha = ConfigurationManager.AppSettings["senha"];

            using (ClientContext context = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(url, usuario, senha))
            {

                CamlQuery query = new CamlQuery();
                List listDespesas = context.Site.RootWeb.GetListByTitle("Despesas");
                ListItemCollection itemsDespesas = listDespesas.GetItems(query);


                context.Load<List>(listDespesas);
                context.Load<ListItemCollection>(itemsDespesas);
                context.ExecuteQuery();


                foreach (var item in itemsDespesas)

                {
                    FieldLookupValue reembolsoID = (FieldLookupValue)item["Reembolso"];
                    double valor = (double)item["Valor"];

                    try
                    {
                        UpdateItem(reembolsoID.LookupId, valor, context);
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("Erro Ao Atualizar item " + ex);
                    }


                }
                Console.WriteLine("Todos os itens foram atualizados");



            }
        }

        public static void UpdateItem(int reembolsoiD, double valor, ClientContext context)
        {
            List listaReembolso = context.Site.RootWeb.GetListByTitle("Reembolsos");


            ListItem itemReembolso = listaReembolso.GetItemById(reembolsoiD);
            context.Load<List>(listaReembolso);
            context.Load<ListItem>(itemReembolso);
            context.ExecuteQuery();
            string status = (string)itemReembolso["Status"].ToString().ToLower();

            if (status == "pendente")
            {

                double valorAtual = Convert.ToDouble(itemReembolso["Total"]);

                itemReembolso["Total"] = valorAtual + valor;
                itemReembolso.Update();

                context.ExecuteQuery();
            }

        }
    }
}



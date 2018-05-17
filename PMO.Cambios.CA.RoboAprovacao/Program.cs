using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using PMO.Cambios.CA.RoboAprovacao.ListaData;
using System.Threading.Tasks;
using System.Collections;
using ActiveUp.Net.Mail;
using PMO.Cambios.CA.RoboAprovacao;

namespace PMO.Cambios.CA.RoboAprovacao
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("----- PORTAL PCFTI Câmbios - ROBÔ APROVAÇÃO -----");
            Console.WriteLine("\n\n");
            Console.WriteLine("====== NÃO FECHAR!!!! ======");

            int vPosInicio;
            int vPosFim;
            int vIntervalo;

            Config itemConfig = new Config();
            //SolicitaçãoDeAdiantamentoDeFornecedoresItem lstItem = null;

            PMODataContext dc = new PMODataContext(new Uri(itemConfig.uri));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

            //Verifica na caixa de email os emails de resposta das aprovações.
            MailRepository rep = new MailRepository("imap.gmail.com", 993, true, @"b2wsistemas@gmail.com", "b2w@123456");

            foreach (Message email in rep.GetUnreadMails("Inbox"))
            {
                if (email.Subject == "RE:Adiantamento de Fornecedores - Análise de solicitação")
                {
                    //Guarda os parâmetros enviados pela resposta de aprovação.

                    //ID do item
                    vPosInicio = email.BodyText.Text.IndexOf("[IDITEM:") + 8;
                    vPosFim = email.BodyText.Text.IndexOf("]", vPosInicio);
                    vIntervalo = vPosFim - vPosInicio;
                    int vIDItem = int.Parse(email.BodyText.Text.Substring(vPosInicio, vIntervalo));

                    //Tipo de Aprovação (TI/FIN)
                    vPosInicio = email.BodyText.Text.IndexOf("[TIPO:") + 6;
                    vPosFim = email.BodyText.Text.IndexOf("]", vPosInicio);
                    vIntervalo = vPosFim - vPosInicio;
                    string vTipo = email.BodyText.Text.Substring(vPosInicio, vIntervalo);

                    //Resposta da Aprovação (S/N)
                    vPosInicio = email.BodyText.Text.IndexOf("[APROVA:") + 8;
                    vPosFim = email.BodyText.Text.IndexOf("]", vPosInicio);
                    vIntervalo = vPosFim - vPosInicio;
                    string vAprova = email.BodyText.Text.Substring(vPosInicio, vIntervalo);
                    if (vAprova == "S")
                        vAprova = "Sim";
                    else
                        vAprova = "Não";
                    
                    //Responsável pela aprovação
                    vPosInicio = email.BodyText.Text.IndexOf("[RESP:") + 6;
                    vPosFim = email.BodyText.Text.IndexOf("]", vPosInicio);
                    vIntervalo = vPosFim - vPosInicio;
                    string vResponsavel = email.BodyText.Text.Substring(vPosInicio, vIntervalo);

                    //Resposta de orçamento
                    vPosInicio = email.BodyText.Text.IndexOf("[ORC:") + 5;
                    vPosFim = email.BodyText.Text.IndexOf("]", vPosInicio);
                    vIntervalo = vPosFim - vPosInicio;
                    string vOrcamento = email.BodyText.Text.Substring(vPosInicio, vIntervalo);
                    if (vOrcamento == "NOK")
                        vOrcamento = "Não OK";
                    
                    //Atualizar item na lista
                    var itemEdit = (from lstSolic in dc.SolicitaçãoDeAdiantamentoDeFornecedores where lstSolic.ID == vIDItem select lstSolic).First();
                    
                    itemEdit.Orçamento.Value = vOrcamento;
                    if (vTipo == "TI")
                    {
                        itemEdit.AprovadoTI.Value = vAprova;
                        itemEdit.ResponsávelTI = vResponsavel;
                        itemEdit.DataTI = DateTime.Today;
                    }

                    if (vTipo == "FIN")
                    {
                        itemEdit.AprovadoFinanceiro.Value = vAprova;
                        itemEdit.ResponsávelFinanceiro = vResponsavel;
                        itemEdit.DataFinanceiro = DateTime.Today;
                    }

                    dc.UpdateObject(itemEdit);
                    dc.SaveChanges();

                    //=====
                    
                    Console.WriteLine("Dados Atualizados!");
                    Console.ReadLine();
                    //Environment.Exit(1);
                }
            }
        }
    }
}

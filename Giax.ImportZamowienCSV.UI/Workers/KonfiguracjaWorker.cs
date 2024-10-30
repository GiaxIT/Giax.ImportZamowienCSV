using System;
using Soneta.Business;
using Soneta.Business.UI;
using Giax.ImportZamowienCSV.UI.Workers;
using Soneta.Handel;
using Soneta.Business.Db;
using System.Linq;

[assembly: Worker(typeof(KonfiguracjaWorker), typeof(DokHandlowe))]
namespace Giax.ImportZamowienCSV.UI.Workers
{
    public class KonfiguracjaWorker
    {

        [Context]
        public Session Session { get; set; }


        [Action("Giax/Importuj zamowienia Amazon CSV/Konfiguracja", Icon = ActionIcon.Fix, Mode = ActionMode.SingleSession | ActionMode.ConfirmSave | ActionMode.Progress)]
        public MessageBoxInformation Konfiguracja()
        {
            using (Session ss = Session.Login.CreateSession(false, true))
            {
                SprawdzCechy(ss);

            }
            return new MessageBoxInformation("Sukces", "Skonfigurowano pomyślnie!");
        }

        (string, FeatureReadOnlyMode, object, string, FeatureTypeNumber)[] cechy =
       {
                ("Giax_ImportAmazon", FeatureReadOnlyMode.Standard, false,"Kontrahenci",FeatureTypeNumber.Bool),

        };

        public void SprawdzCechy(Session ses)

        {
            using (var trans = ses.Logout(true))
            {
                var bmodule = ses.GetBusiness();
                foreach (var p in cechy)
                {
                    var cecha = bmodule.FeatureDefs.Rows.FirstOrDefault(fd => ((FeatureDefinition)fd).Name == p.Item1);
                    if (cecha == null)
                    {
                        var fd = new FeatureDefinition(p.Item4);
                        fd.TypeNumber = p.Item5;
                        fd.Name = p.Item1;
                        fd.ReadOnlyMode = p.Item2;
                        fd.InitValue = p.Item3;
                        bmodule.FeatureDefs.AddRow(fd);
                    }
                }

                trans.CommitUI();
            }
            ses.Save();
        }
    }


}

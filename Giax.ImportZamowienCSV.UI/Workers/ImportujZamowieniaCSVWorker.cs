using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using Giax.ImportZamowienCSV.UI.Model;
using Giax.ImportZamowienCSV.UI.Workers;
using Soneta.Business;
using Soneta.Business.UI;
using Soneta.Handel;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Extensions.DependencyInjection;
using Soneta.CRM;
using Soneta.Towary;
using Soneta.Types;
using Soneta.Kadry;
using static Giax.ImportZamowienCSV.UI.Workers.ImportujZamowieniaCSVWorkerParams;
using Soneta.Magazyny;
using System.Reflection.Emit;
using System.Reflection;
using Soneta.Business.Licence.UI;
using Soneta.Business.Licence;
using Soneta.Business.App;
using DocumentFormat.OpenXml.Presentation;
using static Soneta.Place.WypElementNadgodziny;
using System.Configuration;
using Soneta.Kasa;
using Soneta.Core;
using ServiceStack;
using CsvHelper.Configuration.Attributes;
using Soneta.Tools;
using Soneta.Data.QueryDefinition;
using System.Threading;

[assembly: Worker(typeof(ImportujZamowieniaCSVWorker), typeof(DokHandlowe))]

namespace Giax.ImportZamowienCSV.UI.Workers
{
    public class ImportujZamowieniaCSVWorker
    {
        [Context]
        public Session Session { get; set; }

        [Context]
        public ImportujZamowieniaCSVWorkerParams @params { get; set; }

       


        [Action("Giax/Importuj zamowienia Amazon CSV", Icon = ActionIcon.ArrowUp, Mode = ActionMode.SingleSession | ActionMode.ConfirmSave | ActionMode.Progress)]
      
        public object CSV()
        {
            int added_positions_count = 0;
            int added_orders_count = 0;
            bool czy_kontrahent = false;

            var fs = Session.GetService<IFileSystemService>();
            var data = fs.ReadStream(@params.Plik);

           
           
            string filepath = @params.FilePath;

            List<Pozycja> pozycje;

            try
            {
                pozycje = ReadCSVFile(data);
            }
            catch (Exception ex)
            {
                return new MessageBoxInformation("Błąd", $"Wystąpił błąd podczas odczytu pliku CSV: {ex.Message}");
            }

            List<string> numery_zamowien = GetUniqueOrderNumbers(pozycje);

            List<DokumentHandlowy> zamowienia = new List<DokumentHandlowy>();

            using (var t = Session.Logout(true))
            {
                TraceInfo.HideProgressWindow();
                TraceInfo.ShowProgressWindow();
                for (int i = 0; i < numery_zamowien.Count; i++)
                {
                    
                    foreach (var numer in numery_zamowien)
                    {
                        TraceInfo.WriteProgress("Obiekt: " + i + " (pozostało: " + numery_zamowien.Count + ")");
                    TraceInfo.SetProgressBar(new Percent(i + 1, numery_zamowien.Count));
                        List<Pozycja> filtrowane_pozycje = pozycje.Where(p => p.NumerZamowieniaPO == numer && !p.Dostepnosc.Contains("Anulowano")).ToList();

                        if (!filtrowane_pozycje.Any())
                            continue;

                        added_positions_count += filtrowane_pozycje.Count;
                        added_orders_count++;

                        DokumentHandlowy dokument = new DokumentHandlowy();
                        HandelModule.GetInstance(Session).DokHandlowe.AddRow(dokument);

                        dokument.Definicja = HandelModule.GetInstance(Session).DefDokHandlowych.WgSymbolu[@params.NazwaDefDok];
                        dokument.Obcy.Numer = numer;
                        dokument.Magazyn = HandelModule.GetInstance(Session).Magazyny.Magazyny.WgNazwa[@params.NazwaMag];

                        var pierwszaPozycja = filtrowane_pozycje.First();
                        var lokzalizacja = pierwszaPozycja.Lokalizacja.Substring(0, 4);
                       // var lokzalizacja = "XS";
                        var crmmodule = CRMModule.GetInstance(Session);
                        var lokalziacje_kont = crmmodule.Lokalizacje.CreateView().ToList();

                        foreach (Lokalizacja lok in lokalziacje_kont)
                        {
                            if (lok.Nazwa.Contains(lokzalizacja))
                            {
                                dokument.Kontrahent = (Kontrahent)lok.Kontrahent;
                                dokument.OdbiorcaMiejsceDostawy = lok;
                                czy_kontrahent = true;
                                break;
                            }
                        }

                        if (!czy_kontrahent)
                            return new MessageBoxInformation("Błąd", $"Nie znaleziono kontrahenta dla lokalizacji: {lokzalizacja}");

                        dokument.Data = Date.Parse(filtrowane_pozycje.First().DataZamowienia);
                        dokument.DataOtrzymania = Date.Parse(filtrowane_pozycje.First().DataOtrzymania);



                        if (@params.CzyZaakceptowany) dokument.Potwierdzenie = PotwierdzenieDokumentuHandlowego.Zaakceptowany;
                        if (@params.CzyZatwierdzony) dokument.Potwierdzenie = PotwierdzenieDokumentuHandlowego.Potwierdzony;

                        var towar_modue = TowaryModule.GetInstance(Session).Towary;
                        foreach (var poz in filtrowane_pozycje)
                        {
                            var pozycjaDokHandlowego = Session.AddRow(new PozycjaDokHandlowego(dokument));

                            Towar towar;
                            try
                            {
                               towar = towar_modue.WgEAN[poz.EAN].First();

                            }
                            catch (Exception ex)
                            {
                                return new MessageBoxInformation("Błąd", $"Nie znaleziono towaru dla EAN: {poz.EAN}");
                            }

                            //towar = TowaryModule.GetInstance(Session).Towary.WgEAN["5901035500211"].First();
                            pozycjaDokHandlowego.Towar = towar;
                            pozycjaDokHandlowego.Ilosc = new Quantity(poz.Ilosc, pozycjaDokHandlowego.Towar.Jednostka.Kod);
                            pozycjaDokHandlowego.Cena = new DoubleCy(poz.KosztJednostkowy);
                        }

                        if (!@params.CzyBufor) dokument.Stan = StanDokumentuHandlowego.Zatwierdzony;
                        dokument.Features["Dane kuriera"] = "UPS";
                        dokument.Features["Data_Dostawy_Zam"] = Date.Parse(filtrowane_pozycje.First().DataZamowienia).AddDays(2);
                        zamowienia.Add(dokument);

                       
                    }


                }
                t.CommitUI();
            }
            return new MessageBoxInformation("Sukces", $"Zaimportowane zamówienia {numery_zamowien.Count}");

        }



        public List<Pozycja> ReadCSVFile(object filePath)
        {
            var pozycje = new List<Pozycja>();
            var fs = Session.GetService<IFileSystemService>();
            var data = fs.ReadStream(@params.Plik);

            using (var reader = new StreamReader(data))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Read();
                csv.ReadHeader();

                var secondColumnHeader = csv.HeaderRecord[1];

                string poColumn = "PO";
                string vendorColumn = secondColumnHeader == "Vendor" ? "Vendor" : "Dostawca";
                string warehouseColumn = secondColumnHeader == "Vendor" ? "Warehouse" : "Wysyłka do lokalizacji";
                string externalIdColumn = secondColumnHeader == "Vendor" ? "External ID" : "Identyfikator zewnętrzny";
                string acceptedQuantityColumn = secondColumnHeader == "Vendor" ? "Accepted Quantity" : "Zaakceptowana ilość";
                string unitCostColumn = secondColumnHeader == "Vendor" ? "Unit Cost" : "Koszt jednostkowy";
                string totalCostColumn = secondColumnHeader == "Vendor" ? "Total Cost" : "Całkowity koszt";
                string windowStartColumn = secondColumnHeader == "Vendor" ? "Window Start" : "Data początkowa przedziału czasowego";
                string windowEndColumn = secondColumnHeader == "Vendor" ? "Window End" : "Data końcowa przedziału czasowego";
                string availabilityColumn = secondColumnHeader == "Vendor" ? "Availability" : "Dostępność";

                while (csv.Read())
                {
                    var iloscString = csv.GetField<string>(acceptedQuantityColumn);
                    int ilosc = 0;
                    if (!string.IsNullOrWhiteSpace(iloscString))
                    {
                        int.TryParse(iloscString, NumberStyles.Any, CultureInfo.InvariantCulture, out ilosc);
                    }

                    var kosztJednostkowyString = csv.GetField<string>(unitCostColumn);
                    double kosztJednostkowy = 0.0;
                    if (!string.IsNullOrWhiteSpace(kosztJednostkowyString))
                    {
                        double.TryParse(kosztJednostkowyString, NumberStyles.Any, CultureInfo.InvariantCulture, out kosztJednostkowy);
                    }

                    var calkowityKosztString = csv.GetField<string>(totalCostColumn);
                    double calkowityKoszt = 0.0;
                    if (!string.IsNullOrWhiteSpace(calkowityKosztString))
                    {
                        double.TryParse(calkowityKosztString, NumberStyles.Any, CultureInfo.InvariantCulture, out calkowityKoszt);
                    }

                    var pozycja = new Pozycja
                    {
                        EAN = csv.GetField<string>(externalIdColumn),
                        Ilosc = ilosc,
                        KosztJednostkowy = kosztJednostkowy,
                        NumerZamowieniaPO = csv.GetField<string>(poColumn),
                        DataZamowienia = csv.GetField<string>(windowStartColumn),
                        Dostepnosc = csv.GetField<string>(availabilityColumn),
                        Lokalizacja = csv.GetField<string>(warehouseColumn),
                        DataOtrzymania = csv.GetField<string>(windowEndColumn)
                    };

                    pozycje.Add(pozycja);
                }
            }

            return pozycje;
        }



        static List<string> GetUniqueOrderNumbers(List<Pozycja> pozycje)
        {

            



            return pozycje.Select(p => p.NumerZamowieniaPO).Distinct().ToList();
        }
    }

    public class ImportujZamowieniaCSVWorkerParams : ContextBase
    {
        private string V = "";

        private bool _czyPotwierdzony = false;
        private bool _czyZaakceptowany = false;
        private bool _czyBufor = true;
        private string _nazwaMagazynu = "Magazyn sprzedaży";
        private string _symbolDokumentu = "ZO";

        
        public string Plik { get; set; }

        public object GetListPlik()
        {
            return new FileDialogInfo
            {
                Title = "Wybierz plik",
                DefaultExt = ".csv",
                ForbidMultiSelection = true,
                InitialDirectory = @"C:\"
            };
        }

        public ImportujZamowieniaCSVWorkerParams(Context context) : base(context)
        {
         
        }



        public Object GetList()
        {
            return new FileDialogInfo { Title = "Wybierz pliki z wyciągami", ForbidMultiSelection = false }.AddAllFilesFilter();
        }

       
        public string FilePath
        {
            get { return V; }
            set { V = value; }
        }

        public bool CzyZatwierdzony
        {
            get { return _czyPotwierdzony; }
            set
            {
                _czyPotwierdzony = value;
                _czyZaakceptowany = !_czyPotwierdzony;
                _czyBufor = !_czyPotwierdzony;
            }
        }

        public bool CzyZaakceptowany
        {
            get { return _czyZaakceptowany; }
            set
            {
                _czyZaakceptowany = value;
                _czyPotwierdzony = !_czyZaakceptowany;
                _czyBufor =!_czyZaakceptowany;
            }
        }

        public bool CzyBufor
        {
            get { return _czyBufor; }
            set { 
                _czyBufor = value; 
                _czyZaakceptowany= !_czyBufor;
                _czyPotwierdzony= !_czyBufor;             
            }
        }

        public string NazwaDefDok
        {
            get { return _symbolDokumentu; }
            set { _symbolDokumentu = value; }
        }

        public string NazwaMag
        {
            get { return _nazwaMagazynu; }
            set { _nazwaMagazynu = value; }
        }
    }
}

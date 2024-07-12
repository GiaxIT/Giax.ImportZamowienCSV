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
using Soneta.Business;
using Soneta.Business.UI;
using Soneta.CRM;
using Soneta.Handel;
using Soneta.Towary;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ServiceStack.Text;
using ServiceStack;
using Soneta.Kadry;

[assembly: Worker(typeof(ImportujZamowieniaCSVWorker), typeof(DokHandlowe))]

namespace Giax.ImportZamowienCSV.UI.Workers
{
    public class ImportujZamowieniaCSVWorker
    {

        [Context]
        public Session Session { get; set; }

        [Context]
        public ImportujZamowieniaCSVWorkerParams @params
        {
            get;
            set;
        }


        [Action("Importuj zamowienia Amazonu/CSV", Mode = ActionMode.SingleSession | ActionMode.ConfirmSave | ActionMode.Progress)]
        public MessageBoxInformation CSV()
        {


            int added_positions_count = 0;
            int added_orders_count = 0;
            bool czy_kontrahent = false;
            //#1 Import pliku
          
            string filepath = @params.FilePath;

            if (!File.Exists(filepath))
            {
                return new MessageBoxInformation("Błąd", "Plik nie istnieje.");
            }

            List<Pozycja> pozycje;

            try
            {
                pozycje = ReadCSVFile(@params.FilePath);
               
            }
            catch (Exception ex)
            {
                return new MessageBoxInformation("Błąd", $"Wystąpił błąd podczas odczytu pliku CSV: {ex.Message}");
            }


            // #2 Utworzenie obiektów handlowych

            List<string> numery_zamowien = GetUniqueOrderNumbers(pozycje);
           // using (var session = Session.Login.CreateSession(true, false))
            using (var t = Session.Logout(true))
            {
                foreach (var numer in numery_zamowien)
                {
                    // wyfiltrowac wszystkie pozycje dla konkretnego zamowienia 
                    
                    List<Pozycja> filtrowane_pozycje = pozycje.Where(p => p.NumerZamowieniaPO == numer)
                                                              .Where(p => !p.Dostepnosc.Contains("Anulowano"))
                                                              .ToList();

                    added_positions_count += filtrowane_pozycje.Count();
                    added_orders_count++;

                    DokumentHandlowy dokument = new DokumentHandlowy();
                    HandelModule.GetInstance(Session).DokHandlowe.AddRow(dokument);
                    

                    dokument.Definicja = HandelModule.GetInstance(Session).DefDokHandlowych.ZamówienieOdbiorcy;
                    dokument.Obcy.Numer = numer;
                    
                    //na teraz
                    dokument.Magazyn = HandelModule.GetInstance(Session).Magazyny.Magazyny.WgNazwa["Firma"];
                    //dokument.Magazyn = HandelModule.GetInstance(Session).Magazyny.Magazyny.WgNazwa["Magazyn sprzedaży"];

                    //dodanie kontrahenta po kraju wysyłki
                    var pierwszaPozycja = filtrowane_pozycje.First();
                    //var lokalizacja1 = pierwszaPozycja.Lokalizacja.Substring(0, 4);
                    var lokzalizacja = "SZ01";
                    var crmmodule = CRMModule.GetInstance(Session);
                    var kontrahenci = crmmodule.Kontrahenci.CreateView().ToList();


                    foreach(Kontrahent kont in kontrahenci)
                    {
                        var sa = kont.Lokalizacje.FirstOrDefault();
                        if (sa != null && sa.Kod == lokzalizacja)
                        {
                            dokument.Kontrahent = kont;
                            dokument.OdbiorcaMiejsceDostawy = sa;
                            czy_kontrahent = true;
                            break;
                            
                        }
                    }

                    if(!czy_kontrahent) return new MessageBoxInformation("Błąd", $"Nie znaleziono kontrahenta dla lokalizacji: {lokzalizacja}");

                    dokument.Data = Date.Parse(pozycje.FirstOrDefault().DataZamowienia);
                    dokument.DataOtrzymania = Date.Parse(pozycje.FirstOrDefault().DataOtrzymania);


                    if(@params.CzyZaakceptowany) dokument.Potwierdzenie = PotwierdzenieDokumentuHandlowego.Zaakceptowany; 
                    if(@params.CzyZatwierdzony) dokument.Potwierdzenie = PotwierdzenieDokumentuHandlowego.Potwierdzony;
                    

                    //dodanie pozycji do dokumentu
                    foreach (var poz in filtrowane_pozycje)
                    {
                        var pozycjaDokHandlowego = Session.AddRow(new PozycjaDokHandlowego(dokument));

                        //na demo
                        var towar = TowaryModule.GetInstance(Session).Towary.WgEAN["5901035500211"].First();
                        //var towar = TowaryModule.GetInstance(Session).Towary.WgEAN[poz.EAN].First();

                        pozycjaDokHandlowego.Towar = towar;
                        pozycjaDokHandlowego.Ilosc = new Quantity(poz.Ilosc, pozycjaDokHandlowego.Towar.Jednostka.Kod);
                        pozycjaDokHandlowego.Cena = new DoubleCy(poz.KosztJednostkowy);

                        Session.Events.Invoke();
                    }

                    if (!@params.CzyBufor) dokument.Stan = StanDokumentuHandlowego.Zatwierdzony;


                }
                t.Commit();
                
            }


            return new MessageBoxInformation("Import CSV")
            {
                Text = "Pomyślnie zaimportowano: "+added_orders_count+ " zamowień i " + added_positions_count + " pozycji.",
                OKHandler = () =>
                {
                    using (var t = @params.Session.Logout(true))
                    {
                        t.Commit();
                    }
                    return "Operacja została zakończona";
                }
            };

        }

        public List<Pozycja> ReadCSVFile(string filePath)
        {
            
                var pozycje = new List<Pozycja>();

                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvHelper.CsvReader(reader, CultureInfo.InvariantCulture))
                {
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                    {
                    var iloscString = csv.GetField<string>("Zaakceptowana ilość");
                    int ilosc = 0;
                    if (!string.IsNullOrWhiteSpace(iloscString))
                    {
                        int.TryParse(iloscString, NumberStyles.Any, CultureInfo.InvariantCulture, out ilosc);
                    }

                    var kosztJednostkowyString = csv.GetField<string>("Koszt jednostkowy");
                    double kosztJednostkowy = 0.0;
                    if (!string.IsNullOrWhiteSpace(kosztJednostkowyString))
                    {
                        double.TryParse(kosztJednostkowyString, NumberStyles.Any, CultureInfo.InvariantCulture, out kosztJednostkowy);
                    }

                    var calkowityKosztString = csv.GetField<string>("Całkowity koszt");
                    double calkowityKoszt = 0.0;
                    if (!string.IsNullOrWhiteSpace(calkowityKosztString))
                    {
                        double.TryParse(calkowityKosztString, NumberStyles.Any, CultureInfo.InvariantCulture, out calkowityKoszt);
                    }


                    var pozycja = new Pozycja
                    {
                        EAN = csv.GetField<string>("Identyfikator zewnętrzny"),
                        Ilosc = ilosc,
                        KosztJednostkowy = kosztJednostkowy,
                        NumerZamowieniaPO = csv.GetField<string>("PO"),
                        DataZamowienia = csv.GetField<string>("Data początkowa przedziału czasowego"),
                        Dostepnosc = csv.GetField<string>("Dostępność"),
                        Lokalizacja = csv.GetField<string>("Wysyłka do lokalizacji"),
                        DataOtrzymania = csv.GetField<string>("Data końcowa przedziału czasowego")
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
        private string V = "C:\\Users\\it01.DOMENA\\Downloads\\PurchaseOrderItems.csv";
        private bool _czyPotwierdzony = false;
        private bool _czyZaakceptowany = false;
        private bool _czyBufor = true;

        public ImportujZamowieniaCSVWorkerParams(Context context) : base(context)
        {
            AddRequiredVerifierForProperty(nameof(FilePath));

        }

        
        [Required]
        public string FilePath
        {
            get
            {
                return V;
            }

            set
            {
                V = value;
            }
        }

        public bool CzyZatwierdzony
        {
            get
            {
                return _czyPotwierdzony;
            }

            set
            {
                _czyPotwierdzony = value;
                _czyZaakceptowany = !_czyPotwierdzony; 
            }
        }

        public bool CzyZaakceptowany
        {
            get
            {
                return _czyZaakceptowany;
            }

            set
            {
                _czyZaakceptowany = value;
                _czyPotwierdzony = !_czyZaakceptowany;

            }
        }


        public bool CzyBufor
        {
            get
            {
                return _czyBufor;
            }

            set
            {
                _czyBufor = value;
              
            }
        }


    }

}

using System.Collections.Generic;

namespace Giax.ImportZamowienCSV.UI.Model
{

    public class Pozycja
    {
        public string EAN { get; set; }
        public int Ilosc { set; get; }

        public double KosztJednostkowy {  set; get; }
        public string DataZamowienia { get; set; }

        public string NumerZamowieniaPO { get; set; }

        public string Dostepnosc {  set; get; }

        public string Lokalizacja {  set; get; }

        public string DataOtrzymania { get; set; }
    }


}

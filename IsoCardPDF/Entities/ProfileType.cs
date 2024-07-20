using System.ComponentModel;

namespace IsoCardPDF.Entities
{
    public enum ProfileType
    {
        LA3,
        LA5, // wybijane 
        LA6,
        LAX, // frezowane
        VA4,
        [Description("LA3-LENS")]
        LA3LENS,
        [Description("LAX-LENS")]
        LAXLENS,
        [Description("LA1-Stara")]
        LA1Stara, // bez dalszej obróbki
        LAQ // do malowania
    }
}

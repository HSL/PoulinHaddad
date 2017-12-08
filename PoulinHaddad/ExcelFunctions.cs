using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using static System.Math;

namespace PoulinHaddad
{
  public class ExcelFunctions
  {
    [ExcelFunction(Category = "Poulin Haddad", Description = "Estimate log Pvo:w from log Po:w")]
    public static object LogPvow(
      [ExcelArgument("Log octanol:water partition coefficient")] double logPow,
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("Does drug chemical structure consist of at least one oxygen atom? (TRUE|FALSE)")] bool containsOxygen
      )
    {
      var isValidArg = Enum.TryParse(@class, out IonizationClass ionizationClass);
      if (!isValidArg) return ExcelError.ExcelErrorValue;

      return LogPvow(logPow, ionizationClass, containsOxygen);
    }

    [ExcelFunction(Category = "Poulin Haddad", Description = "Ionization term (tissue)")]
    public static object Iwt(
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("pKa")] object pKa,
      [ExcelArgument(Name = "pKa.base", Description = "pKa.base (zwitterion only)")] object pKa_base
      )
    {
      return Iw(@class, pKa, pKa_base, 7.0);
    }

    [ExcelFunction(Category = "Poulin Haddad", Description = "Ionization term (plasma or extracellular water)")]
    public static object Iwp(
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("pKa")] object pKa,
      [ExcelArgument(Name = "pKa.base", Description = "pKa.base (zwitterion only)")] object pKa_base
      )
    {
      return Iw(@class, pKa, pKa_base, 7.4);
    }

    [ExcelFunction(Category = "Poulin Haddad", Description = "Unbound fraction in plasma")]
    public static object fup(
      [ExcelArgument("Log octanol:water partition coefficient")] double logPow,
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("pKa")] object pKa,
      [ExcelArgument(Name = "pKa.base", Description = "pKa.base (zwitterion only)")] object pKa_base
      )
    {
      return fup(logPow, @class, pKa, pKa_base, false);
    }

    private static object fup(
      double logPow,
      string @class,
      object pKa,
      object pKa_base,
      bool denominatorOnly
      )
    {
      var Pnlw = Pow(10, logPow);

      var o = ExcelFunctions.Iwp(@class, pKa, pKa_base);
      if (o is ExcelError) return o;
      var Iwp = (double)o;

      var Fwp = _table1["plasma"]["Fw"];
      var Fnlp = _table1["plasma"]["Fnl"];

      var denominator = ((1 + Iwp) * Fwp + Pnlw * Fnlp);

      if (denominatorOnly) return denominator;

      var fup = (1 + Iwp) / denominator;

      return fup;
    }

    [ExcelFunction(Category = "Poulin Haddad", Description = "Unbound fraction in tissue")]
    public static object fut(
      [ExcelArgument("Tissue (adipose|bone|brain|gut|heart|kidneys|liver|lungs|muscle|skin|spleen|thymus|blood cells")] string tissue,
      [ExcelArgument("Log octanol:water partition coefficient")] double logPow,
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("Does drug chemical structure consist of at least one oxygen atom? (TRUE|FALSE)")] bool containsOxygen,
      [ExcelArgument("pKa")] object pKa,
      [ExcelArgument(Name = "pKa.base", Description = "pKa.base (zwitterion only)")] object pKa_base
      )
    {
      return fut(tissue, logPow, @class, containsOxygen, pKa, pKa_base, false);
    }

    private static object fut(
      string tissue,
      double logPow,
      string @class,
      bool containsOxygen,
      object pKa,
      object pKa_base,
      bool denominatorOnly
      )
    {
      if (string.IsNullOrWhiteSpace(tissue)) return ExcelError.ExcelErrorValue;
      tissue = tissue.ToLowerInvariant();
      if (tissue == "plasma") return ExcelError.ExcelErrorValue;
      if (!_table1.ContainsKey(tissue)) return ExcelError.ExcelErrorValue;

      object o;
      double Pnlw;
      if (tissue == "adipose")
      {
        o = LogPvow(logPow, @class, containsOxygen);
        if (o is ExcelError) return o;
        var logPvow = (double)o;
        Pnlw = Pow(10, logPvow);
      }
      else
      {
        Pnlw = Pow(10, logPow);
      }

      o = ExcelFunctions.Iwt(@class, pKa, pKa_base);
      if (o is ExcelError) return o;
      var Iwt = (double)o;

      var Fwt = _table1[tissue]["Fw"];
      var Fnlt = _table1[tissue]["Fnl"];

      var denominator = ((1 + Iwt) * Fwt + Pnlw * Fnlt);

      if (denominatorOnly) return denominator;

      var fut = (1 + Iwt) / denominator;

      return fut;
    }

    [ExcelFunction(Category = "Poulin Haddad", Description = "Tissue-plasma ratio")]
    public static object Kp(
      [ExcelArgument("Tissue (adipose|bone|brain|gut|heart|kidneys|liver|lungs|muscle|skin|spleen|thymus|blood cells")] string tissue,
      [ExcelArgument("Log octanol:water partition coefficient")] double logPow,
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("Does drug chemical structure consist of at least one oxygen atom? (TRUE|FALSE)")] bool containsOxygen,
      [ExcelArgument("pKa")] object pKa,
      [ExcelArgument(Name = "pKa.base", Description = "pKa.base (zwitterion only)")] object pKa_base
      )
    {
      object o;

      o = ExcelFunctions.fup(logPow, @class, pKa, pKa_base, true);
      if (o is ExcelError) return o;
      var fup = (double)o;

      o = ExcelFunctions.fut(tissue, logPow, @class, containsOxygen, pKa, pKa_base, true);
      if (o is ExcelError) return o;
      var fut = (double)o;

      var Kp = fut / fup;

      return Kp;
    }

    [ExcelFunction(Category = "Poulin Haddad", Description = "Volume of distribution at steady state")]
    public static object Vss(
      [ExcelArgument("Log octanol:water partition coefficient")] double logPow,
      [ExcelArgument("Ionization class (N|WB|SB|A|Z)")] string @class,
      [ExcelArgument("Does drug chemical structure consist of at least one oxygen atom? (TRUE|FALSE)")] bool containsOxygen,
      [ExcelArgument("pKa")] object pKa,
      [ExcelArgument(Name = "pKa.base", Description = "pKa.base (zwitterion only)")] object pKa_base
      )
    {
      var tissues = _table1.Keys;

      var Vss = 0.0;

      foreach (var tissue in tissues)
      {
        if (tissue == "plasma")
        {
          Vss += _table1[tissue]["V"];
        }
        else
        {
          var o = ExcelFunctions.Kp(tissue, logPow, @class, containsOxygen, pKa, pKa_base);
          if (o is ExcelError) return o;
          var Kp = (double)o;
          Vss += Kp * _table1[tissue]["V"];
        }
      }

      return Vss;
    }

    private static object Iw(string @class, object pKa, object pKa_base, double pH)
    {
      var isValidArg = Enum.TryParse<IonizationClass>(@class, out IonizationClass ionizationClass);
      if (!isValidArg) return ExcelError.ExcelErrorValue;

      switch (ionizationClass)
      {
        case IonizationClass.A:
          if (!(pKa is double)) return ExcelError.ExcelErrorValue;
          return Pow(10.0, pH - (double)pKa);

        case IonizationClass.N: return 0.0;

        case IonizationClass.SB:
        case IonizationClass.WB:
          if (!(pKa is double)) return ExcelError.ExcelErrorValue;
          return Pow(10.0, (double)pKa - pH);

        case IonizationClass.Z:
          if (!(pKa_base is double) || !(pKa is double)) return ExcelError.ExcelErrorValue;
          return Pow(10.0, (double)pKa_base - pH) + Pow(10.0, pH - (double)pKa);

        default: throw new ArgumentOutOfRangeException(nameof(ionizationClass));
      }
    }

    private static object LogPvow(double logPow, IonizationClass ionizationClass, bool containsOxygen)
    {
      if (!containsOxygen || ionizationClass == IonizationClass.SB)
      {
        return 1.0654 * logPow - 0.232;
      }

      return 1.099 * logPow - 1.31;
    }

    private enum IonizationClass { N, WB, SB, A, Z };

    private static IDictionary<string, IDictionary<string, double>> _table1 = new SortedDictionary<string, IDictionary<string, double>>
    {
      ["adipose"] = new SortedDictionary<string, double> { ["Fw"] = 0.15, ["Fnl"] = 0.800, ["V"] = 0.149 },
      ["bone"] = new SortedDictionary<string, double> { ["Fw"] = 0.45, ["Fnl"] = 0.074, ["V"] = 0.13 },
      ["brain"] = new SortedDictionary<string, double> { ["Fw"] = 0.78, ["Fnl"] = 0.068, ["V"] = 0.02 },
      ["gut"] = new SortedDictionary<string, double> { ["Fw"] = 0.76, ["Fnl"] = 0.054, ["V"] = 0.026 },
      ["heart"] = new SortedDictionary<string, double> { ["Fw"] = 0.78, ["Fnl"] = 0.017, ["V"] = 0.044 },
      ["kidneys"] = new SortedDictionary<string, double> { ["Fw"] = 0.76, ["Fnl"] = 0.026, ["V"] = 0.044 },
      ["liver"] = new SortedDictionary<string, double> { ["Fw"] = 0.73, ["Fnl"] = 0.043, ["V"] = 0.036 },
      ["lungs"] = new SortedDictionary<string, double> { ["Fw"] = 0.78, ["Fnl"] = 0.0062, ["V"] = 0.013 },
      ["muscle"] = new SortedDictionary<string, double> { ["Fw"] = 0.71, ["Fnl"] = 0.024, ["V"] = 0.484 },
      ["skin"] = new SortedDictionary<string, double> { ["Fw"] = 0.67, ["Fnl"] = 0.032, ["V"] = 0.08 },
      ["spleen"] = new SortedDictionary<string, double> { ["Fw"] = 0.79, ["Fnl"] = 0.027, ["V"] = 0.0029 },
      ["thymus"] = new SortedDictionary<string, double> { ["Fw"] = 0.78, ["Fnl"] = 0.020, ["V"] = 0.0001 },
      ["blood cells"] = new SortedDictionary<string, double> { ["Fw"] = 0.63, ["Fnl"] = 0.022, ["V"] = 0.0365 },
      ["plasma"] = new SortedDictionary<string, double> { ["Fw"] = 0.96, ["Fnl"] = 0.0073, ["V"] = 0.045 },
    };
  }
}

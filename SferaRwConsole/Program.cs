using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

var symbolTowaru = "ABC-123";
var serwer = "localhost";
var baza = "SUBIEKT_GT";
var uzytkownikSql = "sa";
var hasloSql = "haslo";
var operatorSymbol = "SZEF";
var operatorHaslo = "";
var produktSubiekt = 1;
var autentykacjaSql = 0;
var dokumentRwEnumWartosc = 13;

Console.WriteLine("Start");

object? gt = null;
object? subiekt = null;
object? dokument = null;

try
{
    Console.WriteLine("Łączenie z Sferą");
    var gtType = Type.GetTypeFromProgID("InsERT.GT");
    if (gtType == null)
    {
        throw new InvalidOperationException("Nie znaleziono COM InsERT.GT");
    }

    gt = Activator.CreateInstance(gtType);
    if (gt == null)
    {
        throw new InvalidOperationException("Nie można utworzyć obiektu InsERT.GT");
    }

    SetProperty(gt, "Produkt", produktSubiekt);
    SetProperty(gt, "Serwer", serwer);
    SetProperty(gt, "Baza", baza);
    SetProperty(gt, "Autentykacja", autentykacjaSql);
    SetProperty(gt, "Uzytkownik", uzytkownikSql);
    SetProperty(gt, "UzytkownikHaslo", hasloSql);
    SetProperty(gt, "Operator", operatorSymbol);
    SetProperty(gt, "OperatorHaslo", operatorHaslo);

    subiekt = InvokeMethod(gt, "Uruchom", 0, 0);
    if (subiekt == null)
    {
        throw new InvalidOperationException("Nie można uruchomić Subiekta GT");
    }

    Console.WriteLine($"Pobieranie towaru po symbolu: {symbolTowaru}");
    var towar =
        TryInvokeChain(subiekt, new[]
        {
            new MethodCall("Towary", "Wczytaj", symbolTowaru),
            new MethodCall("TowaryManager", "Wczytaj", symbolTowaru),
            new MethodCall("TowaryManager", "WczytajPoSymbolu", symbolTowaru)
        });

    if (towar == null)
    {
        throw new InvalidOperationException($"Nie znaleziono towaru o symbolu {symbolTowaru}");
    }

    Console.WriteLine("Utworzenie dokumentu RW");
    dokument = TryInvokeOnProperty(subiekt, "SuDokumentyManager", "DodajRW")
               ?? InvokeMethod(ResolveProperty(subiekt, "Dokumenty")!, "Dodaj", dokumentRwEnumWartosc);

    if (dokument == null)
    {
        throw new InvalidOperationException("Nie udało się utworzyć dokumentu RW");
    }

    Console.WriteLine("Dodawanie pozycji");
    var pozycje = ResolveProperty(dokument, "Pozycje");
    if (pozycje == null)
    {
        throw new InvalidOperationException("Nie udało się pobrać listy pozycji dokumentu");
    }

    var pozycja = InvokeMethod(pozycje, "Dodaj", towar);
    if (pozycja == null)
    {
        throw new InvalidOperationException("Nie udało się dodać pozycji do dokumentu");
    }

    SetProperty(pozycja, "IloscJm", 1m);

    Console.WriteLine("Zapis dokumentu");
    InvokeMethod(dokument, "Zapisz");
    TryInvoke(dokument, "Zamknij");

    Console.WriteLine("Sukces");
    return;
}
catch (Exception ex)
{
    Console.WriteLine("Błąd");
    Console.WriteLine(ex.ToString());
}
finally
{
    TryInvoke(subiekt, "Zakoncz");
    TryInvoke(gt, "Zakoncz");

    ReleaseComObject(dokument);
    ReleaseComObject(subiekt);
    ReleaseComObject(gt);
}

void SetProperty(object target, string propertyName, object value)
{
    var property = target.GetType().GetProperty(propertyName);
    if (property == null)
    {
        throw new InvalidOperationException($"Brak właściwości {propertyName}");
    }
    property.SetValue(target, value);
}

object? ResolveProperty(object target, string propertyName)
{
    var property = target.GetType().GetProperty(propertyName);
    return property?.GetValue(target);
}

object? InvokeMethod(object target, string methodName, params object[] args)
{
    var method = target.GetType().GetMethod(methodName);
    if (method == null)
    {
        throw new InvalidOperationException($"Brak metody {methodName}");
    }
    return method.Invoke(target, args);
}

object? TryInvoke(object? target, string methodName, params object[] args)
{
    if (target == null)
    {
        return null;
    }

    var method = target.GetType().GetMethod(methodName);
    return method?.Invoke(target, args);
}

object? TryInvokeOnProperty(object? root, string propertyName, string methodName, params object[] args)
{
    if (root == null)
    {
        return null;
    }

    var property = root.GetType().GetProperty(propertyName);
    if (property == null)
    {
        return null;
    }

    var value = property.GetValue(root);
    if (value == null)
    {
        return null;
    }

    var method = value.GetType().GetMethod(methodName);
    return method?.Invoke(value, args);
}

object? TryInvokeChain(object root, IEnumerable<MethodCall> calls)
{
    foreach (var call in calls)
    {
        var property = root.GetType().GetProperty(call.PropertyName);
        if (property == null)
        {
            continue;
        }

        var propertyValue = property.GetValue(root);
        if (propertyValue == null)
        {
            continue;
        }

        var method = propertyValue.GetType().GetMethod(call.MethodName);
        if (method == null)
        {
            continue;
        }

        var result = method.Invoke(propertyValue, call.Arguments);
        if (result != null)
        {
            return result;
        }
    }

    return null;
}

void ReleaseComObject(object? obj)
{
    if (obj != null && Marshal.IsComObject(obj))
    {
        Marshal.ReleaseComObject(obj);
    }
}

readonly record struct MethodCall(string PropertyName, string MethodName, params object[] Arguments);

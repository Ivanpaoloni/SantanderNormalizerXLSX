using OfficeOpenXml;
using System.Data;
using System.Globalization;

namespace SantanderNormalizer;

class Program
{
    static void Main(string[] args)
    {
        string ruta = "H:\\descargas\\movimientos.xlsx";

        string filePath = args.Length > 0 ? args[0] : ruta;

        if (string.IsNullOrWhiteSpace(filePath))
        {
            Console.WriteLine("Uso: SantanderNormalizer <archivo.xlsx>");
            return;
        }

        if (!File.Exists(filePath))
        {
            Console.WriteLine("Archivo no encontrado.");
            return;
        }

        ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");

        try
        {
            var normalizados = ParseSantanderExtract(filePath);
            Console.WriteLine($"Movimientos leídos: {normalizados.Rows.Count}");

            var output = Path.Combine(
                Path.GetDirectoryName(filePath)!,
                Path.GetFileNameWithoutExtension(filePath) + "_normalizado.xlsx");

            GuardarExcel(normalizados, output);

            Console.WriteLine($"Archivo generado: {output}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    static DataTable ParseSantanderExtract(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var ws = package.Workbook.Worksheets[0];

        int headerRow = -1;

        for (int row = ws.Dimension.Start.Row; row <= ws.Dimension.End.Row; row++)
        {
            bool tieneFecha = false;
            bool tieneDescripcion = false;
            bool tieneImporte = false;

            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var t = NormalizarHeader(ws.Cells[row, col].Text);

                if (t.Contains("fecha")) tieneFecha = true;
                else if (t.Contains("descrip") || t.Contains("concepto")) tieneDescripcion = true;
                else if (t.Contains("caja") || t.Contains("cuenta") || t.Contains("importe") || t.Contains("ahorro") || t.Contains("corriente"))
                    tieneImporte = true;
            }

            if (tieneFecha && tieneDescripcion && tieneImporte)
            {
                headerRow = row;
                break;
            }
        }


        if (headerRow == -1)
            throw new Exception("No se encontró encabezado 'Fecha'.");

        int colFecha = -1, colDescripcion = -1, colCaja = -1, colCuenta = -1;

        int colCount = ws.Dimension.End.Column;

        for (int col = 1; col <= ws.Dimension.End.Column; col++)
        {
            var name = NormalizarHeader(ws.Cells[headerRow, col].Text);

            if (name.Contains("fecha")) colFecha = col;
            else if (name.Contains("descrip") || name.Contains("concepto")) colDescripcion = col;
            else if (name.Contains("caja") || name.Contains("ahorro") || name.Contains("importe")) colCaja = col;
            else if (name.Contains("cuenta") || name.Contains("corriente")) colCuenta = col;
        }

        if (colFecha == -1 || colDescripcion == -1 || (colCaja == -1 && colCuenta == -1))
            throw new Exception("No se pudieron identificar las columnas necesarias.");

        var table = new DataTable();
        table.Columns.Add("Fecha");
        table.Columns.Add("Concepto");
        table.Columns.Add("Importe", typeof(decimal));

        for (int row = headerRow + 1; row <= ws.Dimension.End.Row; row++)
        {
            if (string.IsNullOrWhiteSpace(ws.Cells[row, colFecha].Text))
                break;

            var fecha = ws.Cells[row, colFecha].Text.Trim();
            var concepto = ws.Cells[row, colDescripcion].Text.Trim();

            var cajaRaw = colCaja > 0 ? ws.Cells[row, colCaja].Value : null;
            var cuentaRaw = colCuenta > 0 ? ws.Cells[row, colCuenta].Value : null;

            decimal caja = ParseDecimal(cajaRaw);
            decimal cuenta = ParseDecimal(cuentaRaw);

            decimal importe = caja != 0 ? caja : cuenta;

            table.Rows.Add(fecha, concepto, importe);
        }

        return table;
    }
   
    static decimal ParseDecimal(object value)
    {
        if (value == null)
            return 0;

        decimal result;

        if (value is double d)
            result = (decimal)d;
        else if (value is decimal m)
            result = m;
        else
        {
            var s = value.ToString()?.Trim();
            if (string.IsNullOrEmpty(s))
                return 0;

            // Quitar símbolos
            s = s.Replace("$", "").Replace(" ", "");

            if (!decimal.TryParse(s, NumberStyles.Any, new CultureInfo("es-AR"), out result) &&
                !decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                return 0;
        }

        // 🔧 Ajuste Santander: viene en centavos (ej 1433250 => 14332.50)
        if (Math.Abs(result) > 100000 && result % 1 == 0)
            result = result / 100;

        return result;
    }
    static string NormalizarHeader(string s)
    {
        return s.ToLower()
                .Replace("á", "a")
                .Replace("é", "e")
                .Replace("í", "i")
                .Replace("ó", "o")
                .Replace("ú", "u")
                .Replace("$", "")
                .Replace(".", "")
                .Replace("_", "")
                .Trim();
    }

    static void GuardarExcel(DataTable table, string path)
    {
        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Normalizado");

        ws.Cells["A1"].LoadFromDataTable(table, true);
        ws.Cells.AutoFitColumns();

        package.SaveAs(new FileInfo(path));
    }
}

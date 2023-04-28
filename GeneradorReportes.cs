using System.Collections.Generic;
using ClosedXML.Excel;

namespace ReportesExcel
{
    public class GeneradorReportes
    {
        public void ExportarReporteAExcel(List<Producto> productos, string rutaArchivo)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte de Productos");

               
                worksheet.Cell(1, 1).Value = "ID";
                worksheet.Cell(1, 2).Value = "Nombre";
                worksheet.Cell(1, 3).Value = "Precio";
                worksheet.Cell(1, 4).Value = "Cantidad";

                for (int i = 0; i < productos.Count; i++)
                {
                    var producto = productos[i];
                    worksheet.Cell(i + 2, 1).Value = producto.Id;
                    worksheet.Cell(i + 2, 2).Value = producto.Nombre;
                    worksheet.Cell(i + 2, 3).Value = producto.Precio;
                    worksheet.Cell(i + 2, 4).Value = producto.Cantidad;
                }

                workbook.SaveAs(rutaArchivo);
            }
        }
    }
}



using System;
using System.Collections.Generic;

namespace ReportesExcel
{
    class Program
    {
        static void Main(string[] args)
        {
           
            List<Producto> productos = new List<Producto>
            {
                new Producto { Id = 1, Nombre = "Producto 1", Precio = 10.5m, Cantidad = 5 },
                new Producto { Id = 2, Nombre = "Producto 2", Precio = 15.25m, Cantidad = 3 },
                new Producto { Id = 3, Nombre = "Producto 3", Precio = 8.0m, Cantidad = 10 },
                new Producto { Id = 4, Nombre = "Producto 4", Precio = 20.0m, Cantidad = 1 },
            };

          
            var generadorReportes = new GeneradorReportes();
            string rutaArchivo = "ReporteProductos.xlsx";
            generadorReportes.ExportarReporteAExcel(productos, rutaArchivo);

            Console.WriteLine($"El reporte ha sido generado y guardado en {rutaArchivo}");
            Console.ReadKey();
        }
    }
}


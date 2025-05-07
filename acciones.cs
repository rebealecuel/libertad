using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace libertad
{
    internal class acciones
    {
        private List<Alumno> alumnosList = new List<Alumno>
        {
            new Alumno ("Rebeca",20,"Ladd",112969,DateTime.Today),
            new Alumno ("Maya",19,"Ladd",112901, DateTime.Today),
            new Alumno ("Cindy",20,"Ladd",112816,DateTime.Today),
            new Alumno ("Angela",20,"Ladd",112318,DateTime.Today)
        };

        public List<Alumno> Mostrar()
        {
            return alumnosList;
        }

        public bool ExportarExcel()
        {
            try
            {
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var filePath = Path.Combine(desktopPath, "Alumnos.xlsx");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Lista de Alumnos");
                    worksheet.Cell(1, 1).Value = "Nombre";
                    worksheet.Cell(1, 2).Value = "Edad";
                    worksheet.Cell(1, 3).Value = "Carrera";
                    worksheet.Cell(1, 4).Value = "Matricula";
                    worksheet.Cell(1, 5).Value = "Fecha de nacimiento";

                    int row = 2;
                    foreach (var alumno in alumnosList)
                    {
                        worksheet.Cell(row, 1).Value = alumno.Nombre;
                        worksheet.Cell(row, 2).Value = alumno.Edad;
                        worksheet.Cell(row, 3).Value = alumno.Carrera;
                        worksheet.Cell(row, 4).Value = alumno.Matricula;
                        worksheet.Cell(row, 5).Value = alumno.Fechanacimiento.ToShortDateString();
                        row++;
                    }

                    workbook.SaveAs(filePath);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
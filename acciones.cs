using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace libertad
{
    internal class acciones
    {
        private List<Alumno> alumnosList = new List<Alumno>();
        Correo correo = new Correo();

        public List<Alumno> Mostrar()
        {
            try
            {
                return alumnosList;
            }
            catch (Exception ex) 
            {
                correo.EnviarCorreo(ex.ToString());
                throw;
            }
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
                    worksheet.Cell(1, 0).Value = "Nombre";
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
            catch (Exception ex)
            {
                correo.EnviarCorreo(ex.ToString());
                return false;
            }
        }

        public bool ImportarExcel()
        {
            try
            {
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var filePath = Path.Combine(desktopPath, "Alumnos.xlsx");

                if (!File.Exists(filePath))
                {
                    return false;
                }

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet("Alumno");

                    var rows = worksheet.RowsUsed().Skip(1); // Saltar los encabezados

                    alumnosList.Clear(); // Limpiar la lista antes de importar nuevos datos

                    foreach (var row in rows)
                    {
                        var nombre = row.Cell(1).Value.ToString();

                        // Usar TryParse para evitar excepciones en la conversión
                        int edad = 0;
                        int.TryParse(row.Cell(2).Value.ToString(), out edad); // Intentar convertir a entero

                        var carrera = row.Cell(3).Value.ToString();

                        int matricula = 0;
                        int.TryParse(row.Cell(4).Value.ToString(), out matricula); // Intentar convertir a entero

                        DateTime fechanacimiento;
                        DateTime.TryParse(row.Cell(5).Value.ToString(), out fechanacimiento); // Intentar convertir a DateTime

                        // Agregar el Alumno solo si la conversión fue exitosa
                        alumnosList.Add(new Alumno(nombre, edad, carrera, matricula, fechanacimiento));
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                correo.EnviarCorreo(ex.ToString());
                return false;
            }
        }
    }
}
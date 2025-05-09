using System;
using System.Net;
using System.Net.Mail;
using System.IO;
using ClosedXML.Excel;

namespace libertad
{
    internal class Correo
    {
        public void EnviarCorreo(string error)
        {
            try
            {
                // 1. Crear un archivo Excel con ClosedXML
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet("Errores");
                ws.Cell("A1").Value = "Descripción del Error";
                ws.Cell("A2").Value = error;

                // 2. Guardar el archivo Excel en un MemoryStream (sin crear un archivo físico)
                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    stream.Position = 0; // Asegurarse de que el stream está en la posición inicial

                    // 3. Configuración del correo
                    var mensaje = new MailMessage();
                    mensaje.From = new MailAddress("112969@alumnouninter.mx"); // Cambiar por tu correo de Office 365
                    mensaje.To.Add("ecorrales@uninter.edu.mx");
                    mensaje.Subject = "Informe de Error";
                    mensaje.Body = "Se adjunta el informe de error generado.";

                    // 4. Crear el adjunto del archivo Excel
                    var adjunto = new Attachment(stream, "informe_error.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    mensaje.Attachments.Add(adjunto);

                    // 5. Configuración del servidor SMTP de Office 365
                    var smtpCliente = new SmtpClient("smtp.office365.com")
                    {
                        Port = 587, // El puerto que se usa para SMTP seguro (TLS)
                        Credentials = new NetworkCredential("112969@alumnouninter.mx", "Rebelalobafeliz654321"), // Cambiar por tu correo y contraseña
                        EnableSsl = true
                    };

                    // 6. Enviar el correo
                    smtpCliente.Send(mensaje);
                    Console.WriteLine("Correo enviado con éxito.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al enviar el correo: {ex.Message}");
                throw;
            }
        }
    }
}
using OfficeOpenXml;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Timers;

namespace WebSocketServer
{
    class Program
    {
        static Dictionary<int, VisitaPiso> visitasPorPiso = new Dictionary<int, VisitaPiso>();
        static System.Timers.Timer timer;

        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var httpListener = new HttpListener();
            httpListener.Prefixes.Add("http://localhost:9090/");
            httpListener.Start();

            Console.WriteLine("Servidor WebSocket iniciado. Esperando conexiones...");

            timer = new System.Timers.Timer(120000); // 120000 milisegundos = 2 minutos
            timer.Elapsed += OnTimedEvent;
            timer.AutoReset = true;
            timer.Enabled = true;

            while (true)
            {
                var context = await httpListener.GetContextAsync();
                if (context.Request.IsWebSocketRequest)
                {
                    var webSocketContext = await context.AcceptWebSocketAsync(null);
                    await HandleWebSocketAsync(webSocketContext.WebSocket);
                }
                else
                {
                    context.Response.StatusCode = 400;
                    context.Response.Close();
                }
            }
        }

        static async Task HandleWebSocketAsync(WebSocket webSocket)
        {
            int pisoActual = 1;
            int personasEnAscensor = 0;
            int? pisoSolicitado = null;

            byte[] buffer = new byte[1024];
            WebSocketReceiveResult result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);

            while (!result.CloseStatus.HasValue)
            {
                string message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                Console.WriteLine($"Mensaje recibido: {message}");

                if (int.TryParse(message, out int floor))
                {
                    if (floor >= 1 && floor <= 10)
                    {
                        pisoSolicitado = floor;

                        int personasEntrando = Math.Max(0, Math.Min(10 - personasEnAscensor, 10 - pisoActual));
                        personasEnAscensor += personasEntrando;

                        int personasSaliendo = 0;
                        if (floor != pisoActual)
                        {
                            personasSaliendo = Math.Min(personasEnAscensor, Math.Abs(floor - pisoActual));
                            personasEnAscensor -= personasSaliendo;
                        }

                        while (pisoActual != pisoSolicitado)
                        {
                            if (pisoActual < pisoSolicitado)
                            {
                                pisoActual++;
                            }
                            else
                            {
                                pisoActual--;
                            }

                            if (!visitasPorPiso.ContainsKey(pisoActual))
                            {
                                visitasPorPiso[pisoActual] = new VisitaPiso { Piso = pisoActual, CantidadDeVisitas = 1 };
                            }
                            else
                            {
                                visitasPorPiso[pisoActual].CantidadDeVisitas++;
                            }

                            // Actualizar la cantidad de personas que han visitado el piso
                            visitasPorPiso[pisoActual].PersonasQueHanVisitado += personasEnAscensor;

                            string response = $"El ascensor está en el piso {pisoActual} con {personasEnAscensor} personas.";
                            byte[] responseBuffer = Encoding.UTF8.GetBytes(response);
                            await webSocket.SendAsync(new ArraySegment<byte>(responseBuffer), WebSocketMessageType.Text, true, CancellationToken.None);

                            await Task.Delay(2000);
                        }
                    }
                    else
                    {
                        string errorMessage = "El piso especificado no es válido.";
                        byte[] errorBuffer = Encoding.UTF8.GetBytes(errorMessage);
                        await webSocket.SendAsync(new ArraySegment<byte>(errorBuffer), WebSocketMessageType.Text, true, CancellationToken.None);
                    }
                }

                buffer = new byte[1024];
                result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
            }

            await webSocket.CloseAsync(result.CloseStatus.Value, result.CloseStatusDescription, CancellationToken.None);
        }

        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            GuardarVisitasEnExcel();
        }

        private static void GuardarVisitasEnExcel()
        {
            string fileName = "visitas_pisos.xlsx"; // Nombre del archivo
            string directory = AppDomain.CurrentDomain.BaseDirectory; // Obtiene la ruta del directorio actual

            string filePath = Path.Combine(directory, fileName); // Combina la ruta del directorio con el nombre del archivo

            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                file.Delete(); // Elimina el archivo existente para escribir uno nuevo
            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Visitas");

                worksheet.Cells[1, 1].Value = "Piso";
                worksheet.Cells[1, 2].Value = "Personas que han visitado";
                worksheet.Cells[1, 3].Value = "Cantidad de visitas";

                int row = 2;
                foreach (var visitaPiso in visitasPorPiso.Values)
                {
                    worksheet.Cells[row, 1].Value = visitaPiso.Piso;
                    worksheet.Cells[row, 2].Value = visitaPiso.PersonasQueHanVisitado;
                    worksheet.Cells[row, 3].Value = visitaPiso.CantidadDeVisitas;
                    row++;
                }

                package.Save();
            }

            Console.WriteLine($"Datos guardados en {filePath} - {DateTime.Now}");
        }
    }

    class VisitaPiso
    {
        public int Piso { get; set; }
        public int PersonasQueHanVisitado { get; set; }
        public int CantidadDeVisitas { get; set; }
    }
}

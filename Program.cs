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
        static SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);

        static async Task Main(string[] args)
        {
            // Establece el contexto de licencia de ExcelPackage a no comercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Crea un HttpListener para escuchar las solicitudes entrantes
            var httpListener = new HttpListener();

            // Agrega la URL de prefijo para la escucha
            httpListener.Prefixes.Add("http://localhost:9090/");

            // Inicia el listener
            httpListener.Start();

            // Imprime un mensaje en la consola indicando que el servidor WebSocket está iniciado
            Console.WriteLine("Servidor WebSocket iniciado. Esperando conexiones...");

            // Crea un temporizador que se activará cada 2 minutos
            var timer = new System.Timers.Timer(120000); // 120000 milisegundos = 2 minutos

            // Asigna el método OnTimedEvent para manejar los eventos del temporizador
            timer.Elapsed += OnTimedEvent;

            // Configura el temporizador para que se reinicie automáticamente y lo habilita
            timer.AutoReset = true;
            timer.Enabled = true;

            // Entra en un bucle infinito para manejar las solicitudes entrantes
            while (true)
            {
                // Espera una solicitud HTTP entrante
                var context = await httpListener.GetContextAsync();

                // Verifica si la solicitud es una solicitud WebSocket
                if (context.Request.IsWebSocketRequest)
                {
                    // Acepta la conexión WebSocket
                    var webSocketContext = await context.AcceptWebSocketAsync(null);

                    // Maneja la conexión WebSocket de forma asincrónica
                    await HandleWebSocketAsync(webSocketContext.WebSocket);
                }
                else
                {
                    // Rechaza las solicitudes que no sean WebSocket
                    context.Response.StatusCode = 400;
                    context.Response.Close();
                }
            }
        }

        // Maneja una conexión WebSocket para un ascensor individual.
        // Recibe solicitudes de piso, simula el movimiento del ascensor y envía actualizaciones de estado.
        static async Task HandleWebSocketAsync(WebSocket webSocket)
        {
            // Piso actual del ascensor
            int pisoActual = 1;

            // Número de personas en el ascensor
            int personasEnAscensor = 0;

            // Piso solicitado por un cliente (si lo hay)
            int? pisoSolicitado = null;

            // Buffer para recibir mensajes WebSocket
            byte[] buffer = new byte[1024];

            // Recibe el primer mensaje del cliente
            WebSocketReceiveResult result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);

            // Bucle principal que maneja la comunicación WebSocket
            while (!result.CloseStatus.HasValue)
            {
                // Decodifica el mensaje recibido como una cadena UTF-8
                string message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                Console.WriteLine($"Mensaje recibido: {message}");

                // Intenta convertir el mensaje en un número entero (que representa el piso solicitado)
                if (int.TryParse(message, out int floor))
                {
                    // Valida que el piso solicitado esté entre 1 y 10
                    if (floor >= 1 && floor <= 10)
                    {
                        pisoSolicitado = floor;

                        // Calcula la cantidad de personas que entran en el ascensor (considerando capacidad máxima)
                        int personasEntrando = Math.Max(0, Math.Min(10 - personasEnAscensor, 10 - pisoActual));
                        personasEnAscensor += personasEntrando;

                        // Calcula la cantidad de personas que salen del ascensor (si el piso solicitado es diferente)
                        int personasSaliendo = 0;
                        if (floor != pisoActual)
                        {
                            personasSaliendo = Math.Min(personasEnAscensor, Math.Abs(floor - pisoActual));
                            personasEnAscensor -= personasSaliendo;
                        }

                        // Mueve el ascensor hasta el piso solicitado
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

                            // Registra la visita al piso actual
                            await semaphore.WaitAsync();
                            try
                            {
                                if (!visitasPorPiso.ContainsKey(pisoActual))
                                {
                                    visitasPorPiso[pisoActual] = new VisitaPiso { Piso = pisoActual, CantidadDeVisitas = 1 };
                                }
                                else
                                {
                                    visitasPorPiso[pisoActual].CantidadDeVisitas++;
                                }

                                // Actualiza la cantidad de personas que han visitado el piso
                                visitasPorPiso[pisoActual].PersonasQueHanVisitado += personasEnAscensor;
                            }
                            finally
                            {
                                semaphore.Release();
                            }

                            // Envía una actualización de estado al cliente
                            string response = $"\u001b[36mEl ascensor está en el piso {pisoActual} con {personasEnAscensor} personas.\u001b[0m";
                            byte[] responseBuffer = Encoding.UTF8.GetBytes(response);
                            await webSocket.SendAsync(new ArraySegment<byte>(responseBuffer), WebSocketMessageType.Text, true, CancellationToken.None);

                            // Simula el tiempo de movimiento del ascensor (2 segundos)
                            await Task.Delay(2000);
                        }
                    }
                    else
                    {
                        // Si el piso solicitado no es válido, envía un mensaje de error al cliente
                        string errorMessage = "\u001b[31mEl piso especificado no es válido.\u001b[0m";
                        byte[] errorBuffer = Encoding.UTF8.GetBytes(errorMessage);
                        await webSocket.SendAsync(new ArraySegment<byte>(errorBuffer), WebSocketMessageType.Text, true, CancellationToken.None);
                    }
                }

                // Recibe el siguiente mensaje del cliente
                buffer = new byte[1024];
                result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
            }

            // Cierra la conexión WebSocket
            await webSocket.CloseAsync(result.CloseStatus.Value, result.CloseStatusDescription, CancellationToken.None);
        }

        // Manejador de eventos que se ejecuta periódicamente para guardar las visitas a pisos en un archivo de Excel.
        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            // Llama a la función para guardar los datos en Excel.
            GuardarVisitasEnExcel();
        }

        private static void GuardarVisitasEnExcel()
        {
            // Nombre del archivo de Excel
            string fileName = "visitas_pisos.xlsx";

            // Directorio actual del archivo
            string directory = AppDomain.CurrentDomain.BaseDirectory;

            // Ruta completa del archivo
            string filePath = Path.Combine(directory, fileName);

            // Crea un objeto FileInfo para el archivo
            FileInfo file = new FileInfo(filePath);

            // Verifica si el archivo está abierto por otra aplicación
            if (file.Exists && !IsFileLocked(file))
            {
                file.Delete();
            }
            // Si el archivo no está abierto, intenta eliminarlo para escribir uno nuevo
            else if (!file.Exists)
            {
                // Crea el archivo si no existe
                using (var stream = file.Create()) { }
            }
            else
            {
                Console.WriteLine($"El archivo {fileName} está abierto por otra aplicación o no se puede acceder.");
                return;
            }


            // Crea un nuevo archivo de Excel y guarda la información de visitas a pisos en él.
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // Crea una nueva hoja de trabajo llamada "Visitas".
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Visitas");

                // Agrega encabezados a la hoja de trabajo.
                worksheet.Cells[1, 1].Value = "Piso";
                worksheet.Cells[1, 2].Value = "Personas que han visitado";
                worksheet.Cells[1, 3].Value = "Cantidad de visitas";

                // Recorre los datos de visitas a pisos y los agrega a la hoja de trabajo.
                int row = 2;  // Comienza a escribir datos en la fila 2 (después de los encabezados)
                foreach (var visitaPiso in visitasPorPiso.Values)
                {
                    worksheet.Cells[row, 1].Value = visitaPiso.Piso;
                    worksheet.Cells[row, 2].Value = visitaPiso.PersonasQueHanVisitado;
                    worksheet.Cells[row, 3].Value = visitaPiso.CantidadDeVisitas;
                    row++;  // Avanza a la siguiente fila
                }

                // Guarda los cambios en el archivo de Excel.
                package.Save();
            }

            // Imprime un mensaje en la consola indicando la ubicación y la fecha/hora de guardado del archivo.
            Console.WriteLine($"Datos guardados en {filePath} - {DateTime.Now}");


        }

        // Método para verificar si un archivo está bloqueado por otra aplicación
        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                // Intenta abrir el archivo en modo de lectura y escritura
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                // Si se produce una excepción de E/S, significa que el archivo está bloqueado
                return true;
            }
            finally
            {
                // Cierra el stream si se ha abierto correctamente
                if (stream != null)
                {
                    stream.Close();
                }
            }

            // Si el stream se abrió correctamente, el archivo no está bloqueado
            return false;
        }

    }

    // Clase que representa información sobre las visitas a un piso específico en el sistema de ascensor.
    class VisitaPiso
    {
        // Número del piso al que se refiere la información.
        public int Piso { get; set; }

        // Número total de personas que han visitado el piso.
        public int PersonasQueHanVisitado { get; set; }

        // Número de veces que el piso ha sido visitado.
        public int CantidadDeVisitas { get; set; }
    }
}

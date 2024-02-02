# Servidor WebSocket con la lógica del ascensor

Este proyecto contiene un servidor WebSocket en C# que simula un ascensor y registra las visitas a cada piso.

## Requisitos

- [.NET Core SDK](https://dotnet.microsoft.com/download) instalado en tu sistema.
- Paquete NuGet `EPPlus` para trabajar con archivos de Excel.

## Instalación de paquetes NuGet

Para instalar el paquete `EPPlus`, puedes utilizar el siguiente comando en la terminal:

```bash
dotnet add package EPPlus
```

## Uso

1. Clona este repositorio en tu máquina local:

   ```bash
   git clone <url-del-repositorio>
   ```

2. Abre una terminal y navega hasta el directorio del proyecto.

3. Compila el proyecto utilizando el siguiente comando:

   ```bash
   dotnet build
   ```

4. Ejecuta la aplicación utilizando el siguiente comando:

   ```bash
   dotnet run
   ```

5. El servidor WebSocket estará escuchando en el puerto `9090` del localhost. Asegúrate de tener un cliente WebSocket que pueda conectarse al servidor y enviar mensajes de solicitud de piso.

6. El servidor registrará las visitas a cada piso y guardará los datos en un archivo de Excel llamado `visitas_pisos.xlsx` en el directorio actual. Los datos se guardarán cada 2 minutos.

## Repositorio del cliente

[Cliente-WebSocket](https://github.com/Naeliza/Cliente-WebSocket) 

## Contribución

Si encuentras algún problema o tienes alguna sugerencia para mejorar este servidor WebSocket, ¡no dudes en abrir un issue o enviar un pull request!

## Licencia

Este proyecto está bajo la [Licencia MIT](LICENSE).

# ¿Interesado/da en aprender más?

Sitio web: [Naeliza.com](https://naeliza.com/)

Portafolio: [Portafolio Naomi Céspedes](https://naeliza.netlify.app/#home)

Canal de Youtube: [AE Coding](https://www.youtube.com/@AECoding)

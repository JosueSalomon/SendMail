import nodemailer from 'nodemailer';
import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

// Obtener la ruta del directorio actual
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Configuración del transporter
export const transporter = nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 465,
    secure: true, // true para el puerto 465, false para otros puertos
    auth: {
        user: "josueisacsalomonlanda@gmail.com",
        pass: "xvqj mjdj aoxh zwzg", // Usa una contraseña de aplicación
    },
});

// Función para leer correos desde el archivo Excel
function leerCorreosDesdeExcel(archivo) {
    const workbook = XLSX.readFile(archivo); // Leer el archivo Excel
    const sheetName = workbook.SheetNames[0]; // Obtener la primera hoja
    const sheet = workbook.Sheets[sheetName]; // Obtener los datos de la hoja
    const datos = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convertir los datos en formato JSON
    return datos.map(row => row[0]); // Suponiendo que los correos están en la primera columna
}

// Ruta del archivo Excel
const archivoExcel = path.join(__dirname, 'correos.xlsx'); // Usamos __dirname obtenido con fileURLToPath

// Información del correo
const asunto = "ACTUALIZACIÓN DE CONTRASEÑA DE CAMPUS VIRTUAL";
const mensajePrincipal = `
    Debido a la filtración de contraseñas que han surgido en las últimas semanas en la universidad, 
    le recomendamos actualizar su contraseña institucional lo antes posible. 
    Puede hacerlo ingresando a su portal virtual mediante el siguiente enlace seguro:
`;
const urlVisible = "https://campusvirtual.unah.edu.hn/";
const urlReal = "http://192.168.0.25"; // URL real que redirige a la IP interna

// Verificar si el transporte está listo
transporter.verify().then(() => {
    console.log("Ready for send emails");

    // Leer los correos desde el archivo Excel
    const emails = leerCorreosDesdeExcel(archivoExcel);
    console.log("Correos leídos desde el archivo:", emails);

    // Recorrer la lista de correos y enviar a cada uno
    emails.forEach(email => {
        const DatosCorreo = {
            from: '"UNAH - Seguridad Informática" <josueisacsalomonlanda@gmail.com>',
            to: email,
            subject: asunto,
            html: `
                <!DOCTYPE html>
                <html lang="es">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Actualización de Contraseña</title>
                    <style>
                        body {
                            font-family: 'Arial', sans-serif;
                            background-color: #f4f4f4;
                            margin: 0;
                            padding: 0;
                        }
                        .container {
                            max-width: 600px;
                            margin: 40px auto;
                            background-color: #ffffff;
                            border-radius: 8px;
                            overflow: hidden;
                            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                            border: 1px solid #e6e6e6;
                        }
                        .header {
                            background-color: #0a47a1;
                            color: #ffffff;
                            text-align: center;
                            padding: 20px;
                        }
                        .header img {
                            max-width: 150px;
                            margin-bottom: 10px;
                        }
                        .header h1 {
                            margin: 0;
                            font-size: 20px;
                            font-weight: normal;
                        }
                        .body {
                            padding: 20px;
                            color: #333333;
                        }
                        .body p {
                            line-height: 1.6;
                            font-size: 16px;
                            margin-bottom: 20px;
                        }
                        .body a {
                            color: #0a47a1;
                            text-decoration: none;
                            font-weight: bold;
                        }
                        .cta {
                            text-align: center;
                            margin: 30px 0;
                        }
                        .cta a {
                            background-color: #0a47a1;
                            color: #ffffff;
                            text-decoration: none;
                            padding: 10px 20px;
                            border-radius: 4px;
                            font-size: 16px;
                            display: inline-block;
                        }
                        .cta a:hover {
                            background-color: #083575;
                        }
                        .footer {
                            background-color: #f4f4f4;
                            color: #666666;
                            text-align: center;
                            font-size: 12px;
                            padding: 10px;
                            border-top: 1px solid #e6e6e6;
                        }
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="header">
                            <img src="https://ik.imagekit.io/diancrochet/UNAH-version-horizontal.png" alt="Logo UNAH">
                            <h1>Universidad Nacional Autónoma de Honduras</h1>
                        </div>
                        <div class="body">
                            <p>${mensajePrincipal}</p>
                            <p>Recuerde que mantener la seguridad de sus datos personales es una prioridad.</p>
                            <div class="cta">
                                <a href="${urlReal}">
                                    Ingresar a mi portal virtual
                                </a>
                            </div>
                            <p>
                                Si no solicitó esta acción, puede ignorar este mensaje. 
                                Para más información, visite nuestro sitio web en 
                                <a href="${urlReal}">${urlVisible}</a>.
                            </p>
                        </div>
                        <div class="footer">
                            &copy; 2024 UNAH. Todos los derechos reservados.
                        </div>
                    </div>
                </body>
                </html>
            `,
        };

        // Enviar el correo
        transporter.sendMail(DatosCorreo, (error, info) => {
            if (error) {
                console.error(`Error al enviar el correo a ${email}:`, error);
            } else {
                console.log(`Correo enviado a ${email}:`, info.response);
            }
        });
    });
}).catch(error => {
    console.error("Error al verificar el transporter:", error);
});

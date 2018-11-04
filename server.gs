/*  
	Coded by Jairo Ruiz Saenz - October 2018

	Based on:

	https://www.labnol.org/internet/receive-files-in-google-drive/19697/	
    https://medium.com/@dmccoy/how-to-submit-an-html-form-to-google-sheets-without-google-forms-b833952cc175
    https://ctrlq.org/code/19117-save-gmail-as-pdf?_ga=2.160396157.1718000879.1540091702-379554840.1539588470
*/

function doGet(e) {
    return HtmlService.createHtmlOutputFromFile('forms.html').setTitle("Convocatoria FPS");
}

function createFolder_01(folderName, name, id_number) {
    
    try {
        
        var dropbox3 = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
        if (dropbox3 == null){
            var dropbox3 = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
        }
        
        var dropbox = folderName;
        var folder;
        var folders = DriveApp.getFoldersByName(dropbox);

        if (folders.hasNext()) {
            folder = folders.next();
        } else {
            folder = DriveApp.createFolder(dropbox);
        }
        //-------------------------------------------------------

        var records_file_name = "Registros_convocatoria";
        var record_file;
        var files = DriveApp.getFilesByName(records_file_name);

        if (files.hasNext()) {
            record_file = files.next();
        } else {            
            record_file = SpreadsheetApp.create("Registros_convocatoria");
        }

        //-------------------------------------------------------
        var dropbox_ = id_number;
        var dropbox2 = dropbox_.toString().toUpperCase().replace(".", "").replace(",", "").replace("E", "").replace("-", "").replace("+", "");
        
        var dropbox_2 = name;
        var dropbox2_2 = dropbox_2.toString().toUpperCase().replace(" ", "_").replace("Á,", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U");

        var dropbox2 = ([dropbox2, dropbox2_2].join("_"));

        var folder2;
        var folders2 = folder.getFoldersByName(dropbox2);

        if (folders2.hasNext()) {
            folder2 = folders2.next();
        } else {
            folder2 = folder.createFolder(dropbox2);
        }

        //-------------------------------------------------------
        //var dropbox3 = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
        
        var folder3;
        var folders3 = folder2.getFoldersByName(dropbox3);

        if (folders3.hasNext()) {
            folder3 = folders3.next();
        } else {
            folder3 = folder2.createFolder(dropbox3);
        }
        
    } catch (f) {
        return "- createFolder_01:: " + f.toString();
    }
}


function uploadFileToGoogleDrive(data, file, name, id_number, universidad, email, tipo, file_id_2) {

    try {
        
        var dropbox3 = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
        if (dropbox3 == null){
            var dropbox3 = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
        }
        
        var dropbox = "Entregas Convocatoria";
        var folder;
        var folders = DriveApp.getFoldersByName(dropbox);

        if (folders.hasNext()) {
            folder = folders.next();
        } else {
            //folder = DriveApp.createFolder(dropbox);
        }

        //-------------------------------------------------------        
        
        if (tipo === "2") {

            var records_file_name = "Registros_convocatoria";
            var record_file;
            var files = DriveApp.getFilesByName(records_file_name);
            record_file = files.next();
                        
            var record_date = Utilities.formatDate(new Date(), "GMT-5", "dd-MM-yyyy' 'hh:mm a");
            var spreadsheet = SpreadsheetApp.open(record_file);
            var sheet = spreadsheet.getSheets()[0];
            //sheet.appendRow(['Cotton Sweatshirt XL', 'css004']);
            sheet.appendRow([record_date, id_number, name, universidad, email]);
            
        }

        //-------------------------------------------------------
        var dropbox_ = id_number;
        var dropbox2 = dropbox_.toString().toUpperCase().replace(".", "").replace(",", "").replace("E", "").replace("-", "").replace("+", "");
        
        var dropbox_2 = name;
        var dropbox2_2 = dropbox_2.toString().toUpperCase().replace(" ", "_").replace("Á,", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U");

        var dropbox2 = ([dropbox2, dropbox2_2].join("_"));

        var folder2;
        var folders2 = folder.getFoldersByName(dropbox2);

        if (folders2.hasNext()) {
            folder2 = folders2.next();
        } else {
            //folder2 = folder.createFolder(dropbox2);
        }

        //-------------------------------------------------------
        //var dropbox3 = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
        
        var folder3;
        var folders3 = folder2.getFoldersByName(dropbox3);

        if (folders3.hasNext()) {
            folder3 = folders3.next();
        } else {
            //folder3 = folder2.createFolder(dropbox3);
        }
        
        //-------------------------------------------------------
        var contentType = data.substring(5,data.indexOf(';'));
        var bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,')+7));
        
        var comp_name;
        if (tipo == "1") {
            comp_name = 'CV_';
        } else {
            comp_name = 'Cert_';
        }

        var blob = Utilities.newBlob(bytes, contentType, comp_name + file);    
        var file = folder3.createFile(blob);

        /*----------------------------------------------------*/

        // Get the URL of the document.
        // var url = file.getUrl();
        var file_id = file.getId();

        if (tipo === "2") {
            
            var my_date = Utilities.formatDate(new Date(), "GMT-5", "dd-MM-yyyy' a las 'hh:mm a");

            // Get the email address of the active user - that's you.
            var my_email = Session.getActiveUser().getEmail();

            // Get the name of the document to use as an email subject line.            
            var subject_ = 'FPS - Confirmación aplicación - Mensaje Automático';
            
            var body_ = 'Hola, este es el cuerpo del email';

            // Send yourself an email with a link to the document.        
            // MailApp.sendEmail(recipient = email, subject = subject_, body = body_);
            
            var attachment_01 = DriveApp.getFileById(file_id);
            var attachment_02 = DriveApp.getFileById(file_id_2);

            //var blob = Utilities.newBlob('Insert any HTML content here', 'text/html', 'my_document.html');
        
            var FPS_logo_url = "https://raw.githubusercontent.com/jairoruizsaenz/FPS_ResultadosEncuestas/master/resources/img/top_portada1%20-%20250.jpg";
            var FPS_logo_blob = UrlFetchApp
                         .fetch(FPS_logo_url)
                         .getBlob()
                         .setName("FPS_logo_blob");

            MailApp.sendEmail(recipient = email, subject = subject_, body = body_, {
                name: 'FPS - Confirmación Aplicación',
                attachments: [attachment_01.getAs(MimeType.PDF), attachment_02.getAs(MimeType.PDF)],
                bcc: my_email,
                htmlBody: "<img src='cid:FPSLogo' width='250px'>" + 
                    "<p style='font-size: 16px; font-weight: bold;'> Hola " + name + ",</p>" + 
                    'Gracias por aplicar al programa "Los Mejores por Colombia" del FPS-FNC, hemos recibido tu solicitud el dia de hoy ' + my_date +
                    ", la misma será evaluada y de resultar seleccionado(a) nos comunicaremos." +
                    "<br><p style='font-weight: bold; font-style: italic; color: red;'> --- Este es un mensaje automático, no responder a este correo --- </p>" +
                    "<br><div><strong>FONDO DE PASIVO SOCIAL FERROCARILES NACIONALES DE COLOMBIA</strong></div>" +
                    "<div>Bogota D.C., Colombia</div>" +
                    "<div>Dirección: Calle 13 No. 18-24 Estación de la Sabana (Bogota-Colombia)</div>" +
                    "<div>E-mail: quejasyreclamos@fps.gov.co - comunicaciones@fps.gov.co</div>" +
                    "<div>PBX: (57) (1)3817171 - Ext. 100, 173, 181 y 180 - 2476775</div>" +
                    "<div>Linea Gratuita: 01-8000-09-12206</div>" +
                    "<div>Horarios de Atención: Lunes a Viernes 7:30 a.m - 4:00 p.m Jornada Continua</div>",
                inlineImages:
                {
                    FPSLogo: FPS_logo_blob,
                }
            });
        }

        /*----------------------------------------------------*/

        if (tipo === "2") {
            return "OK";
        } else{            
            return file_id;
        }

    } catch (f) {
        return "- uploadFileToGoogleDrive:: " + f.toString();
    }
}
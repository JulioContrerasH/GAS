function crearPDF() {
    // Google Docs template ID
    var plantillaDocId = '1GBXZwWJLEBJCAh7fmSyuCDUaH7v66_A3HgsjIiQi9OI';
    var hojaId = '1OOcfF7AJgl3B-8yx_ICVmVO87CdQbMTMOuwZ5WgjHw8';
    var hoja = SpreadsheetApp.openById(hojaId).getSheets()[0];
  
    var fila = hoja.getLastRow();
    var nombre = hoja.getRange(fila, 2).getValue(); // Assumes name is in column B
  
    // Open template document and make a copy to preserve the original
    var documentoPlantilla = DriveApp.getFileById(plantillaDocId).makeCopy();
    var copiaId = documentoPlantilla.getId();
    var doc = DocumentApp.openById(copiaId);
    var cuerpo = doc.getBody();
  
    // Replace placeholder with the actual name
    cuerpo.replaceText('{{nombre}}', nombre);
  
    // Save and close temporary document
    doc.saveAndClose();
  
    // Convert document to PDF
    var blob = doc.getAs('application/pdf');
  
    var nombreDelArchivo = 'PDF de ' + nombre + '.pdf';
    blob.setName(nombreDelArchivo);
  
    // Save PDF to Google Drive
    var archivo = DriveApp.createFile(blob);
    archivo.setName('PDF de ' + nombre);
  
    // Option 2: Email the PDF (uncomment next line and provide an email)
    // Assuming email address is in column C
    var email = hoja.getRange(fila, 3).getValue();
    Logger.log(email);
    // Check if email address seems valid
    if (email && email.includes('@')) {
      // Option: Email the PDF
      var asunto = 'Tu PDF Personalizado';
      var cuerpoCorreo = 'Hola,\n\nAdjunto encontrarás tu certificado personalizado.';
      MailApp.sendEmail({
        to: email,
        subject: asunto,
        body: cuerpoCorreo,
        attachments: [blob]
      });
    } else {
      Logger.log('No se encontró una dirección de correo electrónico válida.');
    }
  
    // Delete temporary copy of the template
    DriveApp.getFileById(copiaId).setTrashed(true);
  }  
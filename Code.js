const EMAIL_SENT_COL = "Transition Notice email date";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Notify Campuses')
    .addItem('Send emails to students with blanks in column G', 'sendEmails')
    .addToUi();
}

function sendEmails(sheet=SpreadsheetApp.openById('1qTdUYWNZ5plMB6FlVYH-QTZkJajXaJBnovzrWxmCGac').getSheetByName('Return to HC 24-25')) { //{ .getActiveSheet()) {

  const dataRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Return to HC 24-25').getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();

  // Gets the index of the column named 'Transition Notice email date' (Assumes header names are unique)
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // Converts 2d array into an object array
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[EMAIL_SENT_COL] === ''){
      try {
        const campusInfo = getInfoByCampus(row['Campus']);
        const recipients = campusInfo.recipients;
        const driveLink = campusInfo.driveLink;
        const emailTemplate = getGmailTemplateFromDrafts_(row, driveLink);

        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row, driveLink);

        // Uncomment advanced parameters as needed
        GmailApp.sendEmail(recipients, msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          replyTo: 'john.decker@nisd.net',
        });
        // Edits cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

    function getInfoByCampus(campusValue) {
      switch (campusValue.toLowerCase()) {
        // case 'bernal':
        //   return {
        //    recipients: 'david.laboy@nisd.net, susan.schottler@nisd.net, monica.flores@nisd.net',
        //    driveLink: '1jyY1gPUgj2xLd7K5AWto6l3ictRP2xHQ'
        //   };
        // case 'briscoe':
        //   return {
        //     recipients: 'Joe.bishop@nisd.net, nereida.ollendieck@nisd.net, brigitte.rauschuber@nisd.net, jolanda.bowie@nisd.net',
        //     driveLink: '1gBcGRl700LGfhMPrHx32dZ2j3MwPlqYS'
        //   };
        // case 'connally':
        //   return {
        //     recipients: 'erica.robles@nisd.net, monica.ramirez@nisd.net',
        //     driveLink: '1dFXFBakVlHRaE2M4SQR9u-hVoFRs73-6'
        //   };
        // case 'folks':
        //   return {
        //     recipients: 'yvette.lopez@nisd.net, Miguel.trevino@nisd.net, ann.devlin@nisd.net, angelica.perez@nisd.net, terry.precie@nisd.net',
        //     driveLink: '1iA4d7ju4dU7aOcqr4R6PezzHUG5x1LK9'
        //   };
        // case 'garcia':
        //   return {
        //     recipients: 'Mark.Lopez@nisd.net, Lori.Persyn@nisd.net, Julie.MInnis@nisd.net, mateo.macias@nisd.net, jennifer.christensen@nisd.net',
        //     driveLink: '1nK6IFB5SQkLmJUMo1cFkXob_ngbOECaM'
        //   };
        // case 'hobby':
        //   return {
        //     recipients: 'gregory.dylla@nisd.net, marian.johnson@nisd.net, jose.texidor@nisd.net, beverly.tiffany@nisd.net, lawrence.carranco@nisd.net, jennifer.castro@nisd.net, gabriela.becerra@nisd.net, christina.lora@nisd.net, jesus.alonzo@nisd.net, victoria.denton@nisd.net',
        //     driveLink: '1OmqKg2tjRmGxx_FbdzPm0seMdjaqDr1F'
        //   };
        // case 'hobby magnet':
        //   return {
        //     recipients: 'jaime.heye@nisd.net',
        //     driveLink: '1fYezfiqM1H5FHR_K0p4v2ro0OidExKZ-'
        //   };
        // case 'holmgreen':
        //   return {
        //     recipients: 'sandra.valles@nisd.net, yolanda.carlson@nisd.net',
        //     driveLink: '1nGHHinv5vS9O9XuVyKg1SPIzpleVbhoh'
        //   };
        // case 'jefferson':
        //   return {
        //     recipients: 'Nicole.aguirreGomez@nisd.net, Leti.Chapa@nisd.net, Justine.Hanauer@nisd.net, Monica.cabico@nisd.net, maria-1.martinez@nisd.net',
        //     driveLink: '1kKehUlXwjOnwTjTD0BOuw3L0XmvpXRRF'
        //   };
        // case 'jones':
        //   return {
        //     recipients: 'Javier.lazo@nisd.net, Beatriz.Ramirez@nisd.net, rudolph.arzola@nisd.net, lenni.peed@nisd.net, jesus.villela@nisd.net, loiselle.tejada@nisd.net',
        //     driveLink: '1BGe1rvOSZOmGkXU95l_ZiyODA2yTBFsZ'
        //   };
        // case 'jones magnet':
        //   return {
        //     recipients: 'xavier.maldonado@nisd.net',
        //     driveLink: '1L1roRy3NfUXIm7MmC76YeLo3F_9MhqIE'
        //   };
        // case 'jordan':
        //   return {
        //     recipients: 'Anabel.Romero@nisd.net, Shannon.Zavala@nisd.net, Laurel.Graham@nisd.net, Abigail.Almendarez@nisd.net, Robert.Ruiz@nisd.net, Ryanne.Barecky@nisd.net',
        //     driveLink: '1lOKHdlFWLqDMg4IAOglq1V6om_nHpjnD'
        //   };
        // case 'luna':
        //   return {
        //     recipients: 'moises.ochoa@nisd.net, Karl.Feuge@nisd.net, Amanda.king@nisd.net, Laura.walker@nisd.net, Sherry.mabry@nisd.net, Lisa.richard@nisd.net',
        //     driveLink: '1JBEzGoN44aZlfgw2ZOnMewGqTya0y-dN'
        //   };
        // case 'neff':
        //   return {
        //     recipients: 'Yvonne.Correa@nisd.net, Theresa.Heim@nisd.net, Joseph.Castellanos@nisd.net, Renada.Rodarte@nisd.net, Adriana.Aguero@nisd.net, Priscilla.Vela@nisd.net, michele.adkins@nisd.net, laura-i.sanroman@nisd.net, jennifer.cipollone@nisd.net, jessica.montoya@nisd.net',
        //     driveLink: ''
        //   };
        // case 'pease':
        //   return {
        //     recipients: 'Lynda.Desutter@nisd.net, tamara.campbell-babin@nisd.net, guadalupe.brister@nisd.net, mary.harrington@nisd.net, tiffany.flores@nisd.net, damaris.gutierrez@nisd.net, christie.nickell@nisd.net, arturo.bruns-vargas@nisd.net, mary.hernandez@nisd.net, kathleen.cuevas@nisd.net, tanya.alanis@nisd.net',
        //     driveLink: '1MvxU28snFNspmlryjeV27CgNYWsqGn6S'
        //   };
        // case 'rawlinson':
        //   return {
        //     recipients: 'Blanca.Martinez@nisd.net, Patti.Vlieger@nisd.net, nicole.buentello@nisd.net, elizabeth.smith@nisd.net',
        //     driveLink: '1j71vRF2rN4p75T2SVN_2_DovpfKKPn3s'
        //   };
        // case 'rayburn':
        //   return {
        //     recipients: 'Robert.Alvarado@nisd.net, Maricela.Garza@nisd.net, Aissa.Zambrano@nisd.net, Brandon.masters@nisd.net, jennifer-1.garza@nisd.net',
        //     driveLink: '10sEkjJf2XZ38zBgrnRrfmj5UwVrE_N5h'
        //   };
        // case 'ross':
        //   return {
        //     recipients: 'mahntie.reeves@nisd.net, christina.lozano@nisd.net, claudia.salazar@nisd.net, jason.padron@nisd.net, katherine.vela@nisd.net, faustino.ortega@nisd.net, ivette.ortega@nisd.net, dolores.cardenas@nisd.net',
        //     driveLink: '13C0WgGpwhm0N6WMuMplWeKdNbfl7kv9C'
        //   };
        // case 'rudder':
        //   return {
        //     recipients: 'kevin.vanlanham@nisd.net, ximena.huertamedina@nisd.net',
        //     driveLink: '1jl0pKTbYn16496dOK-AFt1WlpkW9fy64'
        //   };
        // case 'stevenson':
        //   return {
        //     recipients: 'julie.bearden@nisd.net, hilary.pilaczynski@nisd.net, johanna.davenport@nisd.net, amanda.cardenas@nisd.net, chaeleen.garcia@nisd.net ',
        //     driveLink: '1h6V7It6_hVdEjOlul6oIzCW8Jvcw-Xii'
        //   };
        // case 'stinson':
        //   return {
        //     recipients: 'Alexis.Montes@nisd.net, elda.garza@nisd.net, louis.villarreal@nisd.net, jeannette.rainey@nisd.net, crystal.lawrence@nisd.net, Linda.Boyett@nisd.net, lourdes.medina@nisd.net, rick.lane@nisd.net, maria.figueroa@nisd.net',
        //     driveLink: '1hvX1GY7pbgur_hNt58T7uZx_hCW5b496'
        //   };
        // case 'straus':
        //   return {
        //     recipients: 'DanaGilbert-Perry@nisd.net, wendi.peralta@nisd.net, BrandyBergeron@nisd.net, jennifer.myers@nisd.net',
        //     driveLink: '136mdaz16r8lonY5En6Rdc983hhji32k4'
        //   };
        // case 'vale':
        //   return {
        //     recipients: 'jenna.bloom@nisd.net, Brenda.rayburg@nisd.net, daniel.novosad@nisd.net, mary.harrington@nisd.net',
        //     driveLink: '1bNQsGIx-zqSXYRXHO3DBAu5sgACA0WPC'
        //   };
        // case 'zachry':
        //   return {
        //     recipients: 'Richard.DeLaGarza@nisd.net, Juliana.Molina@nisd.net, Randolph.neuenfeldt@nisd.net, jennifer.dever@nisd.net, veronica.poblano@nisd.net, monica.perez@nisd.net, jimann.caliva@nisd.net',
        //     driveLink: '14SAjOwdYe-LELYw7PWl8WlfyiV2h5Rgj'
        //   };
        // case 'zachry magnet':
        //   return {
        //     recipients: 'matthew.patty@nisd.net',
        //     driveLink: '1YgLbUHFu-aCO501mSg1rsMq8jUAqzP6k'
        //   };
        case 'test':
          return {
            recipients: 'alvaro.gomez@nisd.net', //, john.decker@nisd.net, sheila.yeager@nisd.net',
            driveLink: '1I7mkmBa3-sO_eG6f2KlSKyUg-G71ZAYK'
          };
        default:
          return {
            recipients: '',
            driveLink: ''
          };
      }
    }

  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);

  function getGmailTemplateFromDrafts_(row, driveLink) {
    return {
      message: {
        subject: 'AEP Placement Transition Plan',
        text: 'Hardcoded body text',
        html: `Dear ${row['Campus']},<br><br>
              ${row['Student Name']} has nearly completed their assigned placement at NAMS and should be returning to ${row['Campus']} on or around ${row['Return Date']}.<br><br>   
              On their last day of placement, they will be given withdrawal documents and the parents/guardians will have been called and told to contact ${row['Campus']} to set up an appointment to re-enroll and meet with an administrator/counselor.<br><br>  
              Below are links and attachments to a Personalized Transition Plan (with notes from NAMS' assigned social worker), the student's AEP Transition Plan (with grades and notes from their teachers at NAMS), and a link to ${row['Campus']}'s folder with all of the transition plans for this year.<br><br>
              Please let me know if you have any questions or concerns.<br><br>
              Thank you for all you do,<br>
              JD<br><br>
              ${row['Merged Doc URL - Personalized Transition Plan']}<br>
              ${row['Merged Doc URL - AEP Placement Transition Plan']}<br>
              https://drive.google.com/drive/folders/${driveLink}<br>`,
      }, 
      attachments: [], 
      inlineImages: {}
    };

  }

  function fillInTemplateFromObject_(template, row, driveLink) {
    // We have two templates one for plain text and the html body
    // Stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);
    // Token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      if (key === '${driveLink}') {
      return escapeData_(driveLink);
      }
      return escapeData_(row[key.replace(/[{}]+/g, "")] || "", driveLink);
    });
    return  JSON.parse(template_string);
  }

  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };
}

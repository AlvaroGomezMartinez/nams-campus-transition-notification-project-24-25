// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
 
/**
 * @OnlyCurrentDoc
*/
 
/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Date Campus Email Sent";
 
/** 
 * Creates the menu item "Notify Campuses" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Notify Campuses')
      .addItem('Send Emails', 'sendEmails')
      .addToUi();
}
 
/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(subjectLine = "AEP Placement Transition Plan", sheet=SpreadsheetApp.getActiveSheet()) {

  // Skipped the browser prompt below and just set the subject line
  // if (!subjectLine){
  //   subjectLine = Browser.inputBox("Send Campus Emails", 
  //                                     "Type or copy/paste the subject line of the Gmail " +
  //                                     "draft message you would like to mail merge with:",
  //                                     Browser.Buttons.OK_CANCEL);
                                      
  //   if (subjectLine === "cancel" || subjectLine == ""){ 
  //   // If no subject line, finishes up
  //   return;
  //   }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // Gets the data from the passed sheet
  const dataRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Return to HC 23-24').getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift(); 
  
  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        const campusInfo = getInfoByCampus(row['Campus']);
        const recipients = campusInfo.recipients;
        const driveLink = campusInfo.driveLink;

        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row, driveLink);

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(recipients, msgObj.subject, msgObj.text, {
          // htmlBody: msgObj.html,
          // bcc: 'a.bcc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          // attachments: emailTemplate.attachments,
          // inlineImages: emailTemplate.inlineImages
          htmlBody: msgObj.html,
        });
        // Edits cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  /**
     * Determine recipients based on the value in the "Campus" column.
     * @param {string} campusValue value in the "Campus" column
     * @return {object} containing a comma-separated list of recipients and Google Drive Folder link
     */
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
            driveLink: '1gBcGRl700LGfhMPrHx32dZ2j3MwPlqYS'
          };
        default:
          return '';
      }
    }

  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handles inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Gets all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Creates an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
/**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data, driveLink) {
    // We have two templates one for plain text and the html body
    // Stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);


    // Token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      if (key === '{{DriveLink}}') {
      return escapeData_(driveLink);
      }
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return  JSON.parse(template_string);
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
  */
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

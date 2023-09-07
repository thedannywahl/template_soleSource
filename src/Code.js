async function onFormSubmit(e) {
  const exports = `13PM-dUZJs1UTtj-3J9tPg4FmdJIBwOHT`
  const form = mapData(e)
  const template = getTemplate(form)
  const copyId = copyTemplate(template, exports, form)

  populateCopy(copyId, form)

  switch(form.format) {
    case 'Word':
      sendWord(copyId, form)
      break;
    case 'Google':
      sendGoogle(copyId, form)
      break;
    case 'PDF':
      await sendPDF(copyId, form)
      break;
    case 'Email':
      await sendFormattedEmail(copyId, form)
      break;
  }
}

function getTemplate(form) {
  let template
  switch (form.product) {
    case 'Instructure Learning Platform':
      template = `1mMPI5Sz1f2ypQuQSw-2ddYz8PXx1FwY8GU6R3xYyuN0`
      form.color = `#287A9F`
      break;
    case 'Canvas LMS':
      template = `1aZajnF810_BSYUb3r7OYTA6mr1e7ot_bIp2Ozu_uxkE`
      form.color = `#E72429`
      break;
    case 'Canvas Catalog':
      template = `15S9fusl8FvxdiDscOQPLwJw7vO9rVTHvy6NVVd6IM2U`
      form.color = `#E72429`
      break;
    case 'Canvas Studio':
      template = `1nd5atg__8fQKOHQxtH6HjAScnLedJLwI5Ou2VwDJBG0`
      form.color = `#E72429`
      break;
    case 'Canvas Credentials':
      template = `1ET7rK4zF1m6PMtc21mja4F_nITnUbbaYzx082EynPo4`
      form.color = `#E72429`
      break;
    case 'Mastery Connect':
      template = `1yz2wLgMnZYejUe9LPDIg8Usas9GcBR2WSi0sIaaV_yc`
      form.color = `#24A159`
      break;
    case 'Mastery View Assessments & Item Banks':
      template = `14EZbCDHuHmp117KNNyao1Yx84UtfJXHxSetLLHCHpSc`
      form.color = `#24A159`
      break;
    case 'Impact':
      template = `14AIKfnGFlBCG3b0u7VNF-Cx2Xky37xhKvhscplufSNk`
      form.color = `#F76400`
      break;
    case 'LearnPlatform':
      template = `1GaTMa_XR4jy4J27ByDIynOhBLbu33zdNiRkt6N_L1W4`
      form.color = `#0077cc`
      break;
    default:
      template = `1mMPI5Sz1f2ypQuQSw-2ddYz8PXx1FwY8GU6R3xYyuN0`
      form.color = `#287A9F`
  }
  return template
}
function mapData(e) {
  const form = {}
  form.date = e.values[0].split(" ")[0]
  form.clientname = e.values[1]
  form.primarycontact = e.values[2]
  form.rep = {
    name:  e.values[3],
    title: e.values[4],
    phone: e.values[5],
    email: e.values[6]
  }
  form.orderformId = e.values[7].length? e.values[7].split(`=`)[1] : false
  form.format = e.values[8]
  form.product = e.values[9]
  return form
}

function copyTemplate(templateId, copyId, form) {
  const id = DriveApp.getFileById(templateId)
                     .makeCopy(`${form.clientname} - ${form.product} - ${form.date}`, DriveApp.getFolderById(copyId))
                     .getId()
  return id
}

function populateCopy(id, form) {
  const copy = DocumentApp.openById(id)
  if (`To Whom It May Concern` !== form.primarycontact) {
    form.docGreeting = `Dear ${form.primarycontact},`
    form.emailGreeting = `Hi ${form.primarycontact},`
  } else {
    form.docGreeting = `${form.primarycontact}:`
    form.emailGreeting = `${form.primarycontact}:`
  }
  form.aOrAn = `a`
  form.theProduct = form.product
  if (form.product === `Instructure Learning Platform`) {
    form.aOrAn = `an`
    form.theProduct = `the ${form.product}`
  }
  if (form.product === `Impact`) {
    form.aOrAn = `an`
  }
 
  copy.getBody()
    .replaceText(`\\[product\\]`, form.product)
    .replaceText(`\\[clientname\\]`, form.clientname)
    .replaceText(`\\[primarycontact\\]`, form.docGreeting)
    .replaceText(`\\[rep\\.name\\]`, form.rep.name)
    .replaceText(`\\[rep\\.title\\]`, form.rep.title)
    .replaceText(`\\[rep\\.phone\\]`, form.rep.phone)
    .replaceText(`\\[date\\]`, form.date)
    .findText(`\\[rep\\.email\\]`)
      .getElement()
      .setText(form.rep.email)
      .setLinkUrl(`mailto:${form.rep.email}`)
      .setForegroundColor(form.color)
  copy.saveAndClose()
}

async function mergePDF(copyId, form) {
  const PDFexports = `13PQa3ziHHwp_k5rJep2_UAZbDZtydMQO`
  const data = [copyId, form.orderformId].map((id) => new Uint8Array(DriveApp.getFileById(id).getBlob().getBytes()))
  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText())
  const setTimeout = function(f, t) {
    Utilities.sleep(t)
    return f()
  }
  const pdfDoc = await PDFLib.PDFDocument.create()
  for (let i = 0; i < data.length; i++) {
    const pdfData = await PDFLib.PDFDocument.load(data[i])
    const pages = await pdfDoc.copyPages(pdfData, [...Array(pdfData.getPageCount())].map((_, i) => i))
    pages.forEach(page => pdfDoc.addPage(page))
  }
  const bytes = await pdfDoc.save()
  const mergedPDFId = DriveApp.createFile(Utilities.newBlob([...new Int8Array(bytes)],
                                          MimeType.PDF,
                                          `${form.product} sole source for ${form.clientname}`))
                      .getId()
  DriveApp.getFileById(mergedPDFId).moveTo(DriveApp.getFolderById(PDFexports))
  return mergedPDFId
}

function sendWord(copyId, form) {
  const body = sendWordBody(form, copyId)
  const html = HtmlService.createTemplateFromFile('repWordEmail')
  html.form = form
  html.copyId = copyId
  const htmlBody = html.evaluate().getContent()
  const attachments = []
  if (form.orderformId) attachments.push(DriveApp.getFileById(form.orderformId).getAs(MimeType.PDF))
  sendEmail(form, body, {htmlBody: htmlBody, attachments: attachments})
}

function sendGoogle(copyId, form) {
  DriveApp.getFileById(copyId).addEditor(form.rep.email)
  const body = sendGoogleBody(form, copyId)
  const html = HtmlService.createTemplateFromFile('repGoogleEmail')
  html.form = form
  html.copyId = copyId
  const htmlBody = html.evaluate().getContent()
  const attachments = []
  if (form.orderformId) attachments.push(DriveApp.getFileById(form.orderformId).getAs(MimeType.PDF))
  sendEmail(form, body, {htmlBody: htmlBody, attachments: attachments})
}

async function sendPDF(copyId, form) {
  const body = formattedPDFBody(form)
  const html = HtmlService.createTemplateFromFile('repPDFEmail')
  html.form = form
  html.copyId = copyId    
  const htmlBody = html.evaluate().getContent()
  const attachments = []
  const attachmentId = (form.orderformId) ? await mergePDF(copyId, form) : copyId
  attachments.push(DriveApp.getFileById(attachmentId).getAs(MimeType.PDF))
  sendEmail(form, body, {htmlBody: htmlBody, attachments: attachments})
}

async function sendFormattedEmail(copyId, form) {
  const body = formattedEmailBody(form)
  const html = HtmlService.createTemplateFromFile('clientReadyEmail')
  html.form = form
  html.copyId = copyId
  const htmlBody = html.evaluate().getContent()
  const images = {
    logo: DriveApp.getFileById(`13KSdIDdZa6lwY40QmMDXLhD8OrhdpZoN`).setName(`Instructure`),
    hero: DriveApp.getFileById(`13KyHE1F96I7x7hQv7JT14xUQD2v0ZgWY`).setName(`Hero`),
    fb:   DriveApp.getFileById(`13LrsqeL6-vK-cary9X4686BY33QeOI_b`).setName(`Facebook`),
    x:    DriveApp.getFileById(`13MKFDBc1I812MmU9ZguVmQsYpcSBr1ry`).setName(`X`),
    li:   DriveApp.getFileById(`13MqBdGfHQgtXvPtTPSCXqxb84asSQUjI`).setName(`LinkedIn`),
    yt:   DriveApp.getFileById(`13MAMRxu7xPusdW8Xs3EM5ABv4qoV45lh`).setName(`YouTube`)
  }
  const attachments = []
  const attachmentId = (form.orderformId) ? await mergePDF(copyId, form) : copyId
  attachments.push(DriveApp.getFileById(attachmentId).getAs(MimeType.PDF))
  sendEmail(form, body, {htmlBody: htmlBody, attachments: attachments, fromName: form.rep.name, inlineImages: images})
}

function sendEmail(form,
                   body,
                   {
                    htmlBody,
                    inlineImages = {},
                    attachments = [],
                    subject = `${form.product} Sole Source for ${form.clientname}`,
                    replyAddress = `rfps@instructure.com`,
                    fromName = `Instructure Proposal Team`,
                    bcc = `rfps+solesource@instructure.com`
                  } = {}) {
  GmailApp.sendEmail(
    form.rep.email,
    subject,
    body,
    {
      attachments: attachments,
      replyTo: replyAddress,
      name: fromName,
      htmlBody: htmlBody,
      inlineImages: inlineImages,
      bcc: bcc
  })
}

function sendWordBody(form, copyId) {
  return `Hi ${form.rep.name},
  
Your ${form.product} sole source for ${form.clientname} is ready for download: https://docs.google.com/document/d/${copyId}/export?format=docx
If you need any additional support, please ask in #proposal-writers

Thanks,

Instructure Proposal Team
rfps@instructure.com`
}

function sendGoogleBody(form, copyId) {
  return `Hi ${form.rep.name},
  
Your ${form.product} sole source for ${form.clientname} is ready: https://docs.google.com/document/d/${copyId}/view
If you need any additional support, please ask in #proposal-writers

Thanks,

Instructure Proposal Team
rfps@instructure.com`
}

function formattedPDFBody(form) {
  return `Hi ${form.rep.name},

Your ${form.product} sole source for ${form.clientname} is attached.
If you need any additional support, please ask in #proposal-writers.

Thanks,

Instructure Proposal Team
rfps@instructure.com`
}

function formattedEmailBody(form) {
return `Hi ${form.primarycontact},

Your ${form.product} sole source letter from Instructure is attached.
Please don't hesitate to reach out to me if you need anything.

Sincerely,

${form.rep.name}
${form.rep.title}
${form.rep.email}
${form.rep.phone}`
}
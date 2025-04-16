// Global configuration using Script Properties
function getConfig() {
  return {
    templateId: PropertiesService.getScriptProperties().getProperty('templateId') || '15gop9TL6YttWvCD060mrKoGbiLSyp8HXlJ8cS3L00vE',
    folderId: PropertiesService.getScriptProperties().getProperty('folderId') || '1bpcTXV96--uP0AmqfqD3tqFDZV7fyDfJ',
    emailRecipient: PropertiesService.getScriptProperties().getProperty('emailRecipient') || 'pro_sarawak@imi.gov.my'
  };
}

// Serve the web app
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Statistik Ketibaan dan Pelepasan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Country and category data
function getCountryData() {
  const countriesToSave = [
    "Malaysia", "Singapura", "Australia", "New Zealand", "Kanada", "United Kingdom",
    "Hong Kong", "Sri Lanka", "Bangladesh", "India", "Brunei Darussalam",
    "Amerika Syarikat", "China", "Russia", "Amerika Latin", "Negara Arab",
    "Jerman", "Perancis", "Norway, Sweden, Denmark, Finland",
    "Belgium, Luxemburg, Netherlands/Belanda", "Eropah", "Filipina",
    "Thailand", "Taiwan", "Indonesia", "Pakistan", "Jepun", "Korea",
    "Lain-lain Negara", "Vietnam"
  ];

  return {
    countriesToSave,
    "Malaysia": ["Malaysia A", "Malaysia H", "Malaysia K"],
    "Amerika Latin": [
      "Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Costa Rica", "Cuba",
      "Dominican Republic", "Ecuador", "El Salvador", "Guatemala", "Guyana", "Haiti",
      "Honduras", "Mexico", "Nicaragua", "Panama", "Paraguay", "Peru", "Uruguay", "Venezuela"
    ],
    "Negara Arab": [
      "Algeria", "Bahrain", "Mesir/Egypt", "Iraq", "Jordan", "Kuwait", "Lubnan",
      "Libya", "Mauritania", "Maghribi/Morocco", "Oman", "Palestin", "Qatar",
      "Arab Saudi", "Sudan", "Syria", "Tunisia", "Emiriah Arab Bersatu", "Yaman/Yemen"
    ],
    "Eropah": [
      "Austria", "Ireland", "Switzerland", "Czech Republic", "Hungary", "Poland",
      "Slovakia", "Slovenia", "Albania", "Bosnia and Herzegovina", "Bulgaria",
      "Croatia", "Greece", "Montenegro", "North Macedonia", "Romania", "Serbia",
      "Belarus", "Estonia", "Latvia", "Lithuania", "Moldova", "Ukraine",
      "Andorra", "Cyprus", "Iceland", "Liechtenstein", "Malta", "Monaco",
      "San Marino", "Vatican City", "Sepanyol", "Portugal"
    ],
    "Lain-lain Negara": [
      "Afrika Selatan", "Antigua dan Barbuda", "Armenia", "Azerbaijan", "Bahamas",
      "Barbados", "Belize", "Bhutan", "Botswana", "Burkina Faso", "Burundi",
      "Cambodia", "Cameroon", "Cape Verde", "Chad", "Comoros", "Djibouti",
      "Dominika", "Timor-Leste", "Eritrea", "Eswatini", "Ethiopia", "Fiji",
      "Gabon", "Gambia", "Georgia", "Ghana", "Grenada", "Guinea", "Guinea-Bissau",
      "Iran", "Jamaica", "Kazakhstan", "Kenya", "Kiribati", "Kyrgyzstan", "Laos",
      "Lesotho", "Liberia", "Madagascar", "Malawi", "Maldives", "Mali",
      "Marshall Islands", "Mauritius", "Micronesia", "Mongolia", "Mozambique",
      "Myanmar", "Namibia", "Nauru", "Nepal", "Nigeria", "Palau", "Papua New Guinea",
      "Rwanda", "Saint Kitts dan Nevis", "Saint Lucia", "Saint Vincent dan Grenadine",
      "Samoa", "Sao Tome dan Principe", "Senegal", "Seychelles", "Sierra Leone",
      "Kepulauan Solomon", "Somalia", "Sudan Selatan", "Suriname", "Tajikistan",
      "Tanzania", "Togo", "Tonga", "Trinidad dan Tobago", "Turkmenistan", "Turkey",
      "Tuvalu", "Uganda", "Uzbekistan", "Vanuatu", "Zambia", "Zimbabwe"
    ],
    "Norway, Sweden, Denmark, Finland": ["Norway", "Sweden", "Denmark", "Finland"],
    "Belgium, Luxemburg, Netherlands/Belanda": ["Belgium", "Luxemburg", "Belanda"]
  };
}

// Validate input data
function validateData(data) {
  if (!data.month || !/^(Januari|Februari|Mac|April|Mei|Jun|Julai|Ogos|September|Oktober|November|Disember)$/.test(data.month)) {
    throw new Error("Sila pilih bulan yang sah!");
  }
  if (!data.year || !/^\d{4}$/.test(data.year)) {
    throw new Error("Tahun mestilah dalam format YYYY!");
  }
  if (!data.office || data.office.trim() === '') {
    throw new Error("Sila masukkan nama pejabat!");
  }
  if (!data.countries || typeof data.countries !== 'object') {
    throw new Error("Data negara tidak lengkap!");
  }
}

// Save data to Google Sheets
function saveData(data) {
  try {
    validateData(data);
    Logger.log('Saving data for office: ' + data.office);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(data.office);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(data.office);
      sheet.appendRow(["Bulan", "Tahun", "Pejabat", "Negara", "Ketibaan", "Pelepasan"]);
    }

    const countryData = getCountryData();
    countryData.countriesToSave.forEach(country => {
      sheet.appendRow([
        data.month,
        data.year,
        data.office,
        country,
        Number(data.countries[country]?.ketibaan || 0),
        Number(data.countries[country]?.pelepasan || 0)
      ]);
    });

    return "Data berjaya disimpan untuk \"" + data.office + "\"!";
  } catch (e) {
    Logger.log('Error saving data: ' + e.message);
    throw new Error("Gagal menyimpan data: " + e.message);
  }
}

// Generate PDF from template
function generatePDF(data) {
  try {
    validateData(data);
    const config = getConfig();
    Logger.log('Generating PDF for: ' + data.office);
    const templateFile = DriveApp.getFileById(config.templateId);
    const newFile = templateFile.makeCopy('Statistik Ketibaan dan Pelepasan - ' + data.office + ' - ' + data.month + ' ' + data.year);
    const newDoc = DocumentApp.openById(newFile.getId());
    const body = newDoc.getBody();

    // Replace placeholders
    body.replaceText("{PEJABAT}", data.office.toUpperCase());
    body.replaceText("{BULAN}", data.month.toUpperCase());
    body.replaceText("{TAHUN}", data.year);

    // Get tables
    const tables = body.getTables();
    if (tables.length < 2) {
      throw new Error("Templat memerlukan sekurang-kurangnya 2 jadual!");
    }

    // Table 1: Malaysia
    const malaysiaTable = tables[0];
    let malaysiaTotalTiba = 0;
    let malaysiaTotalPergi = 0;
    const malaysiaCountries = getCountryData()["Malaysia"];
    if (malaysiaTable.getNumRows() < malaysiaCountries.length + 2) {
      throw new Error("Jadual Malaysia tidak mempunyai baris yang mencukupi!");
    }
    malaysiaCountries.forEach((country, index) => {
      const row = malaysiaTable.getRow(index + 1);
      row.getCell(0).setText((index + 1).toString());
      row.getCell(1).setText(country);
      const tiba = Number(data.countries[country]?.ketibaan || 0);
      const pergi = Number(data.countries[country]?.pelepasan || 0);
      row.getCell(2).setText(tiba.toLocaleString('en-US'));
      row.getCell(3).setText(pergi.toLocaleString('en-US'));
      malaysiaTotalTiba += tiba;
      malaysiaTotalPergi += pergi;
    });
    const malaysiaTotalRow = malaysiaTable.getRow(malaysiaCountries.length + 1);
    malaysiaTotalRow.getCell(0).setText("");
    malaysiaTotalRow.getCell(2).setText(malaysiaTotalTiba.toLocaleString('en-US'));
    malaysiaTotalRow.getCell(3).setText(malaysiaTotalPergi.toLocaleString('en-US'));

    // Table 2: Other Countries
    const otherTable = tables[1];
    const otherCountries = getCountryData().countriesToSave.filter(c => c !== "Malaysia");
    if (otherTable.getNumRows() < otherCountries.length + 3) {
      throw new Error("Jadual Negara Lain tidak mempunyai baris yang mencukupi!");
    }
    let otherTotalTiba = 0;
    let otherTotalPergi = 0;
    otherCountries.forEach((country, index) => {
      const row = otherTable.getRow(index + 1);
      row.getCell(0).setText((index + 1).toString());
      row.getCell(1).setText(country.toUpperCase());
      const tiba = Number(data.countries[country]?.ketibaan || 0);
      const pergi = Number(data.countries[country]?.pelepasan || 0);
      row.getCell(2).setText(tiba.toLocaleString('en-US'));
      row.getCell(3).setText(pergi.toLocaleString('en-US'));
      otherTotalTiba += tiba;
      otherTotalPergi += pergi;
    });

    // Foreign Visitors Total
    const otherTotalRow = otherTable.getRow(otherCountries.length + 1);
    otherTotalRow.getCell(0).setText("");
    otherTotalRow.getCell(2).setText(otherTotalTiba.toLocaleString('en-US'));
    otherTotalRow.getCell(3).setText(otherTotalPergi.toLocaleString('en-US'));

    // Grand Total
    const grandTotalRow = otherTable.getRow(otherCountries.length + 2);
    grandTotalRow.getCell(0).setText("");
    grandTotalRow.getCell(2).setText((otherTotalTiba + malaysiaTotalTiba).toLocaleString('en-US'));
    grandTotalRow.getCell(3).setText((otherTotalPergi + malaysiaTotalPergi).toLocaleString('en-US'));

    newDoc.saveAndClose();

    // Save PDF to folder
    const folder = DriveApp.getFolderById(config.folderId);
    const pdfFile = newFile.getAs('application/pdf');
    const pdfName = 'Statistik Ketibaan dan Pelepasan - ' + data.office + ' - ' + data.month + ' ' + data.year + '.pdf';
    const savedFile = folder.createFile(pdfFile.setName(pdfName));
    savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Delete temporary Docs file
    newFile.setTrashed(true);

    Logger.log('PDF generated: ' + savedFile.getUrl());
    return savedFile.getUrl();
  } catch (e) {
    Logger.log('Error generating PDF: ' + e.message);
    throw new Error("Gagal menjana PDF: " + e.message);
  }
}

// Send email with PDF attachment
function sendEmail(data) {
  try {
    validateData(data);
    const config = getConfig();
    Logger.log('Sending email for: ' + data.office);
    const pdfName = 'Statistik Ketibaan dan Pelepasan - ' + data.office + ' - ' + data.month + ' ' + data.year + '.pdf';
    const folder = DriveApp.getFolderById(config.folderId);
    const files = folder.getFilesByName(pdfName);
    if (!files.hasNext()) {
      throw new Error("Fail PDF tidak dijumpai!");
    }
    const pdfFile = files.next().getAs('application/pdf');

    MailApp.sendEmail({
      to: config.emailRecipient,
      subject: pdfName,
      body: `Salam,\n\nBerikut adalah statistik ketibaan dan pelepasan untuk ${data.office} bagi bulan ${data.month} ${data.year}.\n\nTerima kasih.`,
      attachments: [pdfFile.setName(pdfName)]
    });

    Logger.log('Email sent to: ' + config.emailRecipient);
    return "Emel berjaya dihantar ke " + config.emailRecipient + "!";
  } catch (e) {
    Logger.log('Error sending email: ' + e.message);
    throw new Error("Gagal menghantar emel: " + e.message);
  }
}

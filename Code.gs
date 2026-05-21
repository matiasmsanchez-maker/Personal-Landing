function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Landing MSB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getLatestFromFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var latest = null;
  var latestDate = '';

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();

    var fullMatch = name.match(/(\d{4}\.\d{2}\.\d{2})/);
    var yearMonthDash = name.match(/(\d{4})-(\d{2})(?!-\d)/);
    var shortMatch = name.match(/(\d{2})\.(\d{2})(?!\.\d)/);

    var dateStr = '';
    if (fullMatch) {
      dateStr = fullMatch[1];
    } else if (yearMonthDash) {
      dateStr = yearMonthDash[1] + '.' + yearMonthDash[2];
    } else if (shortMatch) {
      dateStr = '20' + shortMatch[2] + '.' + shortMatch[1];
    }

    if (dateStr && dateStr > latestDate) {
      latestDate = dateStr;
      latest = { name: name, url: file.getUrl(), date: dateStr };
    }
  }

  return latest;
}

function getLatestSemanalN1() {
  return getLatestFromFolder('1prphI25hs8oJBSrPbBNdFFm9xoWwyOy5');
}

function getLatestB2B2C() {
  return getLatestFromFolder('1Xlg9cUQ5cYN9LqYk_IYNuqkuM6wb13KV');
}

function getLatestCreditRisk() {
  return getLatestFromFolder('1i5eh-3Dohwjte8zFoLT_niSZC7fdAw3n');
}

function getLatestMedia() {
  return getLatestFromFolder('1Gh9MZ8xm2GxWTqF00wqGdwpWjeI8Awro');
}

function getLatestDocByPrefix(prefix) {
  var folder = DriveApp.getFolderById('1Po2EnLn1hegK5nwxgAXCIO7jhAB6eepj');
  var files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  var latest = null;
  var latestTime = 0;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(prefix) === 0) {
      var modified = file.getLastUpdated().getTime();
      if (modified > latestTime) {
        latestTime = modified;
        latest = file;
      }
    }
  }

  if (!latest) return null;

  var doc = DocumentApp.openById(latest.getId());
  var body = doc.getBody().getText();

  return {
    title: latest.getName(),
    content: body,
    url: latest.getUrl()
  };
}

function getDailyMail() {
  return getLatestDocByPrefix('Daily Summary');
}

function getMinutas() {
  return getLatestDocByPrefix('Minutas Pendientes');
}

function getPlanning() {
  var docId = '13GYQCwrP0y96ZF0cudr2_YyMaeTkFyVZnGXNP42dDy8';
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var numChildren = body.getNumChildren();
  var datePattern = /^\d{1,2}[-\/](ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic|jan|apr|aug|dec)/i;

  var sections = [];
  var current = null;

  for (var i = 0; i < numChildren; i++) {
    var el = body.getChild(i);
    var type = el.getType();
    var text = el.asText().getText().trim();
    if (!text) continue;

    if (type === DocumentApp.ElementType.PARAGRAPH && datePattern.test(text)) {
      current = { title: text, bullets: [], links: [] };
      sections.push(current);
      continue;
    }

    if (current && (type === DocumentApp.ElementType.LIST_ITEM || type === DocumentApp.ElementType.PARAGRAPH)) {
      current.bullets.push(text);

      var richText = el.asText();
      var indices = richText.getTextAttributeIndices();
      for (var j = 0; j < indices.length; j++) {
        var url = richText.getLinkUrl(indices[j]);
        if (url) {
          var start = indices[j];
          var end = (j + 1 < indices.length) ? indices[j + 1] : text.length;
          var linkText = text.substring(start, end).trim();
          var exists = current.links.some(function(l) { return l.url === url; });
          if (!exists && linkText) {
            current.links.push({ text: linkText, url: url });
          }
        }
      }
    }
  }

  if (sections.length === 0) return null;

  var first = sections[0];

  return {
    title: first.title,
    content: first.bullets.join('\n'),
    links: first.links,
    docUrl: 'https://docs.google.com/document/d/' + docId + '/edit'
  };
}

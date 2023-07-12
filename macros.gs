function RichTextAuto() {
  var spreadsheet = SpreadsheetApp.getActive();
  var lastRow = spreadsheet.getLastRow();
  var lastCol = spreadsheet.getLastColumn();
  var columns = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
    'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V',
    'W', 'Y', 'Z'
  ];
  for (var row = 0; row < lastRow; row++) {
    for (var col = 0; col < lastCol; col++) {
      var ncell = `${columns[col]}${row + 1}`;
      var cell = spreadsheet.getRangeByName(ncell);
      if ( ! /<[^>]+>/.test(cell.getValue()) ) continue;
      
      var richText = RichTextFormat(cell.getValue());
      cell.setRichTextValue(richText);
    }
  }
};

function RichTextFormat(text) {
    text = text.replace(/\t|\r|\n|<(p|div)> *&nbsp; *<\/(p|div)>/g, '').replace(/[ \t]{2,}/g, ' ').replace(/>( |\t)</g, '><');
    text = decodeHTMLEntities(text).normalize('NFC');

    text = text.replace(/<\/(p|div)>|<br *\/?>/g,'\n').replace(/<\/?(p|div)( [^>]*)?>/g,'');

    var lists  = text.match(/<(o|u)l( [^>]+)?>[^<]*(<li( [^>]+)?>[^<]+<\/li>)+[^<]*<\/(o|u)l>/g);
    if (lists)
    for (l of lists) {
        var fn = /<ol/.test(l) ? ((i) => ++i) : ((i) => '*');
        var list = [];
        var i = 0;
        for (var li of l.match(/<li( [^>]*)?>[^<]+<\/li>/g)) {
            list.push('  ' + (/<ol/.test(l) ? ++i + '.' : '*') + ' ' + li.replace(/<li( [^>]*)?>|<\/li>/g, ''));
        }
        text = text.replace(l, list.join('\n') + '\n');
    }

    var rich = {bolds: [], italics: [], heads: [] }
    var _headers = {
      '<h1>': 32,
      '<h2>': 24,
      '<h3>': 18.72,
      '<h4>': 16,
      '<h5>': 13.28,
      '<h6>': 12
    }
    var italics = text.match(/<(em|i)( [^>]+)?>[^<]+<\/(em|i)>/g);
    var bolds = text.match(/<(b|strong)( [^>]+)?>([^<]+)<\/(b|strong)>/g);
    var heads = text.match(/<h\d( [^>]+)?>[^<]+<\/h\d>/g);
  
    if (bolds)
    for (b of bolds) {
        var i = text.indexOf(b);
        var l = b.replace(/<(b|strong)( [^>]+)?>|<\/(b|strong)>/g, '');
        rich.bolds.push([i, i + l.length]);
        text = text.replace(b, l);
    }
    
    if (italics)
    for (b of italics) {
        var i = text.indexOf(b);
        var l = b.replace(/<(em|i)( [^>]+)?>|<\/(em|i)>/g, '');
        rich.italics.push([i, i + l.length]);
        text = text.replace(b, l);
    }
  
    if (heads)
    for (h of heads) {
        var i = text.indexOf(h);
        var l = h.replace(/<h\d( [^>]+)?>|<\/h\d>/g, '');
        var k = h.replace(/(<h\d).*/,'$1>');
        var px = _headers[k];
        rich.heads.push([i, i + l.length, px]);
        text = text.replace(h, l + '\n');
    }
  
    var newRichText = SpreadsheetApp.newRichTextValue();
    newRichText.setText(text);
    for (b of rich.bolds) {
        newRichText.setTextStyle(b[0], b[1], SpreadsheetApp
            .newTextStyle()
            .setBold(true)
            .build()
        );
    }
    for (i of rich.italics) {
        newRichText.setTextStyle(i[0], i[1], SpreadsheetApp
            .newTextStyle()
            .setItalic(true)
            .build()
        );    
    }
    for (h of rich.heads) {
        newRichText.setTextStyle(h[0], h[1], SpreadsheetApp
            .newTextStyle()
            .setBold(true)
            .setFontSize(h[2])
            .build()
        ); 
    }
    return newRichText.build();
};

function decodeHTMLEntities(text) {
    var entities = {
        '&amp;': '&',
        '&lt;': '<',
        '&gt;': '>',
        '&quot;': '"',
        '&iexcl;': '¡',
        '&iquest;': '¿',
        '&#39;': "'",
        '&#x2F;': '/',
        '&#x60;': '`',
        '&#x3D;': '=',
        '&nbsp;': ' ',
        '&aacute;': 'á',
        '&eacute;': 'é',
        '&iacute;': 'í',
        '&oacute;': 'ó',
        '&uacute;': 'ú',
        '&ntilde;': 'ñ',
        '&Aacute;': 'Á',
        '&Eacute;': 'É',
        '&Iacute;': 'Í',
        '&Oacute;': 'Ó',
        '&Uacute;': 'Ú',
        '&Ntilde;': 'Ñ'
    };
    var search = Object.keys(entities).join('|');
    var regex = new RegExp(search, 'g');
    return text.replace(regex, function(match, entity) {
        return entities[match] || '';
    });
};

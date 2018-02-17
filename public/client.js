/** drop target **/
//var _target = document.getElementById('drop');
var CSF = XLSX;

var _target = document.body;
var events = ""; // Calendar events ICS file contents
/** Spinner **/
var spinner;

var _workstart = function() { spinner = new Spinner().spin(_target); }
var _workend = function() { spinner.stop(); }

/** Alerts **/
var _badfile = function() {
  alertify.alert('This file does not appear to be a valid Excel file.  If we made a mistake, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can take a look.', function(){});
};

var _pending = function() {
  alertify.alert('Please wait until the current file is processed.', function(){});
};

var _large = function(len, cb) {
  alertify.confirm("This file is " + len + " bytes and may take a few moments.  Your browser may lock up during this process.  Shall we play?", cb);
};

var _failed = function(e) {
  console.log(e, e.stack);
  alertify.alert('We unfortunately dropped the ball here.  We noticed some issues with the grid recently, so please test the file using the direct parsers for <a href="/js-xls/">XLS</a> and <a href="/js-xlsx/">XLSX</a> files.  If there are issues with the direct parsers, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can make things right.', function(){});
};

/** Handsontable magic **/
var boldRenderer = function (instance, td, row, col, prop, value, cellProperties) {
  Handsontable.TextCell.renderer.apply(this, arguments);
  $(td).css({'font-weight': 'bold'});
};

var $container, $parent, $window, availableWidth, availableHeight;
var $sql, $sqlpre, sqldb;

var calculateSize = function () {
  var offset = $container.offset();
  availableWidth = Math.max($window.width() - 250,600);
  availableHeight = Math.max($window.height() - 250, 400); 
};

$(document).ready(function() {
  $container = $("#hot"); $parent = $container.parent();
  $window = $(window);
  $window.on('resize', calculateSize);
  $sqlpre = document.getElementById('sqlpre'); 
  $sql = document.getElementById('sql');
  $buttons = document.getElementById('buttons');
});

var _onsheet = function(json, cols) {
  $('#footnote').hide();
  /* add header row for table */
  if(!json) json = [];
  // header row commented
  //json.unshift(function(head){var o = {}; for(i=0;i!=head.length;++i) o[head[i]] = head[i]; return o;}(cols));
  calculateSize();
  /* showtime! */
  $("#hot").handsontable({
    data: json,
    startRows: 5,
    startCols: 3,
    fixedRowsTop: 1,
    stretchH: 'all',
    rowHeaders: true,
    columns: cols.map(function(x) { return {data:x}; }),
    colHeaders: cols,
    cells: function (r,c,p) {
      if(r === 0) this.renderer = boldRenderer;
    },
    width: function () { return availableWidth; },
    height: function () { return availableHeight; },
    stretchH: 'all'
  });
};

/* Result (data) of SQL query to screen:
  SQLResultSet
  SQLResultSetRowList
  Length
*/
function sexqlit(data) {
  console.log('DATA: ', data, data.rows, data.rows.length);
  if(!data || data.length === 0) return; 
  var r = data.rows.item(0);
  var cols = Object.keys(r);
  var json = [];
  for(var i = 0; i < data.rows.length; ++i) {
    r = data.rows.item(i);
    var o = {};
    cols.forEach(function(x) { o[x] = r[x]; });
    json.push(o);
  }
  console.log('JSON: ');
  console.log(json,cols,cols.length);
  _onsheet(json, cols);
  // write events
  writeEvents(json, cols);
  
}


/* SQL query execution

*/
function sexql() {
  $sqlpre.classList.remove("error");
  $sqlpre.classList.remove("info");
  var stmt = $sql.value;
  if(!stmt) return;
  if(stmt.indexOf(";") > -1) stmt = stmt.substr(0, stmt.indexOf(";"));
  $sqlpre.innerText = stmt;
  
  sqldb.transaction(function(tx) {
    tx.executeSql(stmt,[], function(tx, results) {
    $sqlpre.classList.add("info");
      sexqlit(results);
      writeresultFile(results);
    }, function(tx, e) {
    $sqlpre.innerText += "\n" + e + "\n" + (e.message||"") +"\n"+ (e.stack||"");
    $sqlpre.classList.add("error");
    }); 
  });
}

/*
SQL queries
s: query
*/
function prepstmt(s) {
  console.log(s); // INSERT INTO `VK46` 
   // (`Alkaa`, `Päättyy`, `Joukkue`, `Harjoitus`, `Paikka`, `Valmentaja`) VALUES ("15:45","16:50","IE","Jää","Tuplajäät","Outi&Netta");
  sqldb.transaction(function(tx) { 
    tx.executeSql(s, [], function(){}, function(){console.log(arguments); }); 
  });
}

/* Data to db, preprosesses worksheets
   ws: Work sheet
   sname: Sheet name
*/
function prepforsexql(ws, sname) {
  console.log('Sheet: ', sname, '\n');
  if(!ws || !ws['!ref']) return;
  
  var range = CSF.utils.decode_range(ws['!ref']);
  console.log('Range: ', range, '\n');
  if(!range || !range.s || !range.e || range.s > range.e) return;
  /* resolve types */
  
  range.e.c=7;

  var types = new Array(range.e.c-range.s.c+1); 
  
  // names: column names
  var names = new Array(range.e.c-range.s.c+1);  
  var R = range.s.r; // r:row
  for(var C = range.s.c; C<= range.e.c; ++C)
    names[C-range.s.c] = (ws[CSF.utils.encode_cell({c:C,r:R})]||{}).w;
    names[(C-1)-range.s.c] = 'EA'; //b Col 8 EA-column, "ensiapuvastaavat" can be added
  for(var C = range.s.c; C<= range.e.c; ++C)
    for(R = range.s.r+1; R<= range.e.r; ++R)
      switch((ws[CSF.utils.encode_cell({c:C,r:R})]||{}).t) {
        case 'e': break; /* error type */
        case 'b': /* boolean -> number */
        case 'n': if(types[C-range.s.c] !== "TEXT") types[C-range.s.c] = "TEXT"; break;
        case 's': case 'str': types[C-range.s.c] = "TEXT";
        default: break; /* if the cell doesnt exist */
      }
  // for each worksheet

  /* update list, sheet name (sname) */
  //$buttons.innerHTML += "<h2>`" + sname + "`</h2>"
  // ss: header names (ei tarvita kalenterisovelluksessa)
  var ss = ""
  names.forEach(function(n) { if(n) ss += "`" + n + "`<br />"; });
  //$buttons.innerHTML += "<h3>" + ss + "</h3>";
  /* create table */
  prepstmt("CREATE TABLE `" + sname + "` (" + names.map(function(n, i) { return "`" + n + "` " + (types[i]||"TEXT"); }).join(", ") + ");" );
  prepstmt("DROP TABLE `" + sname + "`" );
  prepstmt("CREATE TABLE `" + sname + "` (" + names.map(function(n, i) { return "`" + n + "` " + (types[i]||"TEXT"); }).join(", ") + ");" );

  /* insert data */
  var eventDate = null; // date for each cell even if it is empty in the original xlsx
  //var emptyRow = false;
  for(R = range.s.r+1; R<= range.e.r; ++R) {
    var fields = [], values = [];
    for(var C = range.s.c; C<= range.e.c; ++C) {
      var cell = ws[CSF.utils.encode_cell({c:C,r:R})];
      var nextCell = null;
      if (C+1<= range.e.c) {
        var nextCell = ws[CSF.utils.encode_cell({c:C+1,r:R})];
      }
      if(!cell) {
        if (!nextCell) continue;
      }
      if(cell == null) {
        console.log('NULL\n');
        if (nextCell == null) continue;
      }     
      // last given date to cell A
      if (C===0) {  // Empty date cell
         if (!cell) {
          cell = eventDate;
        } else {  // Date cell with new date
          eventDate = cell;
        }
      }
      fields.push("`" + names[C-range.s.c] + "`");
      // Some junk in cells , empty spaces  etc.
      if (!(cell == null))
        values.push(types[C-range.s.c] === "REAL" ? cell.v : '"' + cell.w + '"');
    }  
    prepstmt("INSERT INTO `" +sname+ "` (" + fields.join(", ") + ") VALUES (" + values.join(",") + ");");
    
  }

}

/*
  Format date (Ma 10.3 -> 20170310)

*/
function formatDate(date) {
  var year = (new Date()).getFullYear();
  var re = /\d+\.\d+/i; // E.g. Ma 3.10 (i: upper/lower case ignored)
  var daymonth = date.match(re);
  
  var tmpmonth = daymonth[0].match(/\.\d+/);
  var tmpday = daymonth[0].match(/\d+\./);
  var month = pad(tmpmonth[0].replace(/^\./,""), 2);
  var day = pad(tmpday[0].replace(/\.$/,""), 2);
 
  //return year.toString() + month + day;
  return '2018' + month + day;
}

function pad(num, size) {
  var s = num+"";
  while (s.length < size) s = "0" + s;
  return s;
}

function removeNbsp(str) {
  var newStr = str.replace(/\u00a0/g, ' ');
  return newStr;
}

/* Workbook: data to db, query executed
  wb: workbook
*/
function sexqlify(wb) {
  $buttons.innerHTML = "";
  document.getElementById('sql').oninput = sexql; // sexql - executes query 

  $sexqls = document.getElementById('sexqls');
  if(typeof openDatabase === 'undefined') {
    sqldiv.innerHTML = '<div class="error"><b>*** WebSQL not available.  Try the <a href="sqljs.html">SQL.js demo</a> ***</b></div>';
    return;
  }
  sqldb = openDatabase('sheetjs','1.0', 'sexql', 3 * 1024 * 1024);

  // Query for all worksheets to find team IA events
  var sheetUnionQuery = "SELECT * FROM (";
  // Worksheets to db
  wb.SheetNames.forEach(function(s, idx, array) { 
    prepforsexql(wb.Sheets[s], s); 
    var num = idx+1; // order number of work sheet to keep order in query results (for filter)
    sheetUnionQuery += "SELECT * , " + num + " AS FILTER FROM " + s + " WHERE trim(Joukkue) LIKE '%IA%' ";
    if (idx < array.length - 1) { 
      sheetUnionQuery += " UNION ALL ";
    }
  });
  sheetUnionQuery += " ) ORDER BY FILTER";
  if(wb.Sheets && wb.Sheets.Data && wb.Sheets.Metadata) 
  document.getElementById('sql').value = "SELECT Format, Importance as Priority, Data.Code, css_class FROM Data JOIN Metadata ON Metadata.code = Data.code WHERE Importance < 3";
  else 
   
   document.getElementById('sql').value = sheetUnionQuery;
   
  sexql(); // SQL query execution, result to the screen

  $sql.disabled = true;
}


var _onwb = function(wb, type, sheetidx) {
  sexqlify(wb);
}

function writeEvents(json, cols) {
  var events = "BEGIN:VCALENDAR\r\nPRODID:-//Paula Leinonen//js-excel2ics//EN\r\nVERSION:1.0\r\n";
  events += "CALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\nX-WR-CALNAME:Luistelu\r\nX-WR-TIMEZONE:Europe/Helsinki\r\n";
  
  for (var i = 0; i < json.length; i++) { // Each row = one event
    var obj = json[i]; // json array with objects
    events += "BEGIN:VEVENT\r\n";
    // DTSTAMP:20171027T183131Z
    var timestamp = new Date();
    var tsyear = timestamp.getFullYear();
    var tsmonth = pad(timestamp.getMonth() + 1,2);
    var tsdate = pad(timestamp.getDate(),2);
    var tshour = pad(timestamp.getHours(),2);
    var tsmin = pad(timestamp.getMinutes(),2);
    var tssec = pad(timestamp.getSeconds(),2);
    var ts = tsyear.toString() + tsmonth.toString() + tsdate.toString() + "T" + tshour.toString() +'0000';



    
    events += "DTSTAMP:" + ts + "\n";
    // UID: unique ID
    events += "UID:" + ts + "-" + i + "\r\n";
    events += "CLASS:PRIVATE\r\n";
    events += "CREATED:" + ts + "\r\n";
    events += "CATEGORIES:1\r\n";

    var coach = "";
    var location = "";
    var starttime = "";
    var endtime = "";
    var tmpstarttime = "";
    var tmpendtime = "";
    var ea = "";

    for (var key in obj) {
      if (obj.hasOwnProperty(key)) {
        var val = obj[key];
        if (val) {
          if (cols[0]===key) {  // Päivä
            var date = formatDate(val);
          } 
        } else {
            continue; // val undefined or null
        }
         switch (key) 
        {
           case cols[0]:  // Päivä
              //var date = formatDate(val);
              events += "DTSTART:" + date;
              break;
        // DTSTART:20171125T120000Z
          case cols[1]: // Alkaa start time
              // find hour
              // pad: adds leading zero
              var tmpshour = val.match(/^\d+/); 
              var shour = pad(parseInt(tmpshour),2); 
              var smin = pad(val.match(/\d+$/),2);
              starttime = shour + ":" + smin;
              tmpstarttime = tmpshour + ":" + smin;
              events += "T" + shour + smin + "00\r\n"; 
              break;
          case cols[2]: // Päättyy End time   
              var tmpehour = val.match(/^\d+/); 
              var ehour = pad(parseInt(tmpehour),2);
              var emin = pad(val.match(/\d+$/),2);
              endtime = ehour + ":" + emin;
              tmpendtime = tmpehour  + ":" + emin;
              events += "DTEND:" + date + "T" + ehour + emin + "00\r\n";
              break;
          case cols[4]: // Harjoitus e.g. SUMMARY:Oheinen (Tua) 12:00-12:45
            eventname = encode_utf8(val);
            break;
          case cols[5]: //  Paikka
            location = encode_utf8(val);
            break;
          case cols[6]: //  Valmentaja
            coach = val;
            break;
          case cols[7]: //  Ensiapuvastaava (EA)
            ea = removeNbsp(val);
            break;
          default: 
               //alert('Default case');
               break;
        } // switch
      } // if (obj.hasOwnProperty(key))
    } // each col in line

         /*  DESCRIPTION:Tua

    LAST-MODIFIED:20171027T183101Z
    LOCATION:Niiralan Monttu\, Hannes Kolehmaisen katu 4\, 70110 Kuopio\, Suomi
    LOCATION:Tuplajäät
URL;TYPE=Map:http://maps.google.com/maps?q=1055%20Fifth
  %20Avenue%2C%20San%20Diego%2C%20California%20&hl=en
*/
    events += "LOCATION:" + location + "\r\n";
    /*var map = getLink(location);
    if (map) {
      events += "URL;TYPE=Map:" + map + "\r\n";
    } */
    events += "DESCRIPTION:" + coach;
    if (ea) {
      events += " - EA: " + ea + "\r\n"; 
    } else 
    {
      events += "\r\n" ;
    }
    events += "SEQUENCE:0\r\n";
    events += "STATUS:CONFIRMED\r\n";
    // SUMMARY:Oheinen (name) 12:00-12:45
    events += "SUMMARY:" + eventname + " (" + coach + ") " + tmpstarttime + "-" + tmpendtime + "\r\n";
    events += "TRANSP:OPAQUE\r\n";
    events += "END:VEVENT\r\n";
    
  };
  events += "END:VCALENDAR\r\n";
  addEvents(events);
  //console.log('EVENTS: ' + events); 
}

//no need for encode/decode
function encode_utf8( s ) {
  var result = [];
  var temp = "";
  for (var i = 0; i < s.length; i++) {
    var charCode = s.charCodeAt(i);
      
      if (s[i] === 'ä') {
          temp += "ä";
      } else {
        temp += s[i];
      }
      //result.push(temp + String.fromCharCode(charCode));
  }
  return temp;
  //return unescape(encodeURIComponent( s ));
}

function decode_utf8( s ) {
  return decodeURIComponent( escape( s ) );
}


/*function getEventName(name) {
  switch (name) {
    case "Jää":
      return "Jää"; 
      break;
    default:
      return name;
      break;
  }
} */

function getLink(loc) {
  switch (loc) {
    case "Niirala":
      return "https://www.google.fi/maps/place/Niiralan+Monttu+Hannes+Kolehmaisen+katu+Kuopio+Suomi/@62.8952734,27.6639871,17z/?hl=fi";
      break;
    case "Lippumäki":
      return "https://www.google.fi/maps/place/Lippumäen+uimahalli/@62.8409009,27.6446969,17z/?hl=fi";
      break;
    case "Toivala", "Tuplajäät":
      return "https://www.google.fi/maps/place/Tuplajäät+Oy/@62.9822246,27.7180837,17z/?hl=fi";
      break;
    default:
      return "";
  }

}

/* Write results to ICalendar file */

function writeresultFile(result) {
  //console.log('RESULT ' + result.rows[1]);
  var textFile = null,
    makeTextFile = function (text) {
      var data = new Blob([text], {type: 'text/plain'});
  
      // If we are replacing a previously generated file we need to
      // manually revoke the object URL to avoid memory leaks.
      if (textFile !== null) {
        window.URL.revokeObjectURL(textFile);
      }
  
      textFile = window.URL.createObjectURL(data);
  
      return textFile;
    };
  
  
    var create = document.getElementById('create'),
      textbox = document.getElementById('textbox');
  
    create.addEventListener('click', function () {

      var link = document.createElement('a');
      link.setAttribute('download', 'events.ics');
      link.href = makeTextFile(textbox.value);
      document.body.appendChild(link);
  
      // wait for the link to be added to the document
      window.requestAnimationFrame(function () {
        var event = new MouseEvent('click');
        link.dispatchEvent(event);
        document.body.removeChild(link);
      });
      
    }, false);
  };
  




/** Drop it like it's hot **/
DropSheet({
  drop: _target,
  on: {
    workstart: _workstart,
    workend: _workend,
    sheet: _onsheet,
    wb: _onwb,
    foo: 'bar'
  },
  errors: {
    badfile: _badfile,
    pending: _pending,
    failed: _failed,
    large: _large,
    foo: 'bar'
  }
})

function addEvents(events){
  var eventBox = document.getElementById("textbox");
  eventBox.value = events;
}

/*
<script type="text/javascript">
  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-36810333-1']);
  _gaq.push(['_setDomainName', 'sheetjs.com']);
  _gaq.push(['_setAllowLinker', true]);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();
  </script>
*/   
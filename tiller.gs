function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Match Against Everything', functionName: 'matchAll'},
    {name: 'Calculate Rents', functionName: 'calculateRents'},
    {name: 'Match Categories and Rules', functionName: 'matchCategoriesAndRules'},
    {name: 'Match Against Amazon', functionName: 'matchAmazon'},
    {name: 'Match Against Mint', functionName: 'matchMint'}
  ];
  spreadsheet.addMenu('Financial', menuItems);
}

function regexEscape(s) {
    // Replaces CHAR with \CHAR to keep every regex check perfectly literal
    return s.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
};

function matchCategoriesAndRules() {
  autoCategorize(true,true,true,false,false);
}

function matchAmazon() {
  autoCategorize(true,false,false,true,false);
}

function matchMint() {
  autoCategorize(true,false,false,false,true);
}

function matchAll() {
  autoCategorize(true,true,true,true,true);
}

function matchTest() {
  autoCategorize(false,true,true,true,true,603,604);
}

var months = "Jan01Feb02Mar03Apr04May05Jun06Jul07Aug08Sep09Oct10Nov11Dec12";
function quickISODate(s) {
  // Tue Nov 04 2017
  s = s.toString();
  var month = months.substr( months.indexOf( s.substr(4,3) )+3, 2 );
  return s.substr(11,4)+'-'+month+'-'+s.substr(8,2);
}

function getVarName(name) {
  name = name.replace(/ /g,'');
  return name.charAt(0).toLowerCase() + name.slice(1);
};

//
// On sheet find each header in the nameList and return a 'zip' (my name) that
// contains a pre-cache of all the values, and the ability to set values, formulas,
// notes, colors, etc using member methods.
//
function zipFetch(sheet,nameList,doFormulas) {
  this.sheet = sheet;
  this.cache = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  if( doFormulas ) {
    this.formula = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getFormulas();
  }

  var headersFound = false;
  for( var row=0 ; !headersFound && row < this.cache.length ; ++row ) {
    var varName;
    for( var i=0 ; i<nameList.length ; ++i ) {
      varName = getVarName(nameList[i]);
      this[varName] = this.cache[row].indexOf(nameList[i]);
      if( this[varName]<0 ) break;
    }
    headersFound = i>=nameList.length && this[varName]>=0;
  }
  if( !headersFound ) {
    Browser.msgBox('Error', 'No header row found containing "'+nameList.join('", "')+'"', Browser.Buttons.OK);
    return;
  }
  this.rowHeader = row-1;
  this.rowStart = row+1-1;
  this.set = function(rowInCache,colInCache,newValue,alignment) {
    this.sheet.getRange(rowInCache+1,colInCache+1).setValue(newValue);
    this.cache[rowInCache][colInCache] = newValue;
    if( alignment ) {
      this.sheet.getRange(rowInCache+1,colInCache+1).setHorizontalAlignment(alignment);
    }
  }
  this.setFormula = function(rowInCache,colInCache,newFormula,alignment) {
    this.sheet.getRange(rowInCache+1,colInCache+1).setFormula(newFormula);
    this.formula[rowInCache][colInCache] = newFormula;
    if( alignment ) {
      this.sheet.getRange(rowInCache+1,colInCache+1).setHorizontalAlignment(alignment);
    }
  }
  this.setNote = function(rowInCache,colInCache,text,append) {
    var range = this.sheet.getRange(rowInCache+1,colInCache+1);
    range.setNote((append ? range.getNote():'')+text);
  }
  this.setFgColor = function(rowInCache,colInCache,color) {
    this.sheet.getRange(rowInCache+1,colInCache+1).setFontColor(color);
  }
  this.setBgColor = function(rowInCache,colInCache,color) {
    this.sheet.getRange(rowInCache+1,colInCache+1).setBackground(color);
  }
  return this;
}

//
// calculateRents()
// Using a tricky method of "filling in" all the rents back to the start date, this simply
// determines how much rent has been given to us by each tenant up to the current date
//
function calculateRents() {
  var spreadsheet = SpreadsheetApp.getActive();

  var transSheet = spreadsheet.getSheetByName('Transactions');
  if( !transSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Transactions".', Browser.Buttons.OK);
    return;
  }

  var rentalsSheet = spreadsheet.getSheetByName('Rentals');
  if( !rentalsSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Rentals".', Browser.Buttons.OK);
    return;
  }

  rentalsSheet.activate();

  //
  // Expects the Rentals sheet to have a column "Unit" that has eg Prairie A, Topper B
  //
  function buildUnitHash(sheet) {
    var z = new zipFetch(sheet,["Unit"]);
    if( !z ) return;

    var anyFound = false;
    var hash = {};
    for( var row=z.rowStart ; row < z.cache.length; ++row ) {
      var v = z.cache[row];
      if( v[z.unit]=='' ) break;
      hash[v[z.unit]] = 1;
      anyFound = true;
    }
    if( !anyFound ) {
      Browser.msgBox('Error', 'No units were found under the Units header.', Browser.Buttons.OK);
      return;
    }
    return hash;
  }
  
  //
  // Builds a transaction list using a list of all rental Units found
  //
  function buildTrans(sheet,unitHash,unitList) {
    var z = new zipFetch(transSheet,["Date", "Brief Description", "Amount", "Category"]);
    if( !z ) return;

    var unitList = [];
    for( var unitName in unitHash ) {
      unitList.push(unitName);
    }
    var reg = new RegExp( 'Rent ('+unitList.join('|')+')' );
    
    var errorList = '';
    var trans = { 'z': z };
    z.unitName = z.cache[0].length;
    z.allocated = z.cache[0].length+1
    for( var row=z.rowStart; row<z.cache.length ; ++row ) {
      v = z.cache[row];
      if( v[z.category] == 'Rental Income' ) {
        var m = (''+v[z.briefDescription]).match( reg );
        if( m != null ) {
          var unit = (''+m[1]).trim();
          v[z.date] = quickISODate(v[z.date]);
          v[z.unit] = unit;
          v[z.allocated] = false;
          trans[unit] = trans[unit] || [];
          trans[unit].unshift(v);
        }
      }
    }
    if( errorList ) {
      Browser.msgBox('Error', 'Errors: '+errorList, Browser.Buttons.OK);
      return;
    }
    return trans;
  }
  
  function toDays(date,offset) {
    var day = 1000*60*60*24;
    var d = new Date(date);
    var result = (d.getTime()/day)+(offset||0);
    return Math.floor(result);
  }
  
  var unitHash = buildUnitHash(rentalsSheet);
  if( !unitHash ) return;
  var trans = buildTrans(transSheet,unitHash);
  if( !trans ) return;

  var z = new zipFetch(rentalsSheet,["Unit","Start Date","End Date","Rent","Late Fee"],true);
  if( !z ) return;
  
  var colorInactive = '#CCCCCC';
  var colorActive = '#FFFFFF';
  var colorPaidOnTimeInFull = '#77CC77';
  var colorPaidOnTimePartial = '#FF5555';
  var colorPaidLateInFull = '#7777CC';
  var colorPaidLateRentOnly = '#C68131';
  var colorPaidLatePartial = '#FF5555';
  var colorPaidNever = '#FF5555';

  //
  // Plugs in the believed amount of rent paid against a certain month.
  //
  function setAmount(row,col,amount) {
    var total = (z.cache[row][col] || 0)+amount;
    var formula = z.formula[row][col] ? z.formula[row][col]+'+'+amount : '='+amount;
    z.cache[row][col] = total;
    z.setFormula(row,col,formula);
    return total;
  }
  
  var balance = [];
  
  var monthColFirst = z.lateFee+2;
  for( var monthCol = monthColFirst ; monthCol < z.cache[0].length ; ++monthCol ) {
    var month = z.cache[z.rowHeader][monthCol];
    if( !month ) {
      break;
    }
    var monthISO = quickISODate(month);
    var monthNext = z.cache[z.rowHeader][monthCol+1];
    if( !monthNext ) {
      break;
    }
    var monthNextISO = quickISODate(monthNext);
    var monthNextDate = toDays(monthNext);
    
    for( var row=z.rowStart ; row<z.cache.length ; ++row ) {
      var v = z.cache[row];
      var unit = v[z.unit].trim();
      if( !unit ) break;
      var dateStart = v[z.startDate];
      var dateEnd = v[z.endDate];
      
      // This allows a tenant to start mid-month, but it does NOT auto-calculate their pro-rate rent. Yet.
      var lastDayOfMonth = new Date(monthNext);
      lastDayOfMonth.setDate(lastDayOfMonth.getDate()-1);
      if( lastDayOfMonth < dateStart || month > dateEnd ) {
        z.set(row,monthCol,'-','right');
        continue;
      }
      balance[row] = balance[row] || 0;
      z.set(row,monthCol,'');
      z.setFormula(row,monthCol,'');
      z.setNote(row,monthCol,'');
      var rent = v[z.rent];
      var lateFee = v[z.lateFee];

      var total = 0;
      if( balance[row] > 0 ) {
        var amount = Math.min(balance[row],rent);
        balance[row] -= amount;
        total = setAmount(row,monthCol,amount);
        z.setNote(row,monthCol,''+total+' credit\n',true); // Money was carried forward from prior payments
        z.setFgColor(row,monthCol,total >= rent ? colorPaidOnTimeInFull : colorPaidOnTimePartial);
      }
      
      var found = false;
      var transList = trans[unit];
      var tz = trans.z;
      // This alorithm assumes that the transactions are sorted ascending!
      for( var i=0 ; i<transList.length ; ++i ) {
        var t = transList[i];
        if( !t[tz.allocated] ) {
          var lateDays = 7;
          var transactionDate = toDays(t[tz.date])
          var lateDate = toDays(month,lateDays);
          if( transactionDate < monthNextDate ) {
            if( unit == 'Topper B' ) {
              var qqq = 1;
            }
            var newBalance = balance[row]+t[tz.amount];
            t[tz.allocated] = 1;
            var amount = Math.min(newBalance,rent-total); // This might not be right because a late fee might apply
            newBalance -= amount;
            balance[row] = newBalance;
            if( amount ) {
              total = setAmount(row,monthCol,amount);
              z.setNote(row,monthCol,''+amount+' paid\n',true);
              if( transactionDate < lateDate ) {
                z.setFgColor(row,monthCol,total >= rent ? colorPaidOnTimeInFull : colorPaidOnTimePartial);
              }
              else {
                z.setFgColor(row,monthCol,total >= rent+lateFee ? colorPaidLateInFull : total == rent ? colorPaidLateRentOnly : colorPaidLatePartial);
              }
            }
          }
        }
      }
      if( total == 0 ) {
        total = setAmount(row,monthCol,0);
        z.setFgColor(row,monthCol,colorPaidNever);
      }
      
      // WARNING: Maybe all this should happend above, when amount == 0 and there is a balance, search back and apply it where you can.
      var n = 1;
      while( balance[row] > 0 && monthCol-n>=monthColFirst  ) {
        if( z.cache[row][monthCol-n] < rent+0 ) {
          // Apply the balance against a prior rent
          var unpaid = (rent+0) - z.cache[row][monthCol-n];
          var amount = Math.min(balance[row],unpaid);
          balance[row] = balance[row]-amount;
          if( amount ) {
            total = setAmount(row,monthCol-n,amount);
            z.setFgColor(row,monthCol-n,total >= rent+lateFee ? colorPaidLateInFull : total == rent ? colorPaidLateRentOnly : colorPaidLatePartial);
            z.setNote(row,monthCol-n,''+amount+' paid\n',true);
          }
        }
        n++;
      }
      if( balance[row] > 0 ) {
        z.setNote(row,monthCol,'balance '+balance[row]+'\n',true);
      }
    }
  }
}

function autoCategorize(doAll,doCats,doRules,doAmazon,doMint,startOnRow,endOnRow) {
  var spreadsheet = SpreadsheetApp.getActive();
  var catSheet = spreadsheet.getSheetByName('Categories');
  if( !catSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Categories".', Browser.Buttons.OK);
    return;
  }

  var ruleSheet = spreadsheet.getSheetByName('Rules');
  if( !ruleSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Rules".', Browser.Buttons.OK);
    return;
  }

  var amazonSheet = spreadsheet.getSheetByName('Amazon');
  if( !amazonSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Amazon".', Browser.Buttons.OK);
    return;
  }

  var mintSheet = spreadsheet.getSheetByName('Mint');
  if( !mintSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Mint".', Browser.Buttons.OK);
    return;
  }

  var transSheet = spreadsheet.getSheetByName('Transactions');
  if( !transSheet ) {
    Browser.msgBox('Error', 'Missing a sheet called "Transactions".', Browser.Buttons.OK);
    return;
  }
  transSheet.activate();

  //
  //
  // buildCategoryHash
  //
  //
  function buildCategoryHash(sheet) {
    var z = new zipFetch(sheet,["Category"]);
    if( !z ) return;

    var hash = {};
    for( var row=z.rowStart ; row < z.cache.length; ++row ) {
      var v = z.cache[row];
      if( v[z.category]!='' ) {
        hash[v[z.category]] = 1;
      }
    }
    return hash;
  }

  
  //
  //
  // buildMap (actually means the Category Map
  // The $ in a description means comma, since comma is used as the separator
  // If a mapping contains catToDetect->catNameToUse then the catNameToUse will be used instead
  //
  //
  function buildMap(sheet) {
    var z = new zipFetch(sheet,["Category","Description Match"]);
    if( !z ) return;

    var tempCategoryHash = {};
    var map = [];
    var briefs = {};
    //var debug = '';
    for( var row=z.rowStart ; row < z.cache.length; ++row ) {
      var v = z.cache[row];
      if( !v[z.category] || v[z.category]=='' ) {
        break;
      }
      var descriptionMatchRaw =  ''+v[z.descriptionMatch];
      var descriptionMatch = '';
      var descriptionReplace = null;
      if( descriptionMatchRaw ) {
        var descPairs = descriptionMatchRaw.replace(/\\,/g,"CoMmA").split(',');
        var descRegex = [];
        for( var i=0 ; i<descPairs.length ; ++i ) {
          var pair = descPairs[i].replace(/CoMmA/g,",").split('->');
          pair[0] = pair[0].trim();
          // This may seem weird, but we are escaping the description so that the later regex check finds only literally what the user typed. It does not accept actual regex in the description match.
          descRegex.push(regexEscape(pair[0]));
          descriptionReplace = descriptionReplace || {};
          descriptionReplace[pair[0].replace(/^/g,'').toLowerCase()] = pair[1] || pair[0];          
        }
        descriptionMatch = new RegExp( '('+descRegex.join('|')+')', 'i' );
      }
      // It is IMPORTANT that every category be represented, even if empty, so that the list of all categories created later will be complete.
      var record = {"category": v[z.category], "descriptionMatch": descriptionMatch, "descriptionReplace": descriptionReplace };
      tempCategoryHash[record.category] = record;
      map.push(record);
    }
    
    // This is a nasty hack to try to accomodate PayPal's habit of putting the words "id:someCategory" into their descriptions
    map.unshift(tempCategoryHash["Transfer"]);
    
    return map;
  }
  
  //
  //
  // buildRules
  //
  //
  function buildRules(sheet) {
    var z = new zipFetch(sheet,["Description Match","Min Amount","Max Amount","New Category","New Description"]);
    if( !z ) return;

    var list = [];
    for( var row=z.rowStart ; row < z.cache.length ; ++row ) {
      var v = z.cache[row];
      if( v[z.descriptionMatch]==='' && v[z.minAmount]==='') {
        continue;
      }
      if( v[z.descriptionMatch] != '' ) {
        v[z.descriptionMatch] = new RegExp( v[z.descriptionMatch], 'i' );
      }
      list.push(v);
    }
    return { 'z': z, 'rule': list };
  }
  
  function clone(obj) {
    if (null == obj || "object" != typeof obj) return obj;
    var copy = obj.constructor();
    for (var attr in obj) {
        if (obj.hasOwnProperty(attr)) copy[attr] = obj[attr];
    }
    return copy;
  }
  
  function floatToPennies(amount) {
      var pennies = ''+amount;
      if( pennies.indexOf('.') < 0 ) {
        pennies += '.00';
      }
      else
      if( pennies.charAt(pennies.length-2) == '.' ) {
        pennies += '0';
      }
      return parseInt(pennies.replace('.',''));
  }
  
  //
  //
  // buildAmazon
  //
  //
  function buildAmazon(sheet) {
    var z = new zipFetch(sheet,["Order Date","Order ID","Title","Seller","Quantity","Item Total"]);
    if( !z ) return;

    z.itemPennies = z.cache[0].length;
    var orders = { 'z': z, 'raw': [], 'byID': {}, 'byDate': {}, 'byPennies': {} };
    for( var row=z.rowStart ; row < z.cache.length ; ++row ) {
      var v = z.cache[row];
      v[z.itemPennies] = floatToPennies(v[z.itemTotal]);
      v[z.orderDate] = quickISODate(v[z.orderDate]);
      var orderId = v[z.orderID];
      if( !orders.byID[orderId] ) {
        orders.byID[orderId] = clone(v);
      }
      else {
        var r = orders.byID[orderId];
        r[z.title] += " AND "+v[z.title];
        r[z.itemTotal] = r[z.itemTotal] + v[z.itemTotal];
        r[z.itemPennies] = r[z.itemPennies] + v[z.itemPennies];
        r[z.quantity] += v[z.quantity];
        orders.byID[orderId] = r;
      }
    }
    for( var id in orders.byID ) {
      var record = orders.byID[id];
      orders.byDate[record[z.orderDate]] = orders.byDate[record[z.orderDate]] || [];
      orders.byDate[record[z.orderDate]].push(record);
      orders.byPennies[record[z.itemPennies]] = orders.byPennies[record[z.itemPennies]] || [];
      orders.byPennies[record[z.itemPennies]].push(record);
    }
    return orders;
  }
  
  //
  // Old function used during the initial conversion. Should not be needed anymore.
  //
  function buildMint(sheet) {
    
    function getMintCategoryToTillerMap() {
      var z = new zipFetch(catSheet,["Category","Mint Equivalents"]);
      if( !z ) return;
      
      var mintMap = {};
      for( var row=z.rowStart ; row < z.cache.length; ++row ) {
        var v = z.cache[row];
        if( !v[z.category] || v[z.category]=='' ) {
          continue;
        }
        if( v[z.mintEquivalents] ) {
          var equivList = v[z.mintEquivalents].split(',');
          for( var i=0 ; i<equivList.length ; ++i ) {
            mintMap[equivList[i].trim()] = v[z.category].trim();
          }
        }
      }
      return mintMap;
    }
    
    var z = new zipFetch(sheet,["Date","Description","Original Description","Amount","Category","Transaction Type"]);
    if( !z ) return;
    var mintToTiller = getMintCategoryToTillerMap();
    z.tillerCategory = z.cache[0].length;
    
    var mint = {z:{}};
    mint.z.description = z.description;
    mint.z.tillerCategory = z.tillerCategory;
    for( var row=z.rowStart ; row<z.cache.length ; ++row ) {
      var v = z.cache[row];
      var mintCategory = v[z.category].trim();
      v[z.tillerCategory] = mintToTiller[mintCategory] || ""; //"fail: "+v[z.category];
      var summary = quickISODate(v[z.date])+':'+(v[z.transactionType]=='debit'?-v[z.amount]:v[z.amount]); //+':'+v[z.originalDescription].toString().replace(/ /g,'').substr(0,10);
      mint[summary] = v;
    }
    return mint;
  }
  
  function descriptionAndCategoryDetermine(description,fullDescription,amount,date) {
    
    var foundDescription = null;
    
    //
    // Test Amazon
    //
    if( doAmazon && amazon && (description.toLowerCase().indexOf("amazon mktplace")>=0 || description.toLowerCase().indexOf("amazon.com")>=0) ) {
      var itemPennies = floatToPennies(amount < 0 ? -amount : amount);
      var aa = amazon;
      var list = amazon.byPennies[itemPennies];
      if( list ) {
        var transactionDateTime = new Date(date).getTime();
        var dayTolerance = 7;  // We only match to a transaction if it happened within seven days of our transaction date
        var oneDay = 1000*60*60*24;
        var bestRecord = list[0];
        var bestDistance = -1;
        for( var i=0 ; i<list.length ; ++i ) {
          var distance = Math.abs(new Date(list[i][amazon.z.orderDate]).getTime() - transactionDateTime)/oneDay;
          if( (bestDistance==-1 || distance < bestDistance) && distance < dayTolerance ) {
            bestRecord = list[i];
            bestDistance = distance;
          }
        }
        if( bestDistance != -1 ) {
          foundDescription = 'Amazon '+bestRecord[amazon.z.orderDate]+' '+bestRecord[amazon.z.title].substr(0,50);
          description = foundDescription;
        }
      }
    }
    
    //
    // Test Rules
    //
    if( doRules ) {
      for( var i=0 ; i<rules.rule.length ; ++i ) {
        var rule = rules.rule[i];
        var descriptionMatch = rule[rules.z.descriptionMatch];
        var minAmount = rule[rules.z.minAmount];
        var maxAmount = rule[rules.z.maxAmount];
        var match = true;
        // There is a debate in my mind. When the description has been tweaked, I think we
        // can only rely fully on the fullDescription, so we check rules against that...
        match = match && (descriptionMatch=='' || (''+fullDescription).search(descriptionMatch)>=0);
        if( maxAmount === 'negative min' && minAmount !== '' ) {
          match = match && (amount === minAmount || amount === -minAmount);
        }
        else
          if( maxAmount==='' && minAmount!=='' ) {
            match = match && (amount === minAmount);
          }
        else {
          match = match && (minAmount==='' || amount >= minAmount);
          match = match && (maxAmount==='' || amount <= maxAmount);
        }
        if( match ) {
          var rr = rules;
          return { newCategory: rule[rules.z.newCategory], newDescription: rule[rules.z.newDescription] };
        }
      }
    }
    
    //
    // Test Category mappings
    // If a mapping contains catToDetect->catNameToUse then the catNameToUse will be used instead
    //
    if( doCats ) {
      for( var i=0 ; i<map.length ; ++i ) {
        var descriptionMatch = map[i].descriptionMatch;
        if( descriptionMatch ) {
          var m = description.match(descriptionMatch);
          if( m != null ) {
            return { newCategory: map[i].category, newDescription: map[i].descriptionReplace[m[0].toLowerCase()] || m[0] };
          }
        }
      }
    }
    
    //
    // Test Mint
    //
    if( doMint && mint ) {
      var summary = quickISODate(date)+':'+amount; //+':'+fullDescription.replace(/ /g,'').substr(0,10);
      var mm = mint;
      var ms = mint[summary];
      if( mint[summary] ) {
        return { 'newDescription': mint[summary][mint.z.description], 'newCategory': mint[summary][mint.z.tillerCategory] };
      }
    }
    
    if( foundDescription ) {
      return { 'newDescription': foundDescription };
    }
    
    return null;
  }
  
  //
  // Tokenizing allows !symbol! substitution as follows:
  // !mmm! - three letter month name
  // !yy! - two letter year name
  //
  var replaceData = null;  
  function tokenize(s,date,amount) {
    
    replaceData = replaceData || {
      mmm: function() {
        return date.toString().substr(4,3);
      },
      yy: function() {
        return date.toString().substr(13,2);
      }
    };
    
    var result = s.replace( /!([a-zA-z]+)!/g, function(match, p1, offset, string) {
      var key = p1.toLowerCase();
      if( key && replaceData[key] ) {
        return replaceData[key]();
      }
      return key;
    });
    return result;    
  }  

  var map,rules,amazon,mint;

  var categoryHash = buildCategoryHash(catSheet);
  
  if( doCats ) {
    //catSheet.activate();
    map = buildMap(catSheet);
    if( !map ) return;
  }
  
  if( doRules ) {
    //ruleSheet.activate();
    rules = buildRules(ruleSheet);
    if( !rules ) return;
  }
  
  if( doAmazon ) {
    //amazonSheet.activate();
    amazon = buildAmazon(amazonSheet);
  }
  
  if( doMint ) {
    //mintSheet.activate();
    mint = buildMint(mintSheet);
  }
  
  transSheet.activate();

  var z = new zipFetch(transSheet,["Date", "Brief Description", "Description", "Full Description", "Amount", "Category"]);
  if( !z ) return;

  var row = startOnRow ? startOnRow-1 : 1;
  while( row < z.cache.length ) {
    // Use search row & col to offset to a cell relative to dateRange start
    //if( row % 10 == 0 ) transSheet.setActiveRange(transSheet.getRange(row,1,1,1));
    
    var vOld = z.cache[row];
    //
    // Weak categories and descriptions are eligible to be replaced with better data.
    //
    var weakDescription = vOld[z.briefDescription] == '' || vOld[z.briefDescription] == vOld[z.description];
    var weakCategory = vOld[z.category] == '' || !categoryHash[vOld[z.category]];
    
    if( weakDescription || weakCategory ) {
      var vNew = descriptionAndCategoryDetermine( vOld[z.description], vOld[z.fullDescription], vOld[z.amount], vOld[z.date] ); // these we do NOT pass because something is weird and mint[1900] we getting truncated to mint[101], map, rules, amazon, mint );

      if( vNew && vNew.newCategory && weakCategory ) {
        z.set(row,z.category,vNew.newCategory);
      }
      if( vNew && vNew.newDescription && weakDescription ) {
        //
        // The descriptions allow !token! substitution
        //
        var briefDescription = vNew.newDescription.indexOf('!')>=0 ? tokenize(vNew.newDescription,vOld[z.date],vOld[z.amount]) : vNew.newDescription;
        z.set(row,z.briefDescription,briefDescription);
        vOld[z.briefDescription] = briefDescription;
      }
      if( vOld[z.briefDescription] == '' ) {
        z.set(row,z.briefDescription,vOld[z.description]);
      }
      //z.set(row,z.briefDescription,"howdy!");
    }
    ++row;
    if( endOnRow && row >= endOnRow ) {
      break;
    }
    if( !doAll && !endOnRow ) break;
  } 
}

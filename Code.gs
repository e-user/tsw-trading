//TSW Trading Script
//Copyright (C) 2014  Alexander Kahl <e-user@fsfe.org>
//
//This program is free software: you can redistribute it and/or modify
//it under the terms of the GNU Affero General Public License as
//published by the Free Software Foundation, either version 3 of the
//License, or (at your option) any later version.
//
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU Affero General Public License for more details.
//
//You should have received a copy of the GNU Affero General Public License
//along with this program.  If not, see <http://www.gnu.org/licenses/>.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Script')
    .addItem('Update Script', 'updateScript')
    .addToUi();
}

function rangeList (range, horizontal) {
  var result = [];
  var i = 1;
  var val = range.getCell(1, 1).getValue();
  var get_value = horizontal 
    ? function (i) { return range.getCell(1, i).getValue(); }
    : function (i) { return range.getCell(i, 1).getValue(); };

  while (val != "") {
    result.push(val);
    i++;
    val = get_value(i);
  }
  
  return result;  
}

function namedRangeList (name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = spreadsheet.getRangeByName(name);
  
  return rangeList(range);
}

function findCellValue (range, value, horizontal) {
  var _ = Underscore.load();
  return _.indexOf(rangeList(range, horizontal), value);
}

function rarities () {
  var _ = Underscore.load();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rarities");
  var rarities = rangeList(spreadsheet.getRange("A:A"));
  var colors = rangeList(spreadsheet.getRange("B:B"));
  
  return _.object(rarities, colors);
}

function types () {
  var _ = Underscore.load();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("types");
  var types = rangeList(spreadsheet.getRange("A:A"));
  var format = rangeList(spreadsheet.getRange("B:B"));
  
  return _.object(types, format);
}

function selling (table) {
  var _ = UnderscoreString.load(Underscore.load(), true);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(table);
  var items = rangeList(spreadsheet.getRange("A:A"));
  var data = spreadsheet.getRange(1, 2, _.max([items.length, 1]), 4).getValues();
  var _rarities = rarities();
  var _types = types();
  
  return _.isEmpty(items) ? [] : _.map(_.zip(items, data), function (vals) {
    var [name, [amount, rarity, price, type]] = vals;
    return {
      name: _.sprintf(_types[type], name), amount: amount, rarity: rarity, color: _rarities[rarity], price: price, type: type
    };
  });  
}

function formatItem (item) {
  var _ = UnderscoreString.load(Underscore.load(), true);
  return _.sprintf("<font face=HUGE>%d x <font face=HUGE color=%s>%s</font> @ %s</font>", item.amount, item.color, item.name, item.price);
}

function updateScript () {
  var _ = UnderscoreString.load(Underscore.load(), true);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Script");
  
  var sort = function (data) {
    return _.chain(data).values().map(function (items) {
      return _.sortBy(items, 'rarity');
    }).flatten().value();
  };
  
  var sellingItems = sort(_.groupBy(selling("Selling"), 'type'));
  var buyingItems = sort(_.groupBy(selling("Buying"), 'type'));
  
  var cell = spreadsheet.getRange(2, 1).getCell(1, 1);
  var template = spreadsheet.getRange(1, 1).getCell(1, 1).getValue();
  
  var s = '<font face=HEADLINE>Selling</face><br>';
  
  s = _.reduce(sellingItems, function (s, item) {
    return s += formatItem(item) + '<br>';
  }, s);
  
  s += '<br><font face=HEADLINE>Buying</face><br>';
  
  s = _.reduce(buyingItems, function (s, item) {
    return s += formatItem(item) + '<br>';
  }, s);
  
  cell.setValue(_.sprintf(template, s, (new Date()).toLocaleDateString('en-US')));
}

function test () {
  //Logger.log(rarities());
  //Logger.log(selling());
  updateScript();
}

/**
 * DND Beyond Character API Script
 * by Jacen Crudo
 * 
 * VERSION v1.0,
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('DND Beyond Tools')
    .addItem('Refresh Character Data', 'refreshCharacterData')
    .addSubMenu(ui.createMenu('Admin Tools')
      .addItem('Inititalize Sheets', 'initSheets')
      .addItem('Clear Event Logs', 'clearLogs')
    )
  .addToUi();

}

/** FUNCTION AUTO REFRESH FUNCTION */
function checkAutoRefresh(){
  try {
    if(SpreadsheetApp.getActiveSpreadsheet().getRangeByName('charAutoRefresh').getValue()) refreshCharacterData();
  }
  catch(e) {
    addEvent('SYSTEM', 'ERROR', '', '', e)
  }
}

/** Function to intitialize all required sheets */
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(!ss.getSheetByName('Characters')) createCharactersSheet();
  if(!ss.getSheetByName('Settings')) createSettingsSheet();
  if(!ss.getSheetByName('Events')) createEventsSheet();
}

/** Function to clear all logs and reset event log IDs */
function clearLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getSheetByName('Events')) {
    ss.deleteSheet(ss.getSheetByName('Events'));
    ss.getRangeByName('eventID').setValue(0);
  }
  createEventsSheet();
}

/** Function to create the Settings Sheet */
function createSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSS = ss.insertSheet('Settings', 0);
  var players = settingsSS.getRange(1, 1, 20, 2).setHorizontalAlignment('center');
  var settings = settingsSS.getRange(1, 3, 20, 3).setHorizontalAlignment('center');
  
  settingsSS.deleteColumns(6, 21);
  settingsSS.deleteRows(21, 980);

  // DND Player Columns
  ss.setNamedRange('players', players);
  settingsSS.setColumnWidths(1, 2, 150);
  players.getBandings().forEach(banding => banding.remove());
  players.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
  settingsSS.getRange(1, 1, 1, 2).setValues([['Player', 'ID']]).setFontWeight('bold');
  
  // Settings Columns
  ss.setNamedRange('settings', settings);
  ss.setNamedRange('charAutoRefresh', settingsSS.getRange(2, 5));
  settingsSS.getRange(2, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  ss.setNamedRange('charLastRefresh', settingsSS.getRange(3, 5));
  ss.setNamedRange('eventID', settingsSS.getRange(4, 5));
  settingsSS.setColumnWidths(3, 3, 200);
  settings.getBandings().forEach(banding => banding.remove());
  settings.applyRowBanding();
  settingsSS.getRange(1, 3, 1, 3).setValues([['Setting', 'Modifier', 'Value']]).setFontWeight('bold');
  settingsSS.getRange(2, 3, 3, 3).setValues([['Auto Refresh', '', ''], ['Last Sys Refresh', '', ''], ['Last Event ID', '', 0]]);
}

/** Function to create the Events Sheet */
function createEventsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSS = ss.insertSheet('Events', 0);
  var header = eventsSS.getRange(1, 1, 1, 7).setHorizontalAlignment('center');

  eventsSS.deleteColumns(8, 19);
  eventsSS.deleteRows(2, 999);

  // Setting Values
  eventsSS.setColumnWidths(2, 2, 150);
  eventsSS.setColumnWidths(4, 4, 200);
  header.getBandings().forEach(banding => banding.remove());
  header.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN);
  header.setValues([['ID', 'Datetime', 'User', 'EventType', 'Player', 'Var1', 'Var2']]);
}

/** Function to create the  Characters Sheet */
function createCharactersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var charactersSS = ss.insertSheet('Characters', 0);
  var header = charactersSS.getRange(1, 1, 1, 9).setHorizontalAlignment('center');
  var characters = charactersSS.getRange(2, 1, 5, 9).setHorizontalAlignment('center');
  var killsFormula = charactersSS.getRange(2, 8, 5);
  var downsFormula = charactersSS.getRange(2, 9, 5);

  charactersSS.deleteColumns(10, 17);
  charactersSS.deleteRows(7, 994);

  ss.setNamedRange('characters', charactersSS.getRange(2, 1, 5, 7));
  charactersSS.setColumnWidth(1, 400);
  charactersSS.setColumnWidth(2, 250);
  charactersSS.setColumnWidths(4, 6, 200);
  charactersSS.setRowHeight(1, 30);
  charactersSS.setRowHeights(2, 5, 50);
  charactersSS.getRange(1, 1, 6).setHorizontalAlignment('left');
  charactersSS.getRange(1, 1, 6, 9).getBandings().forEach(banding => banding.remove());
  charactersSS.getRange(1, 1, 6, 9).applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
  header.setFontWeight('bold').setFontSize(15);
  header.setValues([['Name', 'Class', 'Level', 'XP', 'XP to Level Up', 'Max HP', 'Current HP', 'Kills', 'Downs']]);
  characters.setFontSize(25);

  // SET CONDITIONAL FORMAT RULES
  charactersSS.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#CC0000')
    .setRanges([charactersSS.getRange(2, 7, 5)])
    .build(),
    SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2 / 4 > $G2')
    .setBackground('#E06666')
    .setRanges([charactersSS.getRange(2, 7, 5)])
    .build(),
    SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2 / 2 > $G2')
    .setBackground('#F6B26B')
    .setRanges([charactersSS.getRange(2, 7, 5)])
    .build(),
    SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2 > $G2')
    .setBackground('#FFE599')
    .setRanges([charactersSS.getRange(2, 7, 5)])
    .build()
  ]);

  // SET FORMULAS
  killsFormula.setFormula('=IFNA(SUM(QUERY(Events!$A:$G, "SELECT F WHERE E = \'"&$A2&"\' AND D = \'KILL_REPORT\'", 0)), 0)');
  downsFormula.setFormula('=IFNA(QUERY(Events!$A:$G, "SELECT COUNT(E) WHERE E = \'"&$A2&"\' AND D = \'PLAYER_DOWN\' LABEL COUNT(E) \'\'", 0), 0)');
}

/** DEFAULT ADD EVENT FUNCTION 
 * @param {String} user User running event
 * @param {String} eventType Type of event
 * @param {String} player Player
 * @param {Int} var1 Free integer field
 * @param {String} var2 Free String field
*/
function addEvent(user, eventType, player, var1, var2) {
  var eventsSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  var id = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings').getRange('eventID').getValue()+1;

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings').getRange('eventID').setValue(id);
  eventsSS.getRange(eventsSS.getLastRow() + 1, 1, 1, 7).setValues([
    [
      id,
      new Date(),
      user,
      eventType,
      player,
      var1,
      var2
    ]
  ]);
}

/** Function to refresh data from DND Beyond into the Character sheet */
function refreshCharacterData() {
  var players = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('players').getValues();
  var characters = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('characters');
  var characterData = [];
  var xpLevels = [{level: 1, xp: 0},{level: 2, xp: 300},{level: 3, xp: 900},{level: 4, xp: 2700},{level: 5, xp: 6500},{level: 6, xp: 14000},{level: 7, xp: 23000},{level: 8, xp: 34000},{level: 9, xp: 48000},{level: 10, xp: 64000},{level: 11, xp: 85000},{level: 12, xp: 100000},{level: 13, xp: 120000},{level: 14, xp: 140000},{level: 15, xp: 165000},{level: 16, xp: 195000},{level: 17, xp: 225000},{level: 18, xp: 265000},{level: 19, xp: 305000},{level: 20, xp: 355000}];

  // FILTER DOWN CHARACTERS IF BLANK
  players = players.filter(player => player[1] !== '' && player[0] !== 'Player');
  if(players.length === 0) return;
  
  // GET CHARACTER DATA FROM DND BEYOND
  players.forEach((player) => {
    var data = pullCharacterJSON(player[1]);

    if(data.error) {
      console.log('[ERROR] cannot contact DND Beyond Character service for: ' + player[0] + ' ID:' + player[1])
      characterData.push([player[0],'','','','','', '']);
      addEvent('SYSTEM', 'ERROR', '', '', data.error);
      return;
    }

    var levelUpXp = xpLevels.filter((level) => data.classes[0].level + 1 === level.level);
    var constMods = [];
    var constMod = data.overrideStats[2].value !== null ? data.overrideStats[2].value : data.stats[2].value + data.bonusStats[2].value;
    var healthPoints = data.baseHitPoints + data.temporaryHitPoints;
    var ssHealthPoints = characters.getCell(characterData.length + 1, 7).getValue();
    
    // GETTING ALL CONSTITUTION BONUSES FROM DATA.MODIFIERS
    constMods = constMods.concat(data.modifiers.race.filter((modifier) => modifier.subType === "constitution-score"));
    constMods = constMods.concat(data.modifiers.class.filter((modifier) => modifier.subType === "constitution-score"));
    constMods = constMods.concat(data.modifiers.background.filter((modifier) => modifier.subType === "constitution-score"));
    if(constMods.length > 0) constMods.forEach((mod) => constMod += mod.value);
    constMod = Math.floor((constMod - 10) / 2)
    healthPoints += constMod * data.classes[0].level;

    // AUTO CHARACTER DOWN EVENT
    if(ssHealthPoints !== 0 && ssHealthPoints === healthPoints && healthPoints === data.removedHitPoints) addEvent('SYSTEM', 'PLAYER_DOWN', player[0]);
    
    characterData.push([
      player[0],
      data.classes[0].definition.name,
      data.classes[0].level,
      data.currentXp,
      levelUpXp[0].xp - data.currentXp,
      healthPoints,
      healthPoints - data.removedHitPoints
    ]);
  });

  characters.clearContent();
  characters.setValues(characterData);
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName('charLastRefresh').setValue(new Date());
  addEvent('SYSTEM', 'CHAR_REFRESH');
}

/** Function to pull character JSON data from DND Beyond Character Service API
 * @param {String} id Player ID to pull JSON data
 * 
*/
function pullCharacterJSON(id) {
  try{
    var link = 'https://character-service.dndbeyond.com/character/v5/character/';
    var results = JSON.parse(UrlFetchApp.fetch(link + id, {'method' : 'get'}).getContentText("UTF-8"));
    return results.data;
  } catch(error) {
    return {error}
  }
}

/** Function to create an event tracking player kills */
function killTracker() {
  var player = SpreadsheetApp.getActiveSpreadsheet().getRange('killTracker_input1').getValue();
  var kills = SpreadsheetApp.getActiveSpreadsheet().getRange('killTracker_input2').getValue();
  if(player != '' && kills > 0) {
    addEvent('SYSTEM', 'KILL_REPORT', player, kills);
    SpreadsheetApp.getActiveSpreadsheet().getRange('killTracker_input1').clearContent();
    SpreadsheetApp.getActiveSpreadsheet().getRange('killTracker_input2').clearContent();
  }
  else if(player != '' && kills == '') {
    addEvent('SYSTEM', 'KILL_REPORT', player, 1);
    SpreadsheetApp.getActiveSpreadsheet().getRange('killTracker_input1').clearContent();
    SpreadsheetApp.getActiveSpreadsheet().getRange('killTracker_input2').clearContent();
  }
}

/** Function to create an event tracking player NAT 20 rolls */
function natTracker() {
  var player = SpreadsheetApp.getActiveSpreadsheet().getRange('nat20_input1').getValue();
  if(player != '') {
    addEvent('SYSTEM', 'NAT20_ROLL', player);
    SpreadsheetApp.getActiveSpreadsheet().getRange('nat20_input1').clearContent();
  }
}

/** Function to create an event tracking player CRIT FAIL rolls */
function critTracker() {
  var player = SpreadsheetApp.getActiveSpreadsheet().getRange('critFail_input1').getValue();
  if(player != '') {
    addEvent('SYSTEM', 'CRITFAIL_ROLL', player);
    SpreadsheetApp.getActiveSpreadsheet().getRange('critFail_input1').clearContent();
  }
}

function godRollTracker() {
  var cat = SpreadsheetApp.getActiveSpreadsheet().getRange('godRoll_input1').getValue();
    if(cat != '') {
    addEvent('SYSTEM', 'GODROLL', cat);
    SpreadsheetApp.getActiveSpreadsheet().getRange('godRoll_input1').clearContent();
  }
}
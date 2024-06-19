/**
 * Wuwa DPS Calculator Script
 * by @Maygi
 *
 * This is the script attached to the Wuwa DPS Calculator. Running this is required to update
 * all the calculations. Adjust the CHECK_STATS flag if you'd like it to run the Substat checking
 * logic (to calculate which substats are the most effective). This is toggleable because this
 * causes a considerable runtime increase.
 *
 * V1.3.1 - Havoc MC & Danjin initial release.
 */
var CHECK_STATS = false;

var STANDARD_BUFF_TYPES = ['Normal', 'Heavy', 'Skill', 'Liberation'];
var ELEMENTAL_BUFF_TYPES = ['Glacio', 'Fusion', 'Electro', 'Aero', 'Spectro', 'Havoc'];
var WEAPON_MULTIPLIERS = new Map([
  [1, [1, 1]],
  [20, [3.27, 1.78]],
  [40, [5.62, 2.56]],
  [50, [7.13, 2.94]],
  [60, [8.64, 3.33]],
  [70, [10.15, 3.72]],
  [80, [11.66, 4.11]],
  [90, [12.5, 4.5]]
]);
var CHAR_CONSTANTS = getCharacterConstants();
var skillData = [];
var passiveDamageInstances = [];
var weaponData = {};
var charData = {};
var bonusStats = [];
var queuedBuffs = [];
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var rotationSheet = sheet.getSheetByName('Calculator');
var skillLevelMultiplier = rotationSheet.getRange('AH5').getValue();

/**
 * The main method that runs all the calculations and updates the data.
 * Yes, I know, it's like an 800 line method, so ugly.
 */
function runCalculations() {
  var character1 = rotationSheet.getRange('B7').getValue();
  var character2 = rotationSheet.getRange('B8').getValue();
  var character3 = rotationSheet.getRange('B9').getValue();
  var levelCap = rotationSheet.getRange('F4').getValue();
  var oldDamage = rotationSheet.getRange('G19').getValue();
  var characters = [character1, character2, character3];
  var activeBuffs = {};
  activeBuffs['Team'] = new Set();
  activeBuffs[character1] = new Set();
  activeBuffs[character2] = new Set();
  activeBuffs[character3] = new Set();
  rotationSheet.getRange('H15').setValue(oldDamage);

  charData = {};
  charData[character1] = rowToCharacterInfo(sheet.getRange('A7:Z7').getValues()[0], levelCap);
  charData[character2] = rowToCharacterInfo(sheet.getRange('A8:Z8').getValues()[0], levelCap);
  charData[character3] = rowToCharacterInfo(sheet.getRange('A9:Z9').getValues()[0], levelCap);
  console.log(charData[character1]);

  bonusStats = getBonusStats(character1, character2, character3);

  /**
   * Data for stat analysis.
   */
  let statCheckMap = new Map([
    ['Attack', .09],
    ['Health', .09],
    ['Defense', .1139],
    ['Crit', .084],
    ['Crit Dmg', .168],
    ['Normal', .09],
    ['Heavy', .09],
    ['Skill', .09],
    ['Liberation', .09],
    ['Flat Attack', 50]
  ]);
  var charStatGains = {};
  var charEntries = {};
  for (var j = 0; j < characters.length; j++) {
    charEntries[characters[j]] = 0;
    charStatGains[characters[j]] = {
      'Attack': 0,
      'Health': 0,
      'Defense': 0,
      'Crit': 0,
      'Crit Dmg': 0,
      'Normal': 0,
      'Heavy': 0,
      'Skill': 0,
      'Liberation': 0,
      'Flat Attack': 0
    };
  }
  console.log(charStatGains);

  weaponData = {};
  var weapons = getWeapons();

  var charactersWeaponsRange = rotationSheet.getRange("D7:D9").getValues();
  var weaponRankRange = rotationSheet.getRange("F7:F9").getValues();

  //load echo data into the echo parameter
  var echoes = getEchoes();

  for (var i = 0; i < characters.length; i++) {
    weaponData[characters[i]] = characterWeapon(weapons[charactersWeaponsRange[i][0]], levelCap, weaponRankRange[i][0]);
    charData[characters[i]].echo = echoes[charData[characters[i]].echo];
  }

  skillData = [];
  var effectObjects = getSkills();
  effectObjects.forEach(function(effect) {
    skillData[effect.name] = effect;
  });


  //console.log(activeBuffs);
  var trackedBuffs = []; // Stores the active buffs for each time point.
  var dataCellCol = 'F';
  var dataCellColTeam = 'G';
  var dataCellColDmg = 'H';
  var dataCellColResults = 'D';
  var dataCellColNextSub = 'I';
  var dataCellRowResults = 12;
  var dataCellRowNextSub = 66;

  /**
   * Outro buffs are special, and are saved to be applied to the NEXT character swapped into.
   */
  var queuedBuffsForNext = [];
  var lastCharacter = null;

  var swapped = false;
  var allBuffs = getActiveEffects(); // retrieves all buffs "in play" from the ActiveEffects table.
  allBuffs.sort((a, b) => {
    if (a.type === "Dmg" && b.type !== "Dmg") {
      return -1; // A is a "Dmg" type and B is not, A should come first
    } else if (a.type !== "Dmg" && b.type === "Dmg") {
      return 1; // B is a "Dmg" type and A is not, B should come first
    } else {
      return 0; // Both are "Dmg" types, or neither are, so retain their relative positions
    }
  });
  var weaponBuffsRange = sheet.getSheetByName('WeaponBuffs').getRange("A2:K500").getValues().filter(function(row) {
    return row[0].toString().trim() !== ''; // Ensure that the name is not empty
  });
  var weaponBuffData = weaponBuffsRange.map(rowToWeaponBuffRawInfo);

  var echoBuffsRange = sheet.getSheetByName('EchoBuffs').getRange("A2:K500").getValues().filter(function(row) {
    return row[0].toString().trim() !== ''; // Ensure that the name is not empty
  });
  var echoBuffData = echoBuffsRange.map(rowToEchoBuffInfo);

  for (var i = 0; i < 3; i++) { //loop through characters and add buff data if applicable
    echoBuffData.forEach(echoBuff => {
      if (echoBuff.name.includes(charData[characters[i]].echo.name) || echoBuff.name.includes(charData[characters[i]].echo.echoSet)) {
        var newBuff = createEchoBuff(echoBuff, characters[i]);
        allBuffs.push(newBuff);
        console.log("adding echo buff " + echoBuff.name + " to " + characters[i]);
        console.log(newBuff);
      }
    });
    weaponBuffData.forEach(weaponBuff => {
      if (weaponBuff.name.includes(weaponData[characters[i]].weapon.buff)) {
        var newBuff = rowToWeaponBuff(weaponBuff, weaponData[characters[i]].rank, characters[i]);
        allBuffs.push(newBuff);
      }
    });
  }




  //apply passive buffs
  for (let i = allBuffs.length - 1; i >= 0; i--) {
    let buff = allBuffs[i];
    if (buff.triggeredBy === 'Passive' || buff.duration === 'Passive') {
      buff.duration = 9999;
      console.log("buff " + buff.name + " applies to: " + buff.appliesTo);
      activeBuffs[buff.appliesTo].add(createActiveBuff(buff, 0));
      console.log("adding passive buff : " + buff.name + " to " + buff.appliesTo);

      allBuffs.splice(i, 1); // remove passive buffs from the list afterwards
    }
  }

  console.log("ALL BUFFS:");
  console.log(allBuffs);
  //console.log(weaponBuffsRange[0]);

  let totalDamageMap = new Map([
    ['Normal', 0],
    ['Heavy', 0],
    ['Skill', 0],
    ['Liberation', 0],
    ['Intro', 0],
    ['Outro', 0],
    ['Echo', 0]
  ]);

  for (var i = 21; i <= 60; i++) { //clear the content
    var range = rotationSheet.getRange('F' + i + ':AE' + i);

    range.setValue('');
    range.setFontWeight('normal');
    range.clearNote();
  }
  var statWarningRange = rotationSheet.getRange('I69');
  if (CHECK_STATS) {
    statWarningRange.setValue('The above values are accurate for the latest simulation!')
  } else {
    statWarningRange.setValue('The above values are from a PREVIOUS SIMULATION.')
  }
  var freezeTime = 0;
  for (var i = 21; i <= 60; i++) {
    swapped = false;
    var passiveDamageQueued = null;
    var activeCharacter = rotationSheet.getRange('A' + i).getValue();
    if (lastCharacter != null && activeCharacter != lastCharacter) { //a swap was performed
      swapped = true;
    }
    var currentTime = rotationSheet.getRange('C' + (i + 1)).getValue(); // current time starts from the row below, at the end of this skill cast
    var currentSkill = rotationSheet.getRange('B' + i).getValue(); // the current skill

    if (currentSkill.length == 0) {
      break;
    }
    var classification = rotationSheet.getRange('E' + i).getValue();
    var skillRef = getSkillReference(skillData, currentSkill, activeCharacter);
    if (skillRef.name.includes("Liberation: "))
      freezeTime += skillRef.castTime;
    /*console.log("Active Character: " + activeCharacter + "; Current buffs: " + activeBuffs[activeCharacter] +"; filtering for expired");*/

    var activeBuffsArray = [...activeBuffs[activeCharacter]];
    activeBuffsArray = activeBuffsArray.filter(activeBuff => {
      var endTime = ((activeBuff.buff.type === "StackingBuff")
        ? activeBuff.stackTime
        : activeBuff.startTime) + activeBuff.buff.duration;
      //console.log(activeBuff.buff.name + " end time: " + endTime +"; current time = " + currentTime);
      if (activeBuff.buff.type === "BuffUntilSwap" && swapped) {
        console.log("BuffUntilSwap buff " + activeBuff.name + " was removed");
        return true;
      }
      return currentTime <= endTime; // Keep the buff if the current time is less than or equal to the end time
    });
    activeBuffs[activeCharacter] = new Set(activeBuffsArray); // Convert the array back into a Set
    if (swapped && queuedBuffsForNext.length > 0) { //add outro skills after the buffuntilswap check is performed
      queuedBuffsForNext.forEach(queuedBuff => {
        var outroCopy = Object.assign({}, queuedBuff);
        outroCopy.buff.appliesTo = outroCopy.buff.appliesTo === 'Next' ? activeCharacter : outroCopy.buff.appliesTo;
        if (queuedBuff.buff.appliesTo === 'Team')
          activeBuffs['Team'].add(outroCopy);
        else
          activeBuffs[activeCharacter].add(outroCopy);
        console.log("Added queuedForNext buff [" + queuedBuff.buff.name + "] from " + lastCharacter + " to " + activeCharacter);
        console.log(outroCopy);
      });
      queuedBuffsForNext = [];
    }
    lastCharacter = activeCharacter;
    if (queuedBuffs.length > 0) { //add queued buffs procced from passive effects
      queuedBuffs.forEach(queuedBuff => {
        var found = false;
        var copy = Object.assign({}, queuedBuff);
        copy.buff.appliesTo = copy.buff.appliesTo === 'Next' ? activeCharacter : copy.buff.appliesTo;
        var activeSet = copy.buff.appliesTo === 'Team' ? activeBuffs['Team'] : activeBuffs[activeCharacter];

        activeSet.forEach(activeBuff => { //loop through and look for if the buff already exists
          if (activeBuff.buff.name == copy.buff.name) {
            found = true;
            if (currentTime - activeBuff.stackTime >= activeBuff.buff.stackInterval) {
              console.log(`updating stacks: current: ${activeBuff.stacks}; new stacks: ${copy.stacks}; limit: ${activeBuff.buff.stackLimit}`);
              activeBuff.stacks = Math.min(activeBuff.stacks + copy.stacks, activeBuff.buff.stackLimit);
              activeBuff.stackTime = currentTime; // this actually is not accurate, will fix later. should move forward on multihits
              //console.log("updating stacking buff: " + buff.name);
            }
          }
        });
        if (!found) { //add a new buff
          activeSet.add(copy);
        }
        console.log("Added queued buff [" + queuedBuff.buff.name + "] to " + copy.buff.appliesTo);
      });
      queuedBuffs = [];
    }

    var activeBuffsArrayTeam = [...activeBuffs['Team']];
    activeBuffsArrayTeam = activeBuffsArrayTeam.filter(activeBuff => {
      var endTime = ((activeBuff.buff.type === "StackingBuff")
        ? activeBuff.stackTime
        : activeBuff.startTime) + activeBuff.buff.duration;
      //console.log("current teambuff end time: " + endTime +"; current time = " + currentTime);
      return currentTime <= endTime; // Keep the buff if the current time is less than or equal to the end time
    });
    activeBuffs['Team'] = new Set(activeBuffsArrayTeam); // Convert the array back into a Set

    // check for new buffs triggered at this time and add them to the active list
    allBuffs.forEach(buff => {
      //console.log(buff);
      var activeSet = buff.appliesTo === 'Team' ? activeBuffs['Team'] : activeBuffs[activeCharacter];
      var triggeredBy = buff.triggeredBy;
      if (triggeredBy.includes(';')) { //for cases that have additional conditions, remove them for the initial check
        triggeredBy = triggeredBy.split(';')[0];
      }
      var introOutro = buff.name.includes("Outro") || buff.name.includes("Intro");
      if (triggeredBy.length == 0 && introOutro) {
        triggeredBy = buff.name;
      }
      if (triggeredBy === 'Any')
        triggeredBy = skillRef.name;
      var triggeredByConditions = triggeredBy.split(',');
      //console.log("checking conditions for " + buff.name +"; applies to: " + buff.appliesTo + "; conditions: " + triggeredByConditions)

      // check if any of the conditions in triggeredByConditions match
      var isActivated = triggeredByConditions.some(function(condition) {
        condition = condition.trim();
        var conditionIsSkillName = condition.length > 2;

        if (conditionIsSkillName) {
          //console.log(`passive damage queued: ${passiveDamageQueued != null}, condition: ${condition}, name: ${passiveDamageQueued != null ? passiveDamageQueued.name : "none"}, buff.canActivate: ${buff.canActivate}, owner: ${passiveDamageQueued != null ? passiveDamageQueued.owner : "none"}`);
          if (passiveDamageQueued != null && condition.includes(passiveDamageQueued.name) && (buff.canActivate === passiveDamageQueued.owner || buff.canActivate === 'Team')) {
            console.log("passive damage queued exists - adding new buff " + buff.name);
            passiveDamageQueued.addBuff(buff.type === 'StackingBuff' ? createActiveStackingBuff(buff, currentTime, 1) : createActiveBuff(buff, currentTime));
          }
          // the condition is a skill name, check if it's included in the currentSkill
          var applicationCheck = buff.appliesTo === activeCharacter || buff.appliesTo === 'Team' || introOutro || buff.appliesTo === 'Next' || skillRef.source === activeCharacter;
          if (condition === 'Swap' && !skillRef.name.includes('Intro') && (skillRef.castTime == 0 || skillRef.name.includes('(Swap)'))) { //this is a swap-out skill
            console.log(`application check: ${applicationCheck}, buff.canActivate: ${buff.canActivate}, skillRef.source: ${skillRef.source}, buff.canActivate: ${buff.canActivate}`);
            return applicationCheck && ((buff.canActivate === activeCharacter || buff.canActivate === 'Team') || (skillRef.source === activeCharacter && introOutro));
          } else {
            return currentSkill.includes(condition) && applicationCheck && (buff.canActivate === activeCharacter || buff.canActivate === 'Team');
          }
        } else {
          //console.log(`passive damage queued: ${passiveDamageQueued != null}, condition: ${condition}, name: ${passiveDamageQueued != null ? passiveDamageQueued.name : "none"}, buff.canActivate: ${buff.canActivate}, owner: ${passiveDamageQueued != null ? passiveDamageQueued.owner : "none"}`);
          if (passiveDamageQueued != null && condition.includes(passiveDamageQueued.name) && (buff.canActivate === passiveDamageQueued.owner || buff.canActivate === 'Team')) {
            console.log("passive damage queued exists - adding new buff " + buff.name);
            passiveDamageQueued.addBuff(buff.type === 'StackingBuff' ? createActiveStackingBuff(buff, currentTime, 1) : createActiveBuff(buff, currentTime));
          }
          // the condition is a classification code, check against the classification
          // and ensure that the buff applies to the active character
          return classification.includes(condition) && (buff.canActivate === activeCharacter);
        }
      });
      if (isActivated) { //activate this effect
        var found = false;
        console.log(buff.name + " has been activated");
        if (buff.type === 'Dmg') { //add a new passive damage instance
          //queue the passive damage and snapshot the buffs later
          console.log("adding a new type of passive damage");
          passiveDamageQueued = new PassiveDamage(buff.name, buff.classifications, buff.amount, buff.duration, currentTime, buff.stackLimit, buff.stackInterval, buff.triggeredBy, activeCharacter, i);
          console.log(passiveDamageQueued);
        } else if (buff.type === 'StackingBuff') {
          var stacksToAdd = 1;
          if (buff.stackInterval < skillRef.castTime) { //potentially add multiple stacks
            let maxStacksByTime = buff.stackInterval == 0 ? skillRef.numberOfHits : Math.floor(skillRef.castTime / buff.stackInterval);
            stacksToAdd = Math.min(maxStacksByTime, skillRef.numberOfHits);
          }
          //console.log("this buff applies to: " + buff.appliesTo + "; active char: " + activeCharacter);
          //console.log(buff.name + " is a stacking buff. attempting to add " + stacksToAdd + " stacks");
          activeSet.forEach(activeBuff => { //loop through and look for if the buff already exists
            if (activeBuff.buff.name == buff.name) {
              found = true;
              if (currentTime - activeBuff.stackTime >= activeBuff.buff.stackInterval) {
                activeBuff.stacks = Math.min(activeBuff.stacks + stacksToAdd, buff.stackLimit);
                activeBuff.stackTime = currentTime; // this actually is not accurate, will fix later. should move forward on multihits
                //console.log("updating stacking buff: " + buff.name);
              }
            }
          });
          if (!found) { //add a new stackable buff
            activeSet.add(createActiveStackingBuff(buff, currentTime, Math.min(stacksToAdd, buff.stackLimit)));
            //console.log("adding new stacking buff: " + buff.name);
          }
        } else {
          if (buff.type === "BuffEnergy" && currentTime >= buff.availableIn) { //add energy instead of adding the buff
            //todo energy implementation
            console.log("adding resonance energy: " + buff.amount);
            buff.availableIn = currentTime + buff.stackInterval;
          } else {
            if (buff.name.includes("Outro") || buff.appliesTo === 'Next') { //outro buffs are special and are saved for the next character
              queuedBuffsForNext.push(createActiveBuff(buff, currentTime));
              console.log("queuing buff: " + buff.name);
            } else {
              activeSet.forEach(activeBuff => { //loop through and look for if the buff already exists
                if (activeBuff.buff.name == buff.name) {
                  activeBuff.startTime = currentTime;
                  found = true;
                }
              });
              if (!found) {
                activeSet.add(createActiveBuff(buff, currentTime));
                //console.log("adding new buff: " + buff.name);
              }
            }
          }
        }
      }
    });

    activeBuffsArray = Array.from(activeBuffs[activeCharacter]);
    let buffNames = activeBuffsArray.map(activeBuff => activeBuff.buff.name + (activeBuff.buff.type == "StackingBuff" ? (" x" + activeBuff.stacks) : "")); // Extract the name from each object
    let buffNamesString = buffNames.join(', ');

    //console.log(buffNamesString);
    /*activeBuffs[activeCharacter].forEach(buff => {
      console.log(buff);
    });*/

    //console.log("Writing to: " + (dataCellCol + i) + "; " + buffNamesString);
    rotationSheet.getRange(dataCellCol + i).setValue("(" + activeBuffsArray.length + ") " + buffNamesString);

    activeBuffsArrayTeam = Array.from(activeBuffs['Team']);
    buffNames = activeBuffsArrayTeam.map(activeBuff => activeBuff.buff.name + (activeBuff.buff.type == "StackingBuff" ? (" x" + activeBuff.stacks) : "")); // extract the name from each object
    buffNamesString = buffNames.join(', ');

    console.log(buffNamesString);
    /*activeBuffs['Team'].forEach(buff => {
      console.log(buff);
    });*/

    //console.log("Writing to: " + (dataCellColTeam + i) + "; " + buffNamesString);
    rotationSheet.getRange(dataCellColTeam + i).setValue("(" + activeBuffsArrayTeam.length + ") " + buffNamesString);


    // utility function to set/update total buff amount
    function updateTotalBuffMap(buffCategory, buffType, buffAmount) {
      console.log("updating buff for " + buffCategory + ", which has type " + buffType);
      let key = buffCategory;
      if (buffCategory === 'All') {
        STANDARD_BUFF_TYPES.forEach(buff => {
          let newKey = translateClassificationCode(buff);
          newKey = buffType === 'Deepen' ? `${newKey} (${buffType})` : `${newKey}`;
          let currentAmount = totalBuffMap.get(newKey);
          if (!totalBuffMap.has(newKey)) {
            return;
          }
          totalBuffMap.set(newKey, currentAmount + buffAmount); // Update the total amount
          //console.log("updating buff " + newKey + " to " + (currentAmount) + " (+" + buffAmount + ")");
        });
      } else if (buffCategory === 'AllEle') {
        ELEMENTAL_BUFF_TYPES.forEach(buff => {
          let newKey = translateClassificationCode(buff);
          let currentAmount = totalBuffMap.get(newKey);
          if (!totalBuffMap.has(newKey)) {
            return;
          }
          totalBuffMap.set(newKey, currentAmount + buffAmount); // Update the total amount
          //console.log("updating buff " + newKey + " to " + (currentAmount) + " (+" + buffAmount + ")");
        });
      } else {
        var categories = buffCategory.split(',');
        categories.forEach(category => {
          let newKey = translateClassificationCode(category);
          newKey = buffType === 'Deepen' ? `${newKey} (${buffType})` : `${newKey}`;
          let currentAmount = totalBuffMap.get(newKey);
          let buffKey = buffType === 'Bonus' ? 'Specific' : (buffType === 'Deepen' ? 'Deepen' : 'Multiplier');
          if (buffKey === 'Deepen') { //apply element-specific deepen effects IF MATCH
            if (skillRef.classifications.includes(category)) {
              newKey = 'Deepen';
              currentAmount = totalBuffMap.get(newKey);
              console.log("updating amplify to " + (currentAmount) + " (+" + buffAmount + ")");
              totalBuffMap.set(newKey, currentAmount + buffAmount); // Update the total amount
            }
          } else {
            if (!totalBuffMap.has(newKey)) { //skill-specific buff
              if (skillRef.name.includes(newKey)) {
                console.log("adding new key as " + buffKey + ": " + newKey);
                let currentBonus = totalBuffMap.get(buffKey);
                totalBuffMap.set(buffKey, currentBonus + buffAmount); // Update the total amount
              } else { //add the skill key as a new value for potential procs
                totalBuffMap.set(`${newKey} (${buffKey})`, buffAmount);
                console.log("no match, but adding key " + newKey)
              }
            } else {
              totalBuffMap.set(newKey, currentAmount + buffAmount); // Update the total amount
            }
          }
          //console.log("updating buff " + key + " to " + (currentAmount) + " (+" + buffAmount + ")");
        });
      }
    }

    // Process buff array
    function processBuffs(buffs) {
      buffs.forEach(buffWrapper => {
        let buff = buffWrapper.buff;
        // special buff types are handled slightly differently
        let specialBuffTypes = ['Attack', 'Health', 'Defense', 'Crit', 'Crit Dmg'];
        if (specialBuffTypes.includes(buff.buffType)) {
          updateTotalBuffMap(buff.buffType, '', buff.amount * (buff.type === 'StackingBuff' ? buffWrapper.stacks : 1));
        } else { // for other buffs, just use classifications as is
          updateTotalBuffMap(buff.classifications, buff.buffType, buff.amount * (buff.type === 'StackingBuff' ? buffWrapper.stacks : 1));
        }
      });
    }

    function writeBuffsToSheet(i) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calculator');
      var keysIterator = totalBuffMap.keys();

      var bufferRange = sheet.getRange("I" + i + ":AE" + i);
      var values = [];

      for (let key of keysIterator) {
        values.push(totalBuffMap.get(key));
        //console.log(key + " value : " + totalBuffMap.get(key))
      }

      if (values.length > 23) {
        values = values.slice(0, 23);
      }
      bufferRange.setValues([values])
    }

    let totalBuffMap = new Map([
      ['Attack', 0],
      ['Health', 0],
      ['Defense', 0],
      ['Crit', 0],
      ['Crit Dmg', 0],
      ['Normal', 0],
      ['Heavy', 0],
      ['Skill', 0],
      ['Liberation', 0],
      ['Normal (Deepen)', 0],
      ['Heavy (Deepen)', 0],
      ['Skill (Deepen)', 0],
      ['Liberation (Deepen)', 0],
      ['Physical', 0],
      ['Glacio', 0],
      ['Fusion', 0],
      ['Electro', 0],
      ['Aero', 0],
      ['Spectro', 0],
      ['Havoc', 0],
      ['Specific', 0],
      ['Deepen', 0],
      ['Multiplier', 0],
      ['Flat Attack', 0]
    ]);
    let teamBuffMap = Object.assign({}, totalBuffMap);

    if (totalBuffMap.has(weaponData[activeCharacter].mainStat)) {
      let currentAmount = totalBuffMap.get(weaponData[activeCharacter].mainStat);
      totalBuffMap.set(weaponData[activeCharacter].mainStat, currentAmount + weaponData[activeCharacter].mainStatAmount);
      console.log("adding mainstat " + weaponData[activeCharacter].mainStat + " (+" + weaponData[activeCharacter].mainStatAmount + ") to " + activeCharacter);
    }
    charData[activeCharacter].bonusStats.forEach((statPair) => {
      let stat = statPair[0];
      let value = statPair[1];
      let currentAmount = totalBuffMap.get(stat) || 0;
      totalBuffMap.set(stat, currentAmount + value);
    });
    processBuffs(activeBuffsArray);

    processBuffs(activeBuffsArrayTeam);
    if (passiveDamageQueued != null) { //snapshot passive damage BEFORE team buffs are applied
      //TEMP: move this above activeBuffsArrayTeam and implement separate buff tracking
      passiveDamageQueued.setTotalBuffMap(totalBuffMap);
      passiveDamageInstances.push(passiveDamageQueued);
    }

    rotationSheet.getRange(dataCellColTeam + i).setValue("(" + activeBuffsArrayTeam.length + ") " + buffNamesString);
    writeBuffsToSheet(i);
    if (skillRef.type.includes("Buff")) {
      rotationSheet.getRange(dataCellColDmg + i).setValue(0);
      continue;
    }

    //damage calculations
    //console.log(charData[activeCharacter]);
    //console.log("bonus input stats:");
    //console.log(bonusStats[activeCharacter]);
    //passiveDamageInstances = passiveDamageInstances.filter(passiveDamage => !passiveDamage.canRemove(currentTime));

    passiveDamageInstances.forEach(passiveDamage => {
      console.log("checking proc conditions for " + passiveDamage.name + "; " + passiveDamage.canProc(currentTime));
      if (passiveDamage.canProc(currentTime) && passiveDamage.checkProcConditions(skillRef)) {
        let procs = passiveDamage.handleProcs(currentTime, skillRef.castTime, skillRef.numberOfHits);
        var damageProc = passiveDamage.calculateProc(activeCharacter) * procs;
        var cell = rotationSheet.getRange(dataCellColDmg + passiveDamage.slot);
        var currentDamage = cell.getValue();
        cell.setValue(currentDamage + damageProc);

        var cellInfo = rotationSheet.getRange('H' + passiveDamage.slot);
        cellInfo.setFontWeight('bold');
        cellInfo.setNote(passiveDamage.getNote());
      }
    });
    var damage = skillRef.damage;
    var attack = (charData[activeCharacter].attack + weaponData[activeCharacter].attack) * (1 + totalBuffMap.get('Attack') + bonusStats[activeCharacter].attack);
    var health = (charData[activeCharacter].health + weaponData[activeCharacter].health) * (1 + totalBuffMap.get('health') + bonusStats[activeCharacter].health);
    var defense = (charData[activeCharacter].defense + weaponData[activeCharacter].defense) * (1 + totalBuffMap.get('Defense') + bonusStats[activeCharacter].defense);
    var critMultiplier = (1 - Math.min(1,(charData[activeCharacter].crit + totalBuffMap.get('Crit')))) * 1 + Math.min(1,(charData[activeCharacter].crit + totalBuffMap.get('Crit'))) * (charData[activeCharacter].critDmg + totalBuffMap.get('Crit Dmg'));
    var damageMultiplier = getDamageMultiplier(skillRef.classifications, totalBuffMap);
    var totalDamage = damage * attack * critMultiplier * damageMultiplier * (weaponData[activeCharacter].weapon.name === 'Nullify Damage' ? 0 : 1) * 0.48 * skillLevelMultiplier;
    console.log(`skill damage: ${damage.toFixed(2)}; attack: ${(charData[activeCharacter].attack + weaponData[activeCharacter].attack).toFixed(2)} x ${(1 + totalBuffMap.get('Attack') + bonusStats[activeCharacter].attack).toFixed(2)}; crit mult: ${critMultiplier.toFixed(2)}; dmg mult: ${damageMultiplier.toFixed(2)}; total dmg: ${totalDamage.toFixed(2)}`);
    rotationSheet.getRange(dataCellColDmg + i).setValue(totalDamage);
    charEntries[activeCharacter]++;

    if (CHECK_STATS) {
      statCheckMap.forEach((value, stat) => {
        if (totalDamage > 0) {
          //console.log("current stat:" + stat + " (" + value + ")");
          let currentAmount = totalBuffMap.get(stat);
          totalBuffMap.set(stat, currentAmount + value);
          var attack = (charData[activeCharacter].attack + weaponData[activeCharacter].attack) * (1 + totalBuffMap.get('Attack') + bonusStats[activeCharacter].attack) + totalBuffMap.get('Flat Attack');
          var health = (charData[activeCharacter].health + weaponData[activeCharacter].health) * (1 + totalBuffMap.get('health') + bonusStats[activeCharacter].health);
          var defense = (charData[activeCharacter].defense + weaponData[activeCharacter].defense) * (1 + totalBuffMap.get('Defense') + bonusStats[activeCharacter].defense);
          var critMultiplier = (1 - (charData[activeCharacter].crit + totalBuffMap.get('Crit'))) * 1 + (charData[activeCharacter].crit + totalBuffMap.get('Crit')) * (charData[activeCharacter].critDmg + totalBuffMap.get('Crit Dmg'));
          var damageMultiplier = getDamageMultiplier(skillRef.classifications, totalBuffMap);
          var newTotalDamage = damage * attack * critMultiplier * damageMultiplier * 0.48 * skillLevelMultiplier;
          //console.log("new total dmg: " + newTotalDamage + " vs old: " + totalDamage + "; gain = " + (newTotalDamage / totalDamage - 1));
          charStatGains[activeCharacter][stat] += newTotalDamage - totalDamage;
          totalBuffMap.set(stat, currentAmount); //unset the value after
        }
      });
    }

    //update damage distribution tracking chart
    for (let j = 0; j < skillRef.classifications.length; j += 2) {
      let code = skillRef.classifications.substring(j, j + 2);
      let key = translateClassificationCode(code);
      if (skillRef.name.includes("Intro"))
        key = "Intro";
      if (skillRef.name.includes("Outro"))
        key = "Outro";
      if (totalDamageMap.has(key)) {
        let currentAmount = totalDamageMap.get(key);
        totalDamageMap.set(key, currentAmount + totalDamage); // Update the total amount
      }
    }
  }

  var startRow = dataCellRowNextSub; // Starting at 66
  var startColIndex = SpreadsheetApp.getActiveSpreadsheet().getRange(dataCellColNextSub + "1").getColumn(); // Get the column index for 'I' which is 9

  //console.log(charStatGains[characters[0]]);
  //console.log(charEntries[characters[0]]);

  var finalTime = rotationSheet.getRange('C60').getValue() - freezeTime;
  var finalDamage = rotationSheet.getRange('G19').getValue();
  console.log(`final time: ${rotationSheet.getRange('C60').getValue()} - freezeTime of ${freezeTime}`);
  rotationSheet.getRange('H16').setValue(finalDamage / finalTime);
  if (CHECK_STATS) {
    for (var i = 0; i < characters.length; i++) {
      if (charEntries[characters[i]] > 0) { // Using [characters[i]] to get each character's entry
        var stats = charStatGains[characters[i]];
        var colOffset = 0; // Initialize column offset for each character

        Object.keys(stats).forEach(function(key) {
          stats[key] /= finalDamage; //charEntries[characters[i]];
          // Calculate the range using row and column indices and write data horizontally.
          var cell = rotationSheet.getRange(startRow, startColIndex + colOffset);
          cell.setValue(stats[key]);
          colOffset++; // Move to the next column for the next stat
        });
        console.log(charStatGains[characters[i]]);
        startRow++; // Move to the next row after writing all stats for a character
      }
    }
  }

  var resultIndex = dataCellRowResults;
  console.log(totalDamageMap);
  for (let [key, value] of totalDamageMap) {
    rotationSheet.getRange(dataCellColResults + (resultIndex++)).setValue(value);
  }

  // Output the tracked buffs for each time point (optional)
  trackedBuffs.forEach(entry => {
    Logger.log('Time: ' + entry.time + ', Active Buffs: ' + entry.activeBuffs.join(', '));
  });
}

/**
 * Gets the percentage bonus stats from the stats input.
 */
function getBonusStats(char1, char2, char3) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calculator'); // Replace with your actual sheet name
  var range = sheet.getRange('I11:L13');
  var values = range.getValues();

  // Stats order should correspond to the columns I, J, K, L
  var statsOrder = ['attack', 'health', 'defense', 'energyRecharge'];

  // Character names - must match exactly with names in script
  var characters = [char1, char2, char3];

  var bonusStats = {};

  // Loop through each character row
  for (var i = 0; i < characters.length; i++) {
    var stats = {};
    // Loop through each stat column
    for (var j = 0; j < statsOrder.length; j++) {
      stats[statsOrder[j]] = values[i][j];
    }
    // Assign the stats object to the corresponding character
    bonusStats[characters[i]] = stats;
  }

  return bonusStats;
}


/**
 * Converts a row from the ActiveEffects sheet into an object.
 * @param {Array} row A single row of data from the ActiveEffects sheet.
 * @return {Object} The row data as an object.
 */
function rowToActiveEffectObject(row) {
  var isRegularFormat = row[7] && row[7].toString().trim() !== '';
  var activator = isRegularFormat ? row[10] : row[6];
  if (skillData[row[0]] != null) {
    activator = skillData[row[0]].source;
  }
  //console.log("row: " + row + "; regular: " + isRegularFormat);
  if (isRegularFormat) {
    return {
      name: row[0],       // skill name
      type: row[1],        // The type of buff
      classifications: row[2],       // The classifications this buff applies to, or All if it applies to all.
      buffType: row[3],       // The type of buff - standard, ATK buff, crit buff, elemental buff, etc
      amount: row[4],        // The value of the buff
      duration: row[5],        // How long the buff lasts - a duration is 0 indicates a passive
      active: row[6],     // Should always be TRUE
      triggeredBy: row[7], // The Skill, or Classification type, this buff is triggered by.
      stackLimit: row[8], // The maximum stack limit of this buff.
      stackInterval: row[9], // The minimum stack interval of gaining a new stack of this buff.
      appliesTo: row[10], // The character this buff applies to, or Team in the case of a team buff
      canActivate: activator,
      availableIn: 0 //cooltime tracker for proc-based effects
    };
  } else { //short format for outros and similar
    return {
      name: row[0],
      type: row[1],
      classifications: row[2],
      buffType: row[3],
      amount: row[4],
      duration: row[5],
      // Assuming that for these rows, the 'active' field is not present, thus it should be assumed true
      active: true,
      triggeredBy: "", // No triggeredBy field for this format
      stackLimit: 0, // Assuming 0 as default value if not present
      stackInterval: 0, // Assuming 0 as default value if not present
      appliesTo: row[6],
      canActivate: activator,

      availableIn: 0 //cooltime tracker for proc-based effects
    };
  }
}

/**
 * Turns a row from "ActiveChar" - aka, the skill data -into a skill data object.
 */
function rowToActiveSkillObject(row) {
  if (row[0].includes("Intro") || row[0].includes("Outro")) {
    return {
      name: row[0], // + " (" + row[6] +")",
      type: row[1],
      damage: row[4],
      castTime: row[0].includes("Intro") ? 1.5 : 0,
      dps: 0,
      classifications: row[2],
      numberOfHits: 1,
      source: row[6] //the name of the character this skill belongs to
    }
  } else return {
    name: row[0], // + " (" + row[6] +")",
    type: "",
    damage: row[1],
    castTime: row[2],
    dps: row[3],
    classifications: row[4],
    numberOfHits: row[5],
    source: row[6] //the name of the character this skill belongs to
  }
}

function rowToCharacterInfo(row, levelCap) {
  const bonusTypes = [
    'Normal', 'Heavy', 'Skill', 'Liberation', 'Physical',
    'Glacio', 'Fusion', 'Electro', 'Aero', 'Spectro', 'Havoc'
  ];

  // Map bonus names to their corresponding row values
  const bonusStatsArray = bonusTypes.map((type, index) => {
    return [type, row[13 + index]]; // Row index starts at 13 for bonusNormal and increments for each bonus type
  });
  return {
    name: row[1],
    resonanceChain: row[2],
    weapon: row[3],
    weaponRank: row[5],
    echo: row[6],
    attack: row[8] + CHAR_CONSTANTS[row[1]].baseAttack * WEAPON_MULTIPLIERS.get(levelCap)[0],
    health: row[9] + CHAR_CONSTANTS[row[1]].baseHealth * WEAPON_MULTIPLIERS.get(levelCap)[0],
    defense: row[10] + CHAR_CONSTANTS[row[1]].baseDef * WEAPON_MULTIPLIERS.get(levelCap)[0],
    crit: Math.min(row[11] + 0.05, 1),
    critDmg: row[12] + 1.5,
    bonusStats: bonusStatsArray
  }
};

function rowToCharacterConstants(row) {
  return {
    name: row[0],
    weapon: row[1],
    baseHealth: row[2],
    baseAttack: row[3],
    baseDef: row[4]
  }
}

function rowToWeaponInfo(row) {
  return {
    name: row[0],
    type: row[1],
    baseAttack: row[2],
    baseMainStat: row[3],
    baseMainStatAmount: row[4],
    buff: row[5]
  }
}

/**
 * Turns a row from the "Echo" sheet into an object.
 */
function rowToEchoInfo(row) {
  return {
    name: row[0],
    damage: row[1],
    castTime: row[2],
    echoSet: row[3],
    classifications: row[4],
    numberOfHits: row[5],
    hasBuff: row[6],
    cooldown: row[7]
  }
}

/**
 * Turns a row from the "EchoBuffs" sheet into an object.
 */
function rowToEchoBuffInfo(row) {
  return {
    name: row[0],
    type: row[1],        // The type of buff
    classifications: row[2],       // The classifications this buff applies to, or All if it applies to all.
    buffType: row[3],       // The type of buff - standard, ATK buff, crit buff, elemental buff, etc
    amount: row[4],        // The value of the buff
    duration: row[5],        // How long the buff lasts - a duration is 0 indicates a passive
    triggeredBy: row[6], // The Skill, or Classification type, this buff is triggered by.
    stackLimit: row[7], // The maximum stack limit of this buff.
    stackInterval: row[8], // The minimum stack interval of gaining a new stack of this buff.
    appliesTo: row[9], // The character this buff applies to, or Team in the case of a team buff
    availableIn: 0 //cooltime tracker for proc-based effects
  }
}

/**
 * Creates a new echo buff object out of the given echo.
 */
function createEchoBuff(echoBuff, character) {
  var newAppliesTo = echoBuff.appliesTo === 'Self' ? character : echoBuff.appliesTo;
  return {
    name: echoBuff.name,
    type: echoBuff.type,        // The type of buff
    classifications: echoBuff.classifications,       // The classifications this buff applies to, or All if it applies to all.
    buffType: echoBuff.buffType,       // The type of buff - standard, ATK buff, crit buff, elemental buff, etc
    amount: echoBuff.amount,        // The value of the buff
    duration: echoBuff.duration,        // How long the buff lasts - a duration is 0 indicates a passive
    triggeredBy: echoBuff.triggeredBy, // The Skill, or Classification type, this buff is triggered by.
    stackLimit: echoBuff.stackLimit, // The maximum stack limit of this buff.
    stackInterval: echoBuff.stackInterval, // The minimum stack interval of gaining a new stack of this buff.
    appliesTo: newAppliesTo, // The character this buff applies to, or Team in the case of a team buff
    canActivate: character,
    availableIn: 0 //cooltime tracker for proc-based effects
  }
}


/**
 * Rows of WeaponBuffs raw - these have slash-delimited values in many columns.
 */
function rowToWeaponBuffRawInfo(row) {
  return {
    name: row[0],       // buff  name
    type: row[1],        // the type of buff
    classifications: row[2],       // the classifications this buff applies to, or All if it applies to all.
    buffType: row[3],       // the type of buff - standard, ATK buff, crit buff, deepen, etc
    amount: row[4],        // slash delimited - the value of the buff
    duration: row[5],        // slash delimited - how long the buff lasts - a duration is 0 indicates a passive. for BuffEnergy, this is the Cd between procs
    triggeredBy: row[6], // The Skill, or Classification type, this buff is triggered by.
    stackLimit: row[7], // slash delimited - the maximum stack limit of this buff.
    stackInterval: row[8], // slash delimited - the minimum stack interval of gaining a new stack of this buff.
    appliesTo: row[9], // The character this buff applies to, or Team in the case of a team buff
    availableIn: 0 //cooltime tracker for proc-based effects
  }
}

/**
 * A refined version of a weapon buff specific to a character and their weapon rank.
 */
function rowToWeaponBuff(weaponBuff, rank, character) {
  var newAmount = weaponBuff.amount.includes('/') ? weaponBuff.amount.split('/')[rank] : weaponBuff.amount;
  var newDuration = weaponBuff.duration.includes('/') ? weaponBuff.duration.split('/')[rank] : weaponBuff.duration;
  var newStackLimit = ("" + weaponBuff.stackLimit).includes('/') ? weaponBuff.stackLimit.split('/')[rank] : weaponBuff.stackLimit;
  var newStackInterval = ("" + weaponBuff.stackInterval).includes('/') ? weaponBuff.stackInterval.split('/')[rank] : weaponBuff.stackInterval;
  var newAppliesTo = weaponBuff.appliesTo === 'Self' ? character : weaponBuff.appliesTo;
  return {
    name: weaponBuff.name,       // buff  name
    type: weaponBuff.type,        // the type of buff
    classifications: weaponBuff.classifications,       // the classifications this buff applies to, or All if it applies to all.
    buffType: weaponBuff.buffType,       // the type of buff - standard, ATK buff, crit buff, deepen, etc
    amount: parseFloat(newAmount),        // slash delimited - the value of the buff
    active: true,
    duration: parseFloat(newDuration),        // slash delimited - how long the buff lasts - a duration is 0 indicates a passive
    triggeredBy: weaponBuff.triggeredBy, // The Skill, or Classification type, this buff is triggered by.
    stackLimit: parseFloat(newStackLimit), // slash delimited - the maximum stack limit of this buff.
    stackInterval: parseFloat(newStackInterval), // slash delimited - the minimum stack interval of gaining a new stack of this buff.
    appliesTo: newAppliesTo, // The character this buff applies to, or Team in the case of a team buff
    canActivate: character,
    availableIn: 0 //cooltime tracker for proc-based effects
  }
}

function characterWeapon(pWeapon, pLevelCap, pRank) {
  return {
    weapon: pWeapon,
    attack: pWeapon.baseAttack * WEAPON_MULTIPLIERS.get(pLevelCap)[0],
    mainStat: pWeapon.baseMainStat,
    mainStatAmount: pWeapon.baseMainStatAmount * WEAPON_MULTIPLIERS.get(pLevelCap)[1],
    rank: pRank - 1
  }
}

function createActiveBuff(pBuff, pTime) {
  return {
    buff: pBuff,
    startTime: pTime,
    stacks: 0,
    stackTime: 0
  }
}

function createActiveStackingBuff(pBuff, time, pStacks) {
  return {
    buff: pBuff,
    startTime: time,
    stacks: pStacks,
    stackTime: time
  }
}

/**
 * Creates a passive damage instance that's actively procced by certain attacks.
 */
class PassiveDamage {
  constructor(name, type, damage, duration, startTime, limit, interval, triggeredBy, owner, slot) {
    this.name = name;
    this.type = type;
    this.damage = damage;
    this.duration = duration;
    this.startTime = startTime;
    this.limit = limit;
    this.interval = interval;
    this.triggeredBy = triggeredBy.split(';')[1];
    this.owner = owner;
    this.slot = slot;
    this.lastProc = -999;
    this.numProcs = 0;
    this.procMultiplier = 1;
    this.totalDamage = 0;
    this.totalBuffMap = [];
    this.proccableBuffs = [];
  }

  addBuff(buff) {
    console.log(`adding ${buff.name} as a proccable buff to ${this.name}`);
    this.proccableBuffs.push(buff);
  }

  /**
   * Handles and updates the current proc time according to the skill reference info.
   */
  handleProcs(currentTime, castTime, numberOfHits) {
    let procs = 0;
    let timeBetweenHits = castTime / (numberOfHits > 1 ? numberOfHits - 1 : 1);
    //console.log(`handleProcs called with currentTime: ${currentTime}, castTime: ${castTime}, numberOfHits: ${numberOfHits}`);
    //console.log(`lastProc: ${this.lastProc}, interval: ${this.interval}, timeBetweenHits: ${timeBetweenHits}`);
    if (this.interval > 0) {
      for (let hitIndex = 0; hitIndex < numberOfHits; hitIndex++) {
        let hitTime = currentTime + timeBetweenHits * hitIndex;
        //console.log(`Checking hitIndex ${hitIndex}: hitTime: ${hitTime}, lastProc + interval: ${this.lastProc + this.interval}`);
        if (hitTime - this.lastProc >= this.interval) {
          procs++;
          this.lastProc = hitTime;
          console.log(`Proc occurred at hitTime: ${hitTime}`);
        }
      }
    } else {
      procs = numberOfHits;
    }
    if (this.limit > 0)
      procs = Math.min(procs, this.limit - this.numProcs);
    this.numProcs += procs;
    this.procMultiplier = procs;
    console.log(`Total procs this time: ${procs}`);
    if (procs > 0) {
      this.proccableBuffs.forEach(buff => {
        if (buff.type === 'StackingBuff') {
          var stacksToAdd = 1;
          if (buff.stackInterval < castTime) { //potentially add multiple stacks
            let maxStacksByTime = buff.stackInterval == 0 ? numberOfHits : Math.floor(castTime / buff.stackInterval);
            stacksToAdd = Math.min(maxStacksByTime, numberOfHits);
          }
          buff.stacks = Math.min(activeBuff.stacks + stacksToAdd, buff.stackLimit);
          buff.stackTime = currentTime; // this actually is not accurate, will fix later. should move forward on multihits
        }
        queuedBuffs.push(buff);
      });
    }
    return procs;
  }

  canRemove(currentTime) {
    return this.numProcs >= this.limit && this.limit > 0 || (currentTime - this.startTime > this.duration);
  }

  canProc(currentTime) {
    console.log("can it proc? CT: " + currentTime + "; lastProc: " + this.lastProc + "; interval: " + this.interval);
    return currentTime - this.lastProc >= this.interval - .01;
  }

  /**
   * Sets the total buff map, updating with any skill-specific buffs.
   */
  setTotalBuffMap(totalBuffMap) {
    this.totalBuffMap = totalBuffMap;

    //these may have been set from the skill proccing it
    this.totalBuffMap.set('Specific', 0);
    this.totalBuffMap.set('Deepen', 0);
    this.totalBuffMap.set('Multiplier', 0);

    this.totalBuffMap.forEach((value, stat) => {
      if (stat.includes(this.name)) {
        if (stat.includes('Specific')) {
          let current = totalBuffMap.get('Specific');
          this.totalBuffMap.set('Specific', current + value);
          console.log(`updating damage bonus for ${name} to ${current} + ${value}`);
        } else if (stat.includes('Multiplier')) {
          let current = totalBuffMap.get('Multiplier');
          this.totalBuffMap.set('Multiplier', current + value);
          console.log(`updating damage multiplier for ${this.name} to ${current} + ${value}`);
        } else if (stat.includes('Deepen')) {
          let element = reverseTranslateClassificationCode(stat.split('(')[0].trim());
          if (this.type.includes(element)) {
            let current = totalBuffMap.get('Deepen');
            this.totalBuffMap.set('Deepen', current + value);
            console.log(`updating damage Deepen for ${this.name} to ${current} + ${value}`);
          }
        }
      }
    });

  }

  checkProcConditions(skillRef) {
    console.log("checking proc conditions with skill: " + this.triggeredBy + " vs " + skillRef.name);
    console.log(skillRef);
    if (!this.triggeredBy)
      return false;
    if (this.triggeredBy === 'Any' || skillRef.name.includes(this.triggeredBy))
      return true;
    var triggeredByConditions = this.triggeredBy.split(',');
    triggeredByConditions.forEach(condition => {
      if (skillRef.classifications.includes(condition))
        return true;
    });
    //console.log("failed match");
    return false;
  }

  /**
   * Calculates a proc's damage, and adds it to the total.
   */
  calculateProc(activeCharacter) {
    var bonusAttack = 0;
    /*if (activeCharacter != this.owner) {
      if (charData[this.owner].weapon.includes("Stringmaster")) { //sorry... hardcoding just this once
        bonusAttack = .12 + weaponData[this.owner].rank * 0.03;
      }
    }*/
    var totalBuffMap = this.totalBuffMap;
    var attack = (charData[this.owner].attack + weaponData[this.owner].attack) * (1 + totalBuffMap.get('Attack') + bonusStats[this.owner].attack + bonusAttack);
    var health = (charData[this.owner].health + weaponData[this.owner].health) * (1 + totalBuffMap.get('Health') + bonusStats[this.owner].health);
    var defense = (charData[this.owner].defense + weaponData[this.owner].defense) * (1 + totalBuffMap.get('Defense') + bonusStats[this.owner].defense);
    var critMultiplier = (1 - Math.min(1,(charData[this.owner].crit + totalBuffMap.get('Crit')))) * 1 + Math.min(1,(charData[this.owner].crit + totalBuffMap.get('Crit'))) * (charData[this.owner].critDmg + totalBuffMap.get('Crit Dmg'));
    var damageMultiplier = getDamageMultiplier(this.type, totalBuffMap);
    var totalDamage = this.damage * attack * critMultiplier * damageMultiplier * (weaponData[this.owner].weapon.name === 'Nullify Damage' ? 0 : 1) * 0.48 * skillLevelMultiplier;
    console.log(`passive proc damage: ${this.damage.toFixed(2)}; attack: ${(charData[this.owner].attack + weaponData[this.owner].attack).toFixed(2)} x ${(1 + totalBuffMap.get('Attack') + bonusStats[this.owner].attack).toFixed(2)}; crit mult: ${critMultiplier.toFixed(2)}; dmg mult: ${damageMultiplier.toFixed(2)}; total dmg: ${totalDamage.toFixed(2)}`);
    this.totalDamage += totalDamage * this.procMultiplier;
    this.procMultiplier = 1;
    return totalDamage;
  }

  /**
   * Returns a note to place on the cell.
   */
  getNote() {
    return 'This skill triggered a passive damage effect: ' + this.name + ', which has procced ' + this.numProcs + ' times for ' + this.totalDamage + ' in total.';
  }
}


/**
 * Extracts the skill reference from the skillData object provided, with the name of the current character.
 * Skill data objects have a (Character) name at the end of them to avoid duplicates. Jk, now they don't, but all names MUST be unique.
 */
function getSkillReference(skillData, name, character) {
  return skillData[name/* + " (" + character + ")"*/];
}

function extractNumberAfterX(inputString) {
  var match = inputString.match(/x(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

function reverseTranslateClassificationCode(code) {
  const classifications = {
    'Normal': 'No',
    'Heavy': 'He',
    'Skill': 'Sk',
    'Liberation': 'Rl',
    'Spectro': 'Sp',
    'Fusion': 'Fu',
    'Electro': 'El',
    'Aero': 'Ae',
    'Spectro': 'Sp',
    'Havoc': 'Ha',
    'Physical': 'Ph',
    'Echo': 'Ec'
  };
  return classifications[code] || code; // Default to code if not found
}

function translateClassificationCode(code) {
  const classifications = {
    'No': 'Normal',
    'He': 'Heavy',
    'Sk': 'Skill',
    'Rl': 'Liberation',
    'Sp': 'Spectro',
    'Fu': 'Fusion',
    'El': 'Electro',
    'Ae': 'Aero',
    'Sp': 'Spectro',
    'Ha': 'Havoc',
    'Ph': 'Physical',
    'Ec': 'Echo'
  };
  return classifications[code] || code; // Default to code if not found
}

function getDamageMultiplier(classification, totalBuffMap) {
  let damageMultiplier = 1;
  let damageBonus = 1;
  let damageDeepen = 0;

  // loop through each pair of characters in the classification string
  for (let i = 0; i < classification.length; i += 2) {
    let code = classification.substring(i, i + 2);
    let classificationName = translateClassificationCode(code);
    // if classification is in the totalBuffMap, apply its buff amount to the damage multiplier
    if (totalBuffMap.has(classificationName)) {
      if(STANDARD_BUFF_TYPES.includes(classificationName)) { //check for deepen effects as well
        let deepenName = classificationName + " (Deepen)";
        if (totalBuffMap.has(deepenName))
          damageDeepen +=  totalBuffMap.get(deepenName);
        damageBonus += totalBuffMap.get(classificationName);
      } else {
        damageBonus += totalBuffMap.get(classificationName);
      }
    }
  }
  damageDeepen += totalBuffMap.get('Deepen');
  damageBonus += totalBuffMap.get('Specific');
  return damageMultiplier * damageBonus * (1 + totalBuffMap.get('Multiplier')) * (1 + damageDeepen);
}

function getSkills() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ActiveChar');
  var range = sheet.getDataRange();
  var values = range.getValues();

  // filter rows where the first cell is not empty
  var filteredValues = values.filter(function(row) {
    return row[0].toString().trim() !== ''; // Ensure that the name is not empty
  });

  var objects = filteredValues.map(rowToActiveSkillObject);
  return objects;
}

function getActiveEffects() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ActiveEffects');
  var range = sheet.getDataRange();
  var values = range.getValues();

  var objects = values.map(rowToActiveEffectObject).filter(effect => effect !== null);
  return objects;
}

function getWeapons() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Weapons');
  var range = sheet.getDataRange();
  var values = range.getValues();

  var weaponsMap = {};

  // Start loop at 1 to skip header row
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) { // Check if the row actually contains a weapon name
      var weaponInfo = rowToWeaponInfo(values[i]);
      weaponsMap[weaponInfo.name] = weaponInfo; // Use weapon name as the key for lookup
    }
  }

  return weaponsMap;
}

function getEchoes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Echo');
  var range = sheet.getDataRange();
  var values = range.getValues();

  var echoMap = {};

  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) { // check if the row actually contains an echo name
      var echoInfo = rowToEchoInfo(values[i]);
      echoMap[echoInfo.name] = echoInfo; // Use echo name as the key for lookup
    }
  }

  return echoMap;
}


function getCharacterConstants() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Constants');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var charConstants = {};

  for (var i = 3; i < values.length; i++) {
    if (values[i][0]) { // check if the row actually contains a name
      var charInfo = rowToCharacterConstants(values[i]);
      charConstants[charInfo.name] = charInfo; // use weapon name as the key for lookup
    } else {
      break;
    }
  }

  return charConstants;
}


function test() {
  var effects = getActiveEffects();
  effects.forEach(effect => {
    Logger.log("Name: " + effect.name +
      ", Type: " + effect.type +
      ", Classifications: " + effect.classifications +
      ", Buff Type: " + effect.buffType +
      ", Amount: " + effect.amount +
      ", Duration: " + effect.duration +
      ", Active: " + effect.active +
      ", Triggered By: " + effect.triggeredBy +
      ", Stack Limit: " + effect.stackLimit +
      ", Stack Interval: " + effect.stackInterval
    );
  });
}



function addProduct() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var selection = [];
  for (var i = 0; i < 20; i++) {
    if(!data[i][0])break;
    selection.push(data[1][i])
    //Logger.log(i+":"+data[1][i]);
  }
  
  var size = selection[7];
  var shieldPenalty = selection[6];
  var shieldsSkill = selection[3];
  var shieldSpeedPenalty = calcShieldSpeedPenalty(size, shieldPenalty, shieldsSkill);
  var str = selection[1];
  var armourSkill = selection[4];
  var armourEncumbrance = selection[5];
  var armourSpeedPenalty = calcArmourSpeedPenalty(str, armourSkill, armourEncumbrance);
  var acKeysWeapon1 = {}, acKeysWeapon2 = {};
  for (var i = 3; i < data.length; i++) {
    if(data[i][3]==null)break;
    acKeysWeapon1[data[i][3]] = [];
    acKeysWeapon2[data[i][3]] = [];
  }
  //Logger.log(Object.keys(acKeysWeapon1));
  var dex = selection[2];
  var fighting = selection[3];
  var w1base_damage = selection[8];
  var w1base_delay = selection[9];
  var w1min_delay = selection[10];
  var w1slaying = selection[11];
  var w1brand = selection[12];
  var w1ranged = selection[13];
  var w2base_damage = selection[14];
  var w2base_delay = selection[15];
  var w2min_delay = selection[16];
  var w2slaying = selection[17];
  var w2brand = selection[18];
  var w2ranged = selection[19];

  var acValues = Object.keys(acKeysWeapon1);
  for(var i = 0; i < acValues.length; i++){
    for(var j = 0; j < 28; j++){
      //Logger.log('calcdamage: ' + calcDamage(sPenalty, aPenalty, i));
      // Logger.log(ac);
      // Logger.log(acKeysWeapon1[ac]);
      acKeysWeapon1[acValues[i]].push(
        calcDamage(shieldSpeedPenalty, armourSpeedPenalty, j, str, dex, fighting, w1base_damage, w1base_delay, w1min_delay, w1slaying, w1brand, w1ranged, acValues[i]))
      acKeysWeapon2[acValues[i]].push(
        calcDamage(shieldSpeedPenalty, armourSpeedPenalty, j, str, dex, fighting, w2base_damage, w2base_delay, w2min_delay, w2slaying, w2brand, w2ranged, acValues[i]))
    }  
  }  

  var ratiosAc = {};
  for(var i = 0; i < acValues.length; i++){
    var ratiosArr = [];
    for(var j = 0; j < 28; j++){
      var a = acKeysWeapon1[acValues[i]][j]/acKeysWeapon2[acValues[i]][j];
      ratiosArr.push(acKeysWeapon1[acValues[i]][j]/acKeysWeapon2[acValues[i]][j]);
    }
    ratiosAc[acValues[i]] = [ratiosArr];
  }
  
  for (var i = 3; i < data.length; i++) {
    if(data[i][3]==null)break;
    var ac = data[i][3];
    var range = sheet.getRange("G"+(i+1)+":AH"+(i+1));
    range.setValues(ratiosAc[ac]);
  }

    var rawDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RawDamage");
    for(var i = 0; i < acValues.length; i++){
        var cell = rawDataSheet.getRange("B"+(i+2));
        cell.setValue(acValues[i]);
        cell = rawDataSheet.getRange("A"+(i+2));
        cell.setValue("w1");
        var range = rawDataSheet.getRange("C"+(i+2)+":AD"+(i+2));
        range.setValues([acKeysWeapon1[acValues[i]]]);
        cell = rawDataSheet.getRange("B"+(acValues.length+i+2));
        cell.setValue(acValues[i]);
        cell = rawDataSheet.getRange("A"+(acValues.length+i+2));
        cell.setValue("w2");
        range = rawDataSheet.getRange("C"+(acValues.length+i+2)+":AD"+(acValues.length+i+2));
        range.setValues([acKeysWeapon1[acValues[i]]]);
  }
}

// add value to a dictionary entry, creating the entry if needed
function addToEntry(dict, key, value)
{
    if (dict[key] == null)
        dict[key] = value;
    else
        dict[key] += value;
}

// reduce damage based on defender AC
function applyACReduction(weightedDamage, monsterAC)
{
    let enemy_ac = monsterAC;
    if (enemy_ac > 0) {
        let prevWeightedDamage = weightedDamage;
        weightedDamage = {};

       for (const [damage, weight] of Object.entries(prevWeightedDamage)) {
           var dam = parseInt(damage);
           for (var saved = 0; saved <= enemy_ac; saved++) {
                // damage can't go below zero
                var newDam = Math.max(0, dam - saved);
                addToEntry(weightedDamage, newDam, weight);
            }
        }
    }

    return weightedDamage;
}

function calcShieldPenalty(size, shieldPenalty, shieldsSkill)
{
    var  racialFactor = 0; // default for most medium species
    if (size == "little") {
         racialFactor = 4;
    }
    else if (size == "small") {
         racialFactor = 2;
    }
    else if (size == "large" ) {
        // Formicid is a special case: Due to having an extra set of arms,
        // this medium species gets the reduced shield penalty of large species
         racialFactor = -2;
    }
    var penalty = shieldPenalty;
    penalty = 2 * penalty * penalty;
    penalty *= (270 - shieldsSkill * 10);
    penalty /= (5 * (20 - 3 * racialFactor));
    penalty /= 270;

    return Math.max(0, penalty);
}

function calcShieldSpeedPenalty(size, sPenalty, shieldsSkill) {
    var shieldPenalty = calcShieldPenalty(size, sPenalty, shieldsSkill );


    if (shieldPenalty == 0) {
        return 0;
    }

    // scale for the roll
    var scale = 20;
    var scaledPenalty = shieldPenalty * scale;

    weightedValues = {};

    // DCSS rolls two dice with scaled shieldPenalty sides and takes the lower one
    for (var i = 1; i <= scaledPenalty; i++) {
        for (var j = 1; j <= scaledPenalty; j++) {
            addToEntry(weightedValues, Math.min(i, j), 1);
        }
    }

    weightedValues2 = {};

    for (var v = 1; v <= scaledPenalty; v++) {
        var key = Math.floor(v / scale);
        var rem = v % scale;
        addToEntry(weightedValues2, key, weightedValues[v] * (scale - rem));
        if (rem != 0) {
            addToEntry(weightedValues2, key + 1, weightedValues[v] * rem);
        }
    }

    var avg = getWeightedAverage(weightedValues2);

    // this is in auts, so convert to turns
    return avg / 10;
}

function calcArmourPenalty(str, armourSkill, armourEncumbrance)
{
    // Ref: player.cc

    var base_ev_penalty = armourEncumbrance / 10;

    var penalty =  2 / 5 * base_ev_penalty * base_ev_penalty / (str + 3)

    penalty *= (450 - armourSkill*10) / 450;

    return penalty;
}

function calcArmourSpeedPenalty(str, armourSkill, armourEncumbrance) {

    var penalty = calcArmourPenalty(str, armourSkill, armourEncumbrance);

    // convert from auts to turns
    return penalty / 10;
}

// work out average of weighted values
function getWeightedAverage(weightedValues)
{
    var sum = 0;
    var count = 0;
    for (const [value, weight] of Object.entries(weightedValues)) {
        var val = parseInt(value);
        count += weight;
        sum += (val * weight)
    }
    return count == 0 ? 0 : sum/count;
}

// Ref: attack::calc_damage() method in:
// https://github.com/crawl/crawl/blob/master/crawl-ref/source/attack.cc 
function calcDamage(shieldSpeedPenalty, armourSpeedPenalty, skill, str, dex, fighting, base_damage, base_delay, min_delay, slaying, brand, ranged, monsterAC)
{

    var weaponSkill = skill;
    var stat = ranged ? dex : str;
    var delay = base_delay;

    delay -= 0.1 * weaponSkill/2.0;

    delay = Math.max(delay, min_delay);

    if (brand == "speed") {
        delay = 2.0/3.0 * delay;
    }
    // all possible damage values, weighted by probability
    var weightedDamage = {};
    var prevWeightedDamage;


      if (brand == "heavy") {
          base_damage *= 1.8; // +80% damage
      }
      weightedDamage[base_damage] = 1;
  

    // stat modifier

    prevWeightedDamage = weightedDamage;
    weightedDamage = {};

    // max(1.0, 75 + 2.5 * stat) / 100
    for (const [damage, weight] of Object.entries(prevWeightedDamage)) {
        var dam = parseInt(damage);
        var newDam = Math.floor(dam * Math.max(1.0, 75 + 2.5 * stat));
        newDam = Math.floor(newDam / 100);
        addToEntry(weightedDamage, newDam, weight);
    }
 

    // at this point, damage is randomized to a number between 0 and the full amount
    prevWeightedDamage = weightedDamage;
    weightedDamage = {};

    for (const [damage, weight] of Object.entries(prevWeightedDamage)) {
        var dam = parseInt(damage);
        for (var i = 0; i <= dam; i++) {
            addToEntry(weightedDamage, i, weight);
        }
    }

    // weapon skill modifier
    // [2500 + (random2(you.skill(wpn_skill, 100) + 1))] / 2500
    // = 1 + (0->weapon_skill*100)/2500


      prevWeightedDamage = weightedDamage;
      weightedDamage = {};

      for (const [damage, weight] of Object.entries(prevWeightedDamage)) {
          var dam = parseInt(damage);
          for (var i = 0; i <= weaponSkill*100; i++) {
              var newDam = Math.floor(dam + dam*i/2500);
              addToEntry(weightedDamage, newDam, weight);
          }
      }


    // fighting skill modifier
    // [30 * 100 + (random2(you.skill(SK_FIGHTING, 100) + 1))] / (30 * 100)
    // = 1 + (0->fighting_skill*100)/3000

    prevWeightedDamage = weightedDamage;
    weightedDamage = {};

    for (const [damage, weight] of Object.entries(prevWeightedDamage)) {
        var dam = parseInt(damage);
        for (var i = 0; i <= fighting*100; i++) {
            var newDam = Math.floor(dam + dam*i/3000);
            addToEntry(weightedDamage, newDam, weight);
        }
    }

    // slaying bonus
    // a random number between 0 and effective enchantment (which can be negative)
    var effective_enchant = slaying;
    var slay_bonus_min = Math.min(effective_enchant, 0)
    var slay_bonus_max = Math.max(effective_enchant, 0)

    prevWeightedDamage = weightedDamage;
    weightedDamage = {};

    for (const [damage, weight] of Object.entries(prevWeightedDamage)) {
        var dam = parseInt(damage);
        for (var i = slay_bonus_min; i <= slay_bonus_max; i++) {
            var newDam = dam + i;
            addToEntry(weightedDamage, newDam, weight);
        }
    }

    // apply ac reduction
    weightedDamage = applyACReduction(weightedDamage, monsterAC);

    // work out the weighted average
    var avg_damage = getWeightedAverage(weightedDamage);

    // convert weights to percentages
    var sumWeights = 0;
    for (const [damage, weight] of Object.entries(weightedDamage)) {
        sumWeights += weight;
    }
    for (const [damage, weight] of Object.entries(weightedDamage)) {
        weightedDamage[damage] = 100 * weight / sumWeights;
    }

    var damage_per_hit = {}
    damage_per_hit["base"] = avg_damage;
    damage_per_hit["base_distro"] = weightedDamage;

    damage_per_hit["brand"] = 0.0;
    if (brand == "vorpal") {
        // 0-33% on melee weapons -> avg = 16.7%
        // TODO: handle ranged (apparently 20%)
        damage_per_hit["brand"] = 0.167 * damage_per_hit["base"];
    }
    else if (brand == "flame" || brand == "freeze") {
        // 0-50% -> avg = 25%
        damage_per_hit["brand"] = 0.25 * damage_per_hit["base"];
    }
    else if (brand == "flame+freeze") {
        damage_per_hit["brand"] = 0.5 * damage_per_hit["base"];
    }
    else if (brand == "holy") {
        // 0-150% -> avg = 75%
        damage_per_hit["brand"] = 0.75 * damage_per_hit["base"];
    }
    else if (brand == "drain") {
        // 0-50% + 1+1d3 -> avg = 25% + 2
        damage_per_hit["brand"] = (0.25 * damage_per_hit["base"]) + 2.0;
    }
    else if (brand == "elec") {
        // chance to trigger is 1/4 (1/3 prior to 0.28)
        trigger_chance =  1/4;
        // if triggered, it does 8 + rand2(13) dmg -> 8 + [0 to 12] -> avg = 14
        damage_per_hit["brand"] = 14 * trigger_chance;
    }
    else if (brand == "disrupt") {
        // only found on the unrand artefact "Undeadhunter"
        // has 1/3 chance to inflict random2avg((1 + (dam * 3)), 3);
        // random2avg(x, 3) returns (random2(x) + random2(x+1) + random2(x+1))/3
        // so avg when it triggers is (3*dam + 3*dam+1 + 3*dam+1)/2/3
        // = (9*dam+2)/6
        // divide by 3 because it only triggers 1/3 of the time: avg = (9*dam+2)/18
        damage_per_hit["brand"] = (9.0 * damage_per_hit["base"] + 2.0) / 18.0;
    }
    else if (brand == "silver") {
        // flat 75% on chaotic monsters
        damage_per_hit["brand"] = 0.75 * damage_per_hit["base"];
        //TODO: (1 + random2(damage_done) / 3) on others
    }
    else if (brand == "slay drac") {
        // bonus_dam = 1 + random2(3 * dam / 2);
        // avg = 1 + 75% * dam
        damage_per_hit["brand"] = 1 + 0.75 * damage_per_hit["base"];
    }
    else if (brand == "spect") {
        //damage_per_hit["brand"] = calcSpectralDamage(weapon);
    }
    else if (brand == "discharge") {
        // 1 in 3 chance of casting discharge with an average power of 150
        // damage when it triggers: 3 + random2(5 + pow / 10 + (random2(pow) / 10));
        damage_per_hit["brand"] = (3 + (20 + 15/2) / 2) / 3;
    }

    damage_per_hit["total"] = damage_per_hit["base"] + damage_per_hit["brand"];
    //weapon["damage_per_hit"] = damage_per_hit;

    var delay = base_delay;
    delay -= 0.1 * weaponSkill/2.0;
    
    delay = Math.max(delay, min_delay);
    if (brand == "speed") {
        delay = 2.0/3.0 * delay;
    }
    else if (brand == "heavy") {
        delay = calcHeavyDelay(delay);
    }
    
    delay += shieldSpeedPenalty;

    if (ranged) {
        delay += armourSpeedPenalty;
    }

    var damage_per_turn = {}
    damage_per_turn["base"] = damage_per_hit["base"] / delay;
    damage_per_turn["brand"] = damage_per_hit["brand"] / delay;
    damage_per_turn["total"] = damage_per_hit["total"] / delay;
    return damage_per_turn["total"];
}
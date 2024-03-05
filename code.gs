function Day1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = ss.getSheetByName("Registry");
  const s1 = ss.getSheetByName("Day 1 - Peloton Finishers");

  var regVals = registry.getDataRange().getValues().slice(1);
  var s1Vals = s1.getDataRange().getValues().slice(1);

  const mergedDay1 = s1Vals.map(mergeRegistry);

  function mergeRegistry(row) {
    const bibNum = row[0];
    for (i=0; i<regVals.length; i++) {
      const regBib = regVals[i][0];
      if (bibNum == regBib) {
        var revRow = [bibNum].concat(regVals[i].slice(1), row.slice(1,5).join(":"));
        if (revRow[revRow.length-1] == ":::") {
          return revRow.slice(0,revRow.length-1).concat([""]);
        } else {
          return revRow;
        }
      }
    }
  }

  var dict = {};

  for (i=0; i<mergedDay1.length; i++) {
    const row = mergedDay1[i];
    dict[row[0]] = row.slice(1,row.length);
  }
  
  for (i=0; i<regVals.length; i++) {
    const bib = regVals[i][0];
    if (!(bib in dict)) {
      dict[bib] = regVals[i].slice(1,regVals[i].length).concat([""])
    }
  }
  
  for (var [bib, data] of Object.entries(dict)) {
    const time = data[3];
    if (time != "") {
      dict[bib] = data.slice(0,data.length-1).concat([timeToMs(time)]);
    } 
  }

  calculateRankings(dict, "Day 1 - Results");
  return dict;
}

function Day2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s2 = ss.getSheetByName("Day 2 - TT Finishers");

  var s2Vals = s2.getDataRange().getValues().slice(1);
  var dict = Day1();

  var raceCheck = Object.keys(dict).map(Number);

  for (i=0; i<s2Vals.length; i++) {
    const bib = s2Vals[i][0];
    const time = dict[bib][dict[bib].length-1];
    const newTime = timeToMs(s2Vals[i].slice(1,5).join(":"));

    if (time != "" && newTime != 0) {
      dict[bib][dict[bib].length-1] = time + newTime;
    } else {
      dict[bib][dict[bib].length-1] = "";
    }

    if (raceCheck.includes(bib)) {
      raceCheck.splice(raceCheck.indexOf(bib),1);
    }
  }

  if (raceCheck.length != 0) {
    for (i=0; i<raceCheck.length; i++) {
      dict[raceCheck[i]][dict[raceCheck[i]].length-1] = "";
    }
  }

  calculateRankings(dict, "Day 2 - Results");
  return dict;
}

function Day3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s3 = ss.getSheetByName("Day 3 - Mountain Finishers");

  var s3Vals = s3.getDataRange().getValues().slice(1);
  var dict = Day2();

  var raceCheck = Object.keys(dict).map(Number);

  for (i=0; i<s3Vals.length; i++) {
    const bib = s3Vals[i][0];
    const time = dict[bib][dict[bib].length-1];
    const newTime = timeToMs(s3Vals[i].slice(1,5).join(":"));

    if (time != "" && newTime != 0) {
      dict[bib][dict[bib].length-1] = time + newTime;
    } else {
      dict[bib][dict[bib].length-1] = "";
    }

    if (raceCheck.includes(bib)) {
      raceCheck.splice(raceCheck.indexOf(bib),1);
    }
  }

  if (raceCheck.length != 0) {
    for (i=0; i<raceCheck.length; i++) {
      dict[raceCheck[i]][dict[raceCheck[i]].length-1] = "";
    }
  }
  
  calculateRankings(dict, "GC - Results");
}

function calculateRankings(dict, sheetName) {
  var arr = Object.entries(dict).map(x => [].concat.apply([],x));
  const sorted = arr.sort((a,b) => (a[a.length-1] > b[b.length-1] ? 1:-1));
  var [women, men39, men40, men50, junior, dnf] = [[], [], [], [], [], []];
  var [wL, m39L, m40L, m50L, jL] = [0,0,0,0,0];
  const groupKey = {"Women": women, "Men 39 and Under": men39, "Men 40+": men40, "Men 50+": men50, "Junior": junior};
  const leaderKey = {"Women": wL, "Men 39 and Under": m39L, "Men 40+": m40L, "Men 50+": m50L, "Junior": jL};
  
  for (i=0; i<sorted.length; i++) {
    const group = sorted[i][2];

    if (sorted[i][sorted[i].length-1] == "") {
      sorted[i].unshift("DNF");
      sorted[i].push("");
      dnf.push(sorted[i]);
    } else {
      if (groupKey[group].length == 0) {
        groupKey[group].push([group, "", "", "", "", "",]);
        groupKey[group].push(["Place", "Bib #", "Name", "Team", "Time", "+Time"]);
        sorted[i].unshift(groupKey[group].length-1);
        leaderKey[group] = sorted[i][sorted[i].length-1];
        sorted[i][sorted[i].length-1] = calcDiff(sorted[i][sorted[i].length-1]);
        sorted[i].push("");
        sorted[i].splice(3,1);
        groupKey[group].push(sorted[i]);
      } else {
        sorted[i].unshift(groupKey[group].length-1);
        const diff = calcDiff(sorted[i][sorted[i].length-1] - leaderKey[group]);
        sorted[i][sorted[i].length-1] = calcDiff(sorted[i][sorted[i].length-1]);
        sorted[i].push("+" + diff);
        sorted[i].splice(3,1);
        groupKey[group].push(sorted[i]);
      }
    }
  }

  for (i=0; i<dnf.length; i++) {
    const group = dnf[i][3];
    dnf[i].splice(3,1);
    groupKey[group].push(dnf[i]);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var gcSheet = ss.getSheetByName(sheetName);
  if (gcSheet != null) {
    ss.deleteSheet(gcSheet);
  }
  gcSheet = ss.insertSheet();
  gcSheet.setName(sheetName);

  const allGroups = [women, men39, men40, men50, junior];
  var curRow = 1;

  for (i=0; i<allGroups.length; i++) {
    const numRows = allGroups[i].length;
    const numCols = allGroups[i][0].length;
    gcSheet.getRange(curRow,1,numRows,numCols).setValues(allGroups[i]);
    curRow += numRows + 1;
  }
}

function timeToMs(time) {
  const timeList = time.split(":").map(Number);
  return timeList[0]*3600000+timeList[1]*60000+timeList[2]*1000+timeList[3];
}

function calcDiff(ms) {
  const mss = Math.floor(ms % 1000);
  var secs = Math.floor(ms/1000);
  var mins = Math.floor(secs/60);
  var hrs = Math.floor(mins/60);

  secs = secs % 60;
  mins = mins % 60;
  hrs = hrs % 24;

  return `${hrs.toString().padStart(2,'0')}:${mins.toString().padStart(2,'0')}:${secs.toString().padStart(2,'0')}.${mss.toString().padStart(3,'0')}`
}

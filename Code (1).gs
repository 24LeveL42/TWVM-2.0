function doGet(e) {
  return doPost(e);
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.openById('16u962fiajqS5hat76mrIOUHjgVW4YKZft-UXiZES9o0').getActiveSheet();
    var p = e.parameter;
    var ts = new Date().toLocaleString('en-GB');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp','Location','Date','Time',
        'Track Veh Type','Track Veh MID',
        'Wheel Veh Type','Wheel Veh MID',
        'Unit','Rank name','Defects of Vehicle',
        'Fuel Level (%)','Engine Hr','Mileage'
      ]);
    }

    sheet.appendRow([
      ts,
      p.location||'',
      p.date||'',
      p.time||'',
      p.trackVehType||'',
      p.trackVehMid||'',
      p.wheelVehType||'',
      p.wheelVehMid||'',
      p.unit||'',
      p.rankName||'',
      p.defects||'',
      p.fuelLevel||'',
      p.engineHr||'',
      p.mileage||''
    ]);

    return ContentService.createTextOutput(JSON.stringify({status:'success'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

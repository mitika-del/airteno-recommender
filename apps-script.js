// ============================================================================
// GOOGLE APPS SCRIPT — paste this into Extensions > Apps Script in Google Sheets
// Set trigger: onFormSubmit > From spreadsheet > On form submit
// ============================================================================
//
// Column order matches Google Form field order (v[0] = timestamp):
//   v[1]  date_of_visit
//   v[2]  surveyor_name          (Murshid / Nazir)
//   v[3]  site_type
//   v[4]  site_name
//   v[5]  address
//   v[6]  carpet_area            (sq ft)
//   v[7]  floors
//   v[8]  num_rooms
//   v[9]  room_size
//   v[10] ceiling_height         (ft)
//   v[11] ac_type
//   v[12] layout_upload          (file upload — Drive URL)
//   v[13] other_spaces
//   v[14] install_location
//   v[15] filter_access
//   v[16] inlet_method
//   v[17] power_available
//   v[18] pollution_proximity    (near road / construction — paragraph)
//   v[19] rain_protection
//   v[20] obstructions
//   v[21] doors_to_open
//   v[22] mount_position
//   v[23] site_notes
//   v[24] existing_purifiers
//   v[25] construction_status    (Existing construction / New / Renovation)
//   v[26] health_conditions
// ============================================================================

var API_URL = 'https://airteno-recommender-3grlj4mlx-mitika-dels-projects.vercel.app/api/generate'; // update after vercel deploy

// Surveyor email lookup — set both in Apps Script project properties or hardcode here.
var SURVEYOR_EMAILS = {
  'Murshid': PropertiesService.getScriptProperties().getProperty('MURSHID_EMAIL') || '',
  'Nazir':   PropertiesService.getScriptProperties().getProperty('NAZIR_EMAIL')   || ''
};

// Supported MIME types for Claude vision. PDFs are not supported.
var SUPPORTED_IMAGE_TYPES = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];

function extractDriveFileId(url) {
  if (!url) return null;
  var s = url.toString();
  var m = s.match(/[?&]id=([^&]+)/);
  if (m) return m[1];
  m = s.match(/\/file\/d\/([^/]+)/);
  if (m) return m[1];
  return null;
}

// Handles multi-file upload: takes the first supported image found.
function getLayoutImageData(fieldValue) {
  if (!fieldValue || fieldValue.toString().trim() === '') return null;
  var urls = fieldValue.toString().split(',');
  for (var i = 0; i < urls.length; i++) {
    var fileId = extractDriveFileId(urls[i].trim());
    if (!fileId) continue;
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var mimeType = blob.getContentType();
      if (SUPPORTED_IMAGE_TYPES.indexOf(mimeType) === -1) {
        Logger.log('Layout file skipped — unsupported type: ' + mimeType);
        continue;
      }
      return {
        b64:  Utilities.base64Encode(blob.getBytes()),
        type: mimeType
      };
    } catch (err) {
      Logger.log('Layout image error: ' + err.toString());
    }
  }
  return null;
}

function onFormSubmit(e) {
  var v = e.values;

  var layoutData = getLayoutImageData(v[12]);

  var payload = {
    date_of_visit:        v[1],
    surveyor_name:        v[2],
    surveyor_email:       SURVEYOR_EMAILS[v[2]] || '',
    site_type:            v[3],
    site_name:            v[4],
    address:              v[5],
    carpet_area:          v[6],
    floors:               v[7],
    num_rooms:            v[8],
    room_size:            v[9],
    ceiling_height:       v[10],
    ac_type:              v[11],
    other_spaces:         v[13],
    install_location:     v[14],
    filter_access:        v[15],
    inlet_method:         v[16],
    power_available:      v[17],
    pollution_proximity:  v[18],
    rain_protection:      v[19],
    obstructions:         v[20],
    doors_to_open:        v[21],
    mount_position:       v[22],
    site_notes:           v[23],
    existing_purifiers:   v[24],
    construction_status:  v[25],
    health_conditions:    v[26],
    layout_image_b64:     layoutData ? layoutData.b64  : null,
    layout_image_type:    layoutData ? layoutData.type : null
  };

  var options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(API_URL, options);
    Logger.log('Response: ' + response.getContentText());
  } catch (err) {
    Logger.log('Error calling API: ' + err.toString());
  }
}

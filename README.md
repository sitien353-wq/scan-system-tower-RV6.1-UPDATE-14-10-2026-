// Hàm tiện ích để loại bỏ hậu tố "-1" khỏi chuỗi
function removeMinusOne(str) {
  if (typeof str !== 'string') return str.toString().trim();
  if (str.endsWith("-1")) {
    return str.slice(0, -2).trim(); // Loại bỏ "-1" và khoảng trắng
  }
  return str.trim();
}

function checkAllData() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Không thể truy cập spreadsheet hiện tại.');
    }

    var ui = SpreadsheetApp.getUi();
    if (!ui) {
      throw new Error('Không thể khởi tạo giao diện người dùng.');
    }

    // Hiển thị giao diện "Đang kiểm tra"
    var checkingHtml = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">' +
      '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">' +
      '<style>' +
      'body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #e6f3ff 0%, #d1e8ff 100%); }' +
      '.progress-bar { width: 100%; background-color: #e0e0e0; height: 10px; border-radius: 5px; overflow: hidden; margin-top: 20px; }' +
      '.progress { width: 0; height: 100%; background-color: #10b981; animation: progress 2s infinite; }' +
      '@keyframes progress { 0% { width: 0; } 100% { width: 100%; } }' +
      '</style>' +
      '</head>' +
      '<body class="flex items-center justify-center min-h-screen">' +
      '<div class="bg-white p-6 rounded-lg shadow-xl max-w-md w-full text-center">' +
      '<h2 class="text-2xl font-bold text-gray-800 mb-4">Đang kiểm tra...</h2>' +
      '<h3 class="text-xl font-semibold text-blue-800 mb-4">CẤU TẠO MÃ KHUÔN LINH KIỆN</h3>' +
      '<p class="text-gray-600 mt-4">Vui lòng chờ trong khi dữ liệu được kiểm tra.</p>' +
      '<div class="progress-bar"><div class="progress"></div></div>' +
      '<audio id="startSound" autoplay><source src="https://www.soundjay.com/buttons/beep-02.mp3" type="audio/mpeg"></audio>' +
      '<script>' +
      'document.getElementById("startSound").play().catch(function(error) { console.log("Lỗi phát âm thanh: ", error); });' +
      '</script>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(400).setHeight(300);
    ui.showModalDialog(checkingHtml, 'Đang kiểm tra');

    // Đợi 1.5 giây để hiển thị giao diện "Đang kiểm tra"
    Utilities.sleep(1500);

    // Định nghĩa các sheet cần kiểm tra
    var sheetsToCheck = {
      'VA001-B': { checkFunction: checkSheet8794B },
      'VA015-A': { checkFunction: checkSheet7040GA },
      'VA015-B': { checkFunction: checkSheet7040GB },
      'VA015-E': { checkFunction: checkSheet7040GE },
      'VS001-A': { checkFunction: checkSheetVS001A },
      'VS001-B': { checkFunction: checkSheetVS001B },
      'VS001-C': { checkFunction: checkSheetVS001C },
      'VA019-A': { checkFunction: checkSheetVA019A },
      'VA019-B': { checkFunction: checkSheetVA019B },
      'VA019-C': { checkFunction: checkSheetVA019C },
      'VA003': { checkFunction: checkSheetVA003 },
      'VS051': { checkFunction: checkSheetVS051 },
      'VA016-A': { checkFunction: checkSheet7041GA },
      'VA016-B': { checkFunction: checkSheet7041GB },
    };

    // Tạo hoặc lấy sheet log và tự động điều chỉnh cột
    var logSheet = spreadsheet.getSheetByName('CheckLog') || spreadsheet.insertSheet('CheckLog');
    if (!logSheet) {
      throw new Error('Không thể tạo hoặc truy cập sheet CheckLog.');
    }
    logSheet.clear();
    logSheet.appendRow(['Thời gian', 'Tên Sheet', 'Trạng thái', 'Số lỗi', 'Thông báo']);

    // Thu thập kết quả và đánh dấu lỗi
    var results = {};
    var startTime = new Date();
    var startTimeFormatted = Utilities.formatDate(startTime, 'GMT+7', 'HH:mm:ss');
    var erroredSheet = null;
    var hasErrors = false;
    for (var sheetName in sheetsToCheck) {
      try {
        var sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          results[sheetName] = { valid: false, errorCount: 0, message: 'Sheet không tìm thấy', details: [] };
          logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Sheet không tìm thấy']);
          hasErrors = true;
          continue;
        }
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        if (lastRow < 7) {
          results[sheetName] = { valid: false, errorCount: 0, message: 'Không có dữ liệu từ hàng 7', details: [] };
          logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Không có dữ liệu từ hàng 7']);
          hasErrors = true;
          continue;
        }
       

        var result = sheetsToCheck[sheetName].checkFunction(sheet);
        results[sheetName] = result;
        logSheet.appendRow([new Date(), sheetName, result.valid ? 'OK' : 'LỖI', result.errorCount, result.message]);
        if (!result.valid && !erroredSheet) {
          erroredSheet = sheet;
          hasErrors = true;
        }
      } catch (e) {
        results[sheetName] = { valid: false, errorCount: 0, message: 'Lỗi xử lý sheet: ' + e.message, details: [] };
        logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Lỗi xử lý sheet: ' + e.message]);
        hasErrors = true;
      }
    }
    var endTime = new Date();
    var endTimeFormatted = Utilities.formatDate(endTime, 'GMT+7', 'HH:mm:ss');
    var duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log('Thời gian thực thi: ' + duration + ' giây');

    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.autoResizeColumns(1, 5);
      for (var i = 1; i <= lastRow; i++) {
        logSheet.setRowHeight(i, 21);
      }
    }

            var htmlOutput = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">' +
      '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">' +
      '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">' +
      '<style>' +
      'body { font-family: "Roboto", sans-serif; ' +
      (hasErrors 
        ? 'background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%); }'
        : 'background: linear-gradient(135deg, #e6f3ff 0%, #d1e8ff 100%); }') +
      '.time-badge { background-color: #e2e8f0; color: #2c3e50; padding: 8px 12px; border-radius: 20px; display: inline-block; margin: 4px; }' +
      'table { width: 100%; border-collapse: separate; border-spacing: 0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }' +
      'th { background-color: ' + (hasErrors ? '#dc2626' : '#10b981') + '; color: white; padding: 12px; text-align: left; }' +
      'td { padding: 12px; text-align: left; border: 1px solid #ddd; vertical-align: top; }' +
      'tr:nth-child(even) { background-color: #f8fafc; }' +
      'tr:hover { background-color: ' + (hasErrors ? '#fca5a5' : '#e6fffa') + '; transition: background-color 0.3s; }' +
      '.error-row { background-color: #dc2626; color: white; border: 2px solid #991b1b; }' +
      '.error-row:hover { background-color: #b91c1c; border-color: #7f1d1d; }' +
      '.status-icon { margin-right: 8px; }' +
      '.error-detail { font-size: 0.95em; line-height: 1.6; }' +
      '.error-detail ul { margin: 8px 0; padding-left: 22px; list-style-type: disc; }' +
      '.error-item { margin-bottom: 10px; padding: 10px; background-color: rgba(220, 38, 38, 0.15); border-left: 5px solid #dc2626; border-radius: 6px; }' +
      '.error-pos { font-weight: bold; color: #991b1b; }' +
      '.error-current { font-family: monospace; background: #ffffff; padding: 3px 8px; border-radius: 4px; font-weight: bold; color: #000000; border: 1px solid #999999; }' +
      '.error-expected { font-family: monospace; background: #ffffcc; padding: 3px 8px; border-radius: 4px; font-weight: bold; color: #000000; border: 1px solid #999999; }' +
      '.no-detail { color: #666; font-style: italic; }' +
      '</style>' +
      '</head>' +
      '<body class="flex items-center justify-center min-h-screen" onload="playSoundIfErrors(\'' + (hasErrors ? 'true' : 'false') + '\')">' +
      '<div class="bg-white p-8 rounded-lg shadow-2xl max-w-6xl w-full overflow-x-auto">' +
      '<h3 class="text-2xl font-semibold ' + (hasErrors ? 'text-red-700' : 'text-blue-800') + ' text-center mb-4">CẤU TẠO MÃ KHUÔN LINH KIỆN</h3>' +
      '<h2 class="text-3xl font-bold text-gray-800 text-center mb-6">Kết Quả Kiểm Tra</h2>' +
      '<div class="flex justify-center space-x-3 mb-6">' +
      '<span class="time-badge">Bắt đầu: ' + startTimeFormatted + '</span>' +
      '<span class="time-badge">Kết thúc: ' + endTimeFormatted + '</span>' +
      '<span class="time-badge">Thời gian: ' + duration.toFixed(2) + 's</span>' +
      '</div>' +
      '<div class="overflow-x-auto">' +
      '<table class="min-w-full">' +
      '<thead><tr>' +
      '<th class="px-4 py-3">Tên Sheet</th>' +
      '<th class="px-4 py-3">Trạng thái</th>' +
      '<th class="px-4 py-3">Số lỗi</th>' +
      '<th class="px-4 py-3">Thông báo tổng quát</th>' +
      '<th class="px-4 py-3" style="width: 450px;">Chi tiết lỗi cụ thể</th>' +
      '</tr></thead>' +
      '<tbody>' +
      (function() {
        var rows = [];
        function escapeHtml(text) {
          if (text === null || text === undefined) return '';
          return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
        }

        for (var sheetName in results) {
          var result = results[sheetName];
          var detailContent = '';

          if (result.valid) {
            detailContent = '<div class="text-green-600 font-medium">✓ Tất cả dữ liệu hợp lệ</div>';
          } else {
            if (result.details && Array.isArray(result.details) && result.details.length > 0) {
              detailContent = '<div class="error-detail"><ul>';
              result.details.forEach(function(err) {
                if (typeof err === 'object' && err !== null) {
                  var row = err.row || '?';
                  var col = err.col ? String.fromCharCode(64 + parseInt(err.col)) : '?';
                  var position = 'Hàng ' + row + ', Cột ' + col;

                  detailContent += '<li class="error-item">' +
                    '<span class="error-pos">' + position + ':</span> ' + escapeHtml(err.message || 'Lỗi không xác định') + '<br>';
                  if (err.current !== undefined) detailContent += '→ Hiện tại: <span class="error-current">' + escapeHtml(err.current) + '</span><br>';
                  if (err.expected !== undefined) detailContent += '→ Mong đợi: <span class="error-expected">' + escapeHtml(err.expected) + '</span>';
                  detailContent += '</li>';
                } else {
                  detailContent += '<li class="error-item">' + escapeHtml(err) + '</li>';
                }
              });
              detailContent += '</ul></div>';
            } else if (result.message) {
              detailContent = '<div class="error-item">' + escapeHtml(result.message) + '</div>';
            } else {
              detailContent = '<div class="no-detail">Có lỗi nhưng không có thông tin chi tiết</div>';
            }
          }

          rows.push('<tr' + (result.valid ? '' : ' class="error-row"') + '>' +
                    '<td class="border px-4 py-3 font-semibold">' + escapeHtml(sheetName) + '</td>' +
                    '<td class="border px-4 py-3"><i class="status-icon fa ' + (result.valid ? 'fa-check-circle text-green-600' : 'fa-times-circle text-white') + '"></i> <strong>' + (result.valid ? 'OK' : 'LỖI') + '</strong></td>' +
                    '<td class="border px-4 py-3 text-center text-lg font-bold">' + result.errorCount + '</td>' +
                    '<td class="border px-4 py-3">' + (result.message ? escapeHtml(result.message) : '-') + '</td>' +
                    '<td class="border px-4 py-3">' + detailContent + '</td>' +
                    '</tr>');
        }
        return rows.join('');
      })() +
      '</tbody></table>' +
      '</div>' +
      (hasErrors ? '<audio id="errorSound"><source src="https://www.soundjay.com/buttons/beep-01a.mp3" type="audio/mpeg"></audio>' : '') +
      '<script>' +
      'function playSoundIfErrors(hasErrors) {' +
      '  if (hasErrors === "true") {' +
      '    var audio = document.getElementById("errorSound");' +
      '    if (audio) audio.play().catch(function(error) { console.log("Lỗi phát âm thanh: ", error); });' +
      '  }' +
      '}' +
      '</script>' +
      '<div class="text-center mt-8">' +
      '<button onclick="google.script.host.close()" class="bg-' + (hasErrors ? 'red' : 'blue') + '-600 text-white px-8 py-4 rounded-lg hover:bg-' + (hasErrors ? 'red' : 'blue') + '-700 transition font-semibold text-lg shadow-lg">Đóng</button>' +
      '</div>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(950).setHeight(650);

    ui.showModalDialog(htmlOutput, hasErrors ? 'CẢNH BÁO HỆ THỐNG:  Phát hiện lỗi...Thông tin đã gửi đến quản lý để xử lý kịp thời ' : 'Tất cả dữ liệu đều OK');

    if (hasErrors) {
      try {
        var emailSubject = hasErrors ? 'CẢNH BÁO: Kết quả kiểm tra dữ liệu - Có lỗi' : 'Kết quả kiểm tra dữ liệu - Tất cả OK';
        var emailBody = `
  <h2 style="color:#dc2626;">⚠️ CẢNH BÁO: Phát hiện lỗi trong kiểm tra mã khuôn</h2>
  <p><strong>Thời gian bắt đầu:</strong> ${startTimeFormatted}</p>
  <p><strong>Thời gian kết thúc:</strong> ${endTimeFormatted}</p>
  <p><strong>Thời gian thực thi:</strong> ${duration.toFixed(2)} giây</p>
  <hr>
  <h3>📋 Tóm tắt kết quả</h3>
  <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">
    <tr style="background-color: #10b981; color: white;">
      <th style="padding: 10px;">Tên Sheet</th>
      <th style="padding: 10px;">Trạng thái</th>
      <th style="padding: 10px;">Số lỗi</th>
    </tr>
    ${Object.keys(results).map(function(sheetName) {
      var result = results[sheetName];
      return `<tr style="background-color: ${result.valid ? '#f0fdf4' : '#fee2e2'};">
                <td style="padding: 10px;"><strong>${sheetName}</strong></td>
                <td style="padding: 10px; color: ${result.valid ? '#166534' : '#dc2626'}; font-weight: bold;">${result.valid ? 'OK' : 'LỖI'}</td>
                <td style="padding: 10px; text-align: center; font-weight: bold; color: ${result.valid ? '#166534' : '#dc2626'};">${result.errorCount}</td>
              </tr>`;
    }).join('')}
  </table>
  <hr>
  <h3>🔍 Chi tiết lỗi cụ thể</h3>
  ${Object.keys(results).filter(sheetName => !results[sheetName].valid && results[sheetName].details && results[sheetName].details.length > 0).map(sheetName => {
    var result = results[sheetName];
    var detailList = result.details.map(err => {
      if (typeof err === 'object' && err !== null) {
        var position = err.row && err.col ? `Hàng ${err.row}, Cột ${String.fromCharCode(64 + parseInt(err.col))}` : 'Vị trí không xác định';
        var msg = err.message || 'Lỗi không xác định';
        var current = err.current !== undefined ? `<strong>Hiện tại:</strong> <span style="background:#fee2e2; padding:2px 6px; border-radius:4px;">${err.current}</span>` : '';
        var expected = err.expected !== undefined ? `<strong>Mong đợi:</strong> <span style="background:#fef3c7; padding:2px 6px; border-radius:4px;">${err.expected}</span>` : '';
        return `<li style="margin:8px 0; padding:10px; background:#fff5f5; border-left:4px solid #dc2626; border-radius:4px;">
                  <strong>${position}:</strong> ${msg}<br>
                  ${current ? current + '<br>' : ''}
                  ${expected}
                </li>`;
      } else {
        return `<li style="margin:8px 0;">${err}</li>`;
      }
    }).join('');
    return `<h4 style="color:#dc2626;">📌 Sheet: ${sheetName} (${result.errorCount} lỗi)</h4>
            <ul style="padding-left:20px;">${detailList}</ul>`;
  }).join('<hr>')}
  <hr>
  <p><strong>Vui lòng mở file Google Sheets để xem chi tiết đầy đủ và sửa lỗi (các ô lỗi đã được tô đỏ).</strong></p>
  <p>Trân trọng,<br><em>Script kiểm tra tự động</em></p>
`;
        var emailAddress = 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn';
        MailApp.sendEmail({
          to: emailAddress,
          subject: emailSubject,
          htmlBody: emailBody
        });
      } catch (e) {
        Logger.log('Lỗi gửi email: ' + e.message);
      }
    }

    if (erroredSheet) {
      spreadsheet.setActiveSheet(erroredSheet);
      erroredSheet.setActiveRange(erroredSheet.getRange(7, 1));
    }
  } catch (e) {
    Logger.log('Lỗi chính trong checkAllData: ' + e.message);
    ui.alert('Lỗi nghiêm trọng', 'Đã xảy ra lỗi: ' + e.message + '\nVui lòng kiểm tra log để biết chi tiết.', ui.ButtonSet.OK);
  }
}



function checkSheet7040GA(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*A$/) !== null;
    var isValidB = (colB.includes('7189-7040-30') && colB.match(/.*A$/)) ||
                   (colB.includes('7172-1333') && colB.match(/.*D$/)) ||
                   (colB.includes('7137-3079') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng A', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7040-30A, 7172-1333D hoặc 7137-3079A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}
 

function checkSheet7040GB(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*B$/) !== null;
    var isValidB = (colB.includes('7189-7040-30') && colB.match(/.*F$/)) ||
                    (colB.includes('7189-7040') && colB.match(/.*B$/)) ||
                   (colB.includes('7189-7040-40') && colB.match(/.*B$/)) ||
                   (colB.includes('7172-1333') && colB.match(/.*D$/)) ||
                   (colB.includes('7137-3079') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng B', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7040-30F, 7189-7040-40B, 7172-1333D hoặc 7137-3079A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheet8794B(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues(); // A, B, C
  var errors = [];
  var details = [];
  var valid = true;

  var lastLot = "";       
  var lastBEnd = "";      

  const FIXED_CODE_C = '7158-7116';  

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var rawA = (data[i][0] || '').toString().trim();
    var rawB = (data[i][1] || '').toString().trim();
    var rawC = (data[i][2] || '').toString().trim();

    // === KIỂM TRA MỚI: ĐUÔI PHẦN SAU $ CỦA CỘT A PHẢI LÀ 'B' ===
    var currentLot = '';
    var endA = '';  // chữ cuối phần lot

    if (rawA.includes('$')) {
      var partsA = rawA.split('$');
      currentLot = partsA[1].trim();
      if (currentLot.length > 0) {
        endA = currentLot.slice(-1).toUpperCase();
      }
    }

    // Báo lỗi nếu:
    // - Không có $ → NG
    // - Có $ nhưng phần sau rỗng → NG
    // - Đuôi không phải 'B' → NG
    if (!rawA.includes('$') || currentLot === '' || endA !== 'B') {
      errors.push({row: row, col: 1});
      details.push({
        row: row, 
        col: 1, 
        message: 'Phần sau $ của cột A phải tồn tại và kết thúc bằng B',
        current: rawA,
        expected: 'ví dụ: ...$...B'
      });
      valid = false;
    }

    // === CỘT B (giữ nguyên) ===
    if (rawB !== '') {
      var currentBEnd = '';
      if (rawB.includes('$')) {
        currentBEnd = rawB.split('$')[1].slice(-1).toUpperCase();
      } else {
        currentBEnd = rawB.slice(-1).toUpperCase();
      }

      if (currentBEnd !== 'A' && currentBEnd !== 'B' && currentBEnd !== 'C') {
        errors.push({row: row, col: 2});
        details.push({row: row, col: 2, message: 'Đuôi B không phải A/B/C', current: currentBEnd});
        valid = false;
      } else if (lastLot && currentLot === lastLot && lastBEnd && currentBEnd !== lastBEnd) {
        errors.push({row: row, col: 2});
        details.push({row: row, col: 2, message: 'Xung đột đuôi B trong cùng lot', current: currentBEnd, expected: lastBEnd});
        valid = false;
      }

      lastLot = currentLot;
      lastBEnd = currentBEnd;
    }

    // === CỘT C (giữ nguyên) ===
    if (rawC !== '') {
      var codeC = '';
      var endC = '';
      if (rawC.includes('$')) {
        var partsC = rawC.split('$');
        codeC = partsC[0].trim();
        endC = partsC[1].slice(-1).toUpperCase();
      } else {
        codeC = rawC;
        endC = rawC.slice(-1).toUpperCase();
      }

      if (!codeC.includes(FIXED_CODE_C) || (endC !== 'A' && endC !== 'B')) {
        errors.push({row: row, col: 3});
        details.push({row: row, col: 3, message: 'Mã C không hợp lệ hoặc đuôi không phải A/B', current: rawC});
        valid = false;
      }
    }

    // === QUY TẮC PHỤ THUỘC B → C (giữ nguyên) ===
    if (rawB !== '' && rawC !== '') {
      var bEnd = rawB.includes('$') ? rawB.split('$')[1].slice(-1).toUpperCase() : rawB.slice(-1).toUpperCase();
      var cEnd = rawC.includes('$') ? rawC.split('$')[1].slice(-1).toUpperCase() : rawC.slice(-1).toUpperCase();

      if ((bEnd === 'A' || bEnd === 'B') && cEnd === 'B') {
        errors.push({row: row, col: 3});
        details.push({row: row, col: 3, message: 'Đuôi C không được B khi đuôi B là A hoặc B', current: cEnd, expected: 'A'});
        valid = false;
      }
      if (bEnd === 'C' && cEnd === 'A') {
        errors.push({row: row, col: 3});
        details.push({row: row, col: 3, message: 'Đuôi C không được A khi đuôi B là C', current: cEnd, expected: 'B'});
        valid = false;
      }
    }
  }

  // Tô đỏ tất cả lỗi (bao gồm cột A nếu có)
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return { 
    valid: valid, 
    errorCount: errors.length, 
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details 
  };
}
function checkSheet7040GE(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*E$/) !== null;
    var isValidB = (colB.includes('7189-7040-30') && colB.match(/.*E$/)) ||
                   (colB.includes('7172-1333') && colB.match(/.*C$/)) ||
                   (colB.includes('7137-3079') && colB.match(/.*C$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng E', current: colA, expected: 'Kết thúc bằng E'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7040-30E, 7172-1333C hoặc 7137-3079C'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS001A(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*A$/) !== null;
    var isValidB = (colB.includes('7186-1896') && colB.match(/.*A$/)) ||
                   (colB.includes('7172-0221') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng A', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7186-1896A hoặc 7172-0221A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS001B(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*B$/) !== null;
    var isValidB = (colB.includes('7186-1896') && colB.match(/.*B$/)) ||
                   (colB.includes('7172-0221') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng B', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7186-1896B hoặc 7172-0221A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS001C(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*C$/) !== null;
    var isValidB = (colB.includes('7186-1896') && colB.match(/.*C$/)) ||
                   (colB.includes('7172-0221') && colB.match(/.*B$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng C', current: colA, expected: 'Kết thúc bằng C'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7186-1896C hoặc 7172-0221B'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVA019A(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;

  // Thêm: Theo dõi lot thực tế (sau $) → chữ cuối B và C đầu tiên gặp + danh sách row
  var lotRealMap = {}; // key: lotReal (ví dụ "251209A"), value: { bEnd: 'A', cEnd: 'B', rows: [...] }

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';

    if (colA === "" && colB === "" && colC === "") continue;

    // === LOGIC GỐC CỦA VA019A - GIỮ NGUYÊN HOÀN TOÀN KHÔNG THAY ĐỔI ===
    var isValidA = colA !== '' && colA.match(/.*A$/) !== null;

    var isValidB = colB === '' ||
                   (colB.includes('7189-3342-40') && colB.match(/.*A$/)) ||
                   (colB.includes('7189-3342-90') && colB.match(/.*A$/)) ||
                   (colB.includes('7189-3342-30') && colB.match(/.*C$/));

    var isValidC = colC === '' ||
                   (colC.includes('7158-6425') && (colC.match(/.*A$/) || colC.match(/.*B$/)));

    // Quy tắc tổ hợp gốc
    var isValidRow = false;
    if (colA !== '' && colB !== '' && colC !== '') {
      var colCEnd = colC.slice(-1).toUpperCase();
      isValidRow =
        (colA.includes('7289-3342-40') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-40') && colB.match(/.*A$/) &&
         colC.includes('7158-6425') && colCEnd === 'B') ||
        (colA.includes('7289-3342-90') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-90') && colB.match(/.*A$/) &&
         colC.includes('7158-6425') && colCEnd === 'B') ||
        (colA.includes('7289-3342-30') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-30') && colB.match(/.*C$/) &&
         colC.includes('7158-6425') && colCEnd === 'B') ||
        (colA.includes('7289-3342-30') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-30') && colB.match(/.*C$/) &&
         colC.includes('7158-6425') && colCEnd === 'A');
    }

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7189-3342-40A, 7189-3342-90A hoặc 7189-3342-30C'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425 với đuôi A hoặc B'});
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({row: row, col: null, message: 'Tổ hợp A-B-C không hợp lệ'});
      valid = false;
    }

        // === QUY TẮC MỚI: KIỂM SOÁT CHỮ CUỐI CỦA CẢ CỘT B VÀ CỘT C KHI CÙNG LOT SAU $ ===
    if (colA && colA.includes('$') && colA.match(/.*A$/) !== null) {
      var parts = colA.split('$');
      var lotReal = parts[parts.length - 1].trim(); // lot thực tế sau $ cuối cùng

      var currentBEnd = colB ? colB.slice(-1).toUpperCase() : '';
      var currentCEnd = colC ? colC.slice(-1).toUpperCase() : '';

      if (lotRealMap[lotReal] === undefined) {
        // Lần đầu gặp lot thực tế này → ghi nhớ chữ cuối B và C + lưu dòng đầu tiên
        lotRealMap[lotReal] = {
          bEnd: currentBEnd,
          cEnd: currentCEnd,
          rows: [row]  // vẫn lưu danh sách rows để tiện sau này nếu cần
        };
      } else {
        // Đã gặp trước đó
        lotRealMap[lotReal].rows.push(row);

        // Kiểm tra xung đột (chỉ khi ô không trống)
        var conflictB = (currentBEnd !== '' && lotRealMap[lotReal].bEnd !== '' && lotRealMap[lotReal].bEnd !== currentBEnd);
        var conflictC = (currentCEnd !== '' && lotRealMap[lotReal].cEnd !== '' && lotRealMap[lotReal].cEnd !== currentCEnd);

        if (conflictB || conflictC) {
          // CHỈ TÔ ĐỎ DÒNG HIỆN TẠI (dòng gây xung đột)
          errors.push({row: row, col: 1});
          errors.push({row: row, col: 2});
          errors.push({row: row, col: 3});

          var conflictMsg = 'Xung đột đuôi trong cùng lot ' + lotReal;
          if (conflictB) conflictMsg += ' (B: ' + currentBEnd + ' ≠ ' + lotRealMap[lotReal].bEnd + ')';
          if (conflictC) conflictMsg += ' (C: ' + currentCEnd + ' ≠ ' + lotRealMap[lotReal].cEnd + ')';

          details.push({row: row, col: null, message: conflictMsg});

          valid = false;

          // Tùy chọn: cũng có thể đánh dấu dòng đầu tiên của lot để người dùng biết "chuẩn là gì"
          // Nhưng theo yêu cầu "chỉ bôi đỏ dòng lỗi" → không tô dòng đầu
        }
      }
    
    }
  }

  // Tô đỏ lỗi
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return { 
    valid: valid, 
    errorCount: errors.length, 
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details 
  };
}

function checkSheetVA019B(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  var lotRealMap = {};

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';

    if (colA === "" && colB === "" && colC === "") continue;

    var colBEnd = colB ? String(colB).slice(-1).toUpperCase() : '';
    var colCEnd = colC ? String(colC).slice(-1).toUpperCase() : '';

    var prefixA = String(colA).includes('$') ? String(colA).split('$')[0].trim() : String(colA).trim();
    var prefixB = String(colB).includes('$') ? String(colB).split('$')[0].trim() : String(colB).trim();
    var codeA = prefixA.replace(/^7289-/, '');
    var codeB = prefixB.replace(/^7189-/, '');
    var isSamePrefix = (codeA === codeB);
    var is90 = (codeA === '3342-90');

    var isValidA = /B$/i.test(colA);
    var isValidB = String(colB).includes('7189-3342') && (colBEnd === 'B' || colBEnd === 'C' || colBEnd === 'F');
    var isValidC = String(colC).includes('7158-6425') && (colCEnd === 'C' || colCEnd === 'E');

    // Cặp đuôi hợp lệ: B-C, C-C, F-E  |  F-C / B-E / C-E = NG
    var pairOK =
      (colBEnd === 'B' && colCEnd === 'C') ||
      (colBEnd === 'C' && colCEnd === 'C') ||
      (colBEnd === 'F' && colCEnd === 'E');

    // -90 chỉ được F-E
    var combo90OK = !is90 || (colBEnd === 'F' && colCEnd === 'E');

    var isValidRow = isSamePrefix && isValidA && pairOK && combo90OK;

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7189-3342 đuôi B / C / F'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425 đuôi C hoặc E'});
      valid = false;
    }
    if (!isSamePrefix) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      details.push({
        row: row, col: null,
        message: 'Phần trước $ của A và B không khớp',
        current: prefixA + ' | ' + prefixB,
        expected: '3342 với 3342, hoặc 3342-90 với 3342-90'
      });
      valid = false;
    }
    if (colBEnd === 'F' && colCEnd === 'C') {
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({
        row: row, col: null,
        message: 'GF + GC không hợp lệ',
        current: 'B=' + colBEnd + ' / C=' + colCEnd,
        expected: 'GF phải đi với GE (F + E)'
      });
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({
        row: row, col: null,
        message: 'Tổ hợp A-B-C không hợp lệ',
        current: prefixA + ' | B=' + colBEnd + ' | C=' + colCEnd,
        expected: is90 ? '3342-90 chỉ được F + E' : '3342: B+C hoặc C+C hoặc F+E (không F+C)'
      });
      valid = false;
    }

    if (colA && String(colA).includes('$') && isValidA) {
      var lotReal = String(colA).split('$').pop().trim();
      if (lotRealMap[lotReal] === undefined) {
        lotRealMap[lotReal] = { bEnd: colBEnd, cEnd: colCEnd, rows: [row] };
      } else {
        lotRealMap[lotReal].rows.push(row);
        var conflictB = (colBEnd !== '' && lotRealMap[lotReal].bEnd !== '' && lotRealMap[lotReal].bEnd !== colBEnd);
        var conflictC = (colCEnd !== '' && lotRealMap[lotReal].cEnd !== '' && lotRealMap[lotReal].cEnd !== colCEnd);
        if (conflictB || conflictC) {
          errors.push({row: row, col: 1});
          errors.push({row: row, col: 2});
          errors.push({row: row, col: 3});
          var conflictMsg = 'Xung đột đuôi trong cùng lot ' + lotReal;
          if (conflictB) conflictMsg += ' (B: ' + colBEnd + ' ≠ ' + lotRealMap[lotReal].bEnd + ')';
          if (conflictC) conflictMsg += ' (C: ' + colCEnd + ' ≠ ' + lotRealMap[lotReal].cEnd + ')';
          details.push({row: row, col: null, message: conflictMsg});
          valid = false;
        }
      }
    }
  }

  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return {
    valid: valid,
    errorCount: errors.length,
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details
  };
}
function checkSheetVA019C(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;

  // Thêm: Theo dõi lot thực tế (sau $) → chữ cuối B và C đầu tiên gặp + danh sách row
  var lotRealMap = {}; // key: lotReal (ví dụ "251209C"), value: { bEnd: 'D', cEnd: 'A', rows: [...] }

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';

    if (colA === "" && colB === "" && colC === "") continue;

    // === LOGIC GỐC CỦA VA019C - GIỮ NGUYÊN HOÀN TOÀN KHÔNG THAY ĐỔI ===
    var isValidA = colA.match(/.*C$/) !== null;
    var isValidB = (colB.includes('7189-3342-30') && colB.match(/.*D$/)) ||
                   (colB.includes('7189-3342-30') && colB.match(/.*E$/));
    var isValidC = (colC.includes('7158-6425') && colC.match(/.*A$/)) ||
                   (colC.includes('7158-6425') && colC.match(/.*D$/));

    var colCEnd = colC ? colC.slice(-1).toUpperCase() : '';

    var isValidRow = (colA.includes('7289-3342-30') && colA.match(/.*C$/) &&
                      colB.includes('7189-3342-30') && colB.match(/.*D$/) &&
                      colC.includes('7158-6425') && colCEnd === 'D') ||
                     (colA.includes('7289-3342-30') && colA.match(/.*C$/) &&
                      colB.includes('7189-3342-30') && colB.match(/.*D$/) &&
                      colC.includes('7158-6425') && colCEnd === 'A') ||
                     (colA.includes('7289-3342-30') && colA.match(/.*C$/) &&
                      colB.includes('7189-3342-30') && colB.match(/.*E$/) &&
                      colC.includes('7158-6425') && colCEnd === 'D');

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: 'Kết thúc bằng C'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7189-3342-30D hoặc 7189-3342-30E'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425 với đuôi A hoặc D'});
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({row: row, col: null, message: 'Tổ hợp A-B-C không hợp lệ'});
      valid = false;
    }

       // === QUY TẮC MỚI: KIỂM SOÁT CHỮ CUỐI CỦA CẢ CỘT B VÀ CỘT C KHI CÙNG LOT SAU $ ===
    if (colA && colA.includes('$') && colA.match(/.*C$/) !== null) {
      var parts = colA.split('$');
      var lotReal = parts[parts.length - 1].trim();

      var currentBEnd = colB ? colB.slice(-1).toUpperCase() : '';
      var currentCEnd = colC ? colC.slice(-1).toUpperCase() : '';

      if (lotRealMap[lotReal] === undefined) {
        lotRealMap[lotReal] = {
          bEnd: currentBEnd,
          cEnd: currentCEnd,
          rows: [row]
        };
      } else {
        lotRealMap[lotReal].rows.push(row);

        var conflictB = (currentBEnd !== '' && lotRealMap[lotReal].bEnd !== '' && lotRealMap[lotReal].bEnd !== currentBEnd);
        var conflictC = (currentCEnd !== '' && lotRealMap[lotReal].cEnd !== '' && lotRealMap[lotReal].cEnd !== currentCEnd);

        if (conflictB || conflictC) {
          errors.push({row: row, col: 1});
          errors.push({row: row, col: 2});
          errors.push({row: row, col: 3});

          var conflictMsg = 'Xung đột đuôi trong cùng lot ' + lotReal;
          if (conflictB) conflictMsg += ' (B: ' + currentBEnd + ' ≠ ' + lotRealMap[lotReal].bEnd + ')';
          if (conflictC) conflictMsg += ' (C: ' + currentCEnd + ' ≠ ' + lotRealMap[lotReal].cEnd + ')';

          details.push({row: row, col: null, message: conflictMsg});

          valid = false;
        }
      }
    
    }
  }

  // Tô đỏ lỗi
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return { 
    valid: valid, 
    errorCount: errors.length, 
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details 
  };
}

function checkSheetVA003(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';
    if (colA === "" && colB === "" && colC === "") continue;
    var isValidA = (colA.includes('7283-4578') && colA.match(/.*A$/)) ||
                   (colA.includes('7298-8296') && colA.match(/.*A$/));
    var isValidB = (colB.includes('7183-4578') && colB.match(/.*A$/)) ||
                   (colB.includes('7198-8296') && colB.match(/.*A$/));
    var isValidC = colC.includes('7158-6425') && colC.match(/.*A$/);
    var isValidRow = (colA.includes('7283-4578') && colA.match(/.*A$/) &&
                      colB.includes('7183-4578') && colB.match(/.*A$/) &&
                      colC.includes('7158-6425') && colC.match(/.*A$/)) ||
                     (colA.includes('7298-8296') && colA.match(/.*A$/) &&
                      colB.includes('7198-8296') && colB.match(/.*A$/) &&
                      colC.includes('7158-6425') && colC.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: '7283-4578A hoặc 7298-8296A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7183-4578A hoặc 7198-8296A'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425A'});
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({row: row, col: null, message: 'Tổ hợp A-B-C không hợp lệ'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS051(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var colAEnd = colA.slice(-1).toUpperCase();
    var colBEnd = colB.slice(-1).toUpperCase();
    var isValidRow = (colA.includes('7286-7859') && colAEnd === 'A' && colB.includes('7186-7859') && colBEnd === 'A') ||
                     (colA.includes('7286-7859') && colAEnd === 'B' && colB.includes('7186-7859') && colBEnd === 'B');
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      details.push({row: row, col: null, message: 'Tổ hợp A-B không hợp lệ', current: colA + ' & ' + colB, expected: '7286-7859A & 7186-7859A hoặc 7286-7859B & 7186-7859B'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheet7041GA(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;

    // Cột A phải kết thúc bằng GA 
    var isValidA = colA.match(/.*A$/i) !== null;

    // Cột B chỉ được 3 mã sau và đúng đuôi GA 
    var isValidB = (colB.includes('7189-7041-30') && colB.match(/.*A$/i)) ||
                   (colB.includes('7172-1334') && colB.match(/.*A$/i)) ||
                   (colB.includes('7137-3080') && colB.match(/.*A$/i));

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng A', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7041-30A, 7172-1334A hoặc 7137-3080A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}
function checkSheet7041GB(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;

    // Cột A phải kết thúc bằng B
    var isValidA = colA.match(/.*B$/i) !== null;

    // Cột B chỉ được 3 mã sau và đúng đuôi 
    var isValidB = (colB.includes('7189-7041-30') && colB.match(/.*B$/i)) ||
                   (colB.includes('7172-1334') && colB.match(/.*A$/i)) ||
                   (colB.includes('7137-3080') && colB.match(/.*A$/i));

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng B', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7041-30B, 7172-1334A hoặc 7137-3080A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('checkAllData')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// HỆ THỐNG: KIỂM TRA MÃ KHUÔN & TEM NHÃN – BẢN HOÀN HẢO CUỐI CÙNG
// LB-VA019A: CỘT A CỐ ĐỊNH LB-VA019A-7289-3342-A | CỘT B CHỈ 3 MÃ | CỘT C PHẢI A
// ═══════════════════════════════════════════════════════════════════════
function checkLabelQuick() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const RULES = [
    { sheet: 'LB-VA019A', fixedA: 'LB-VA019A-7289-3342-A', codesB: ['7289-3342-30', '7289-3342-40', '7289-3342-90'], endC: 'A' },
    { sheet: 'LB-VA019B', fixedA: 'LB-VA019B-7289-3342-B', codesB: '7289-3342', endC: 'B' },
    { sheet: 'LB-VA019C', fixedA: 'LB-VA019C-7289-3342-C', codesB: '7289-3342-30', endC: 'C' },
    { sheet: 'LB-VA015A', fixedA: 'LB-VA015A-7289-7040-30-A', codesB: '7289-7040-30', endC: 'A' },
    { sheet: 'LB-VA015B', fixedA: 'LB-VA015B-7289-7040-30-B', codesB: ['7289-7040', '7289-7040-30', '7289-7040-40'], endC: 'B' },
    { sheet: 'LB-VA015E', fixedA: 'LB-VA015E-7289-7040-30-E', codesB: '7289-7040-30', endC: 'E' },
    { sheet: 'LB-VA016A', fixedA: 'LB-VA016A-7289-7041-30-A', codesB: '7289-7041-30', endC: 'A' },
    { sheet: 'LB-VA016B', fixedA: 'LB-VA016B-7289-7041-30-B', codesB: '7289-7041-30', endC: 'B' },
    { sheet: 'LB-VA001B', fixedA: 'LB-VA001B-7283-8794-70-B', codesB: '7283-8794-70', endC: 'B' },
    { sheet: 'LB-VA006', fixedA: 'LB-VA006-7283-0651-70-A', codesB: '7283-0651-70', endC: 'A' },
    { sheet: 'LB-VA018', fixedA: 'LB-VA018-7289-8063-30-A', codesB: '7289-8063-30', endC: 'B' },
    { sheet: 'LB-VS001A', fixedA: 'LB-VS001A-7286-1896-A', codesB: '7286-1896', endC: 'A' },
    { sheet: 'LB-VS001B', fixedA: 'LB-VS001B-7286-1896-B', codesB: '7286-1896', endC: 'B' },
    { sheet: 'LB-VS001C', fixedA: 'LB-VS001C-7286-1896-C', codesB: '7286-1896', endC: 'C' },
    { sheet: 'LB-VS002', fixedA: 'LB-VS002-7286-1902-A', codesB: '7286-1902', endC: 'A' },
    { sheet: 'LB-VS004', fixedA: 'LB-VS004-7286-1903-A', codesB: '7286-1903', endC: 'A' },
    { sheet: 'LB-VS012C', fixedA: 'LB-VS012C-7287-8821-C', codesB: '7287-8821', endC: 'C' },
    { sheet: 'LB-VS051B', fixedA: 'LB-VS051B-7286-7859-B', codesB: '7286-7859', endC: 'B' }
  ];

  let totalErrors = 0;
  const errorDetails = [];       // Chi tiết lỗi để gửi email
  const sheetStatus = [];        // Chỉ để hiển thị OK/LỖI ngắn gọn trong log cũ (nếu cần)

  RULES.forEach(r => {
    const sheet = ss.getSheetByName(r.sheet);
    if (!sheet) {
      errorDetails.push(`<li><strong>Không tìm thấy sheet:</strong> ${r.sheet}</li>`);
      sheetStatus.push(`Không tìm thấy sheet: ${r.sheet}`);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 7) {
      errorDetails.push(`<li><strong>Không có dữ liệu:</strong> ${r.sheet}</li>`);
      sheetStatus.push(`Không có dữ liệu: ${r.sheet}`);
      return;
    }

    const startRow = 7;
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).setBackground(null);

    let sheetHasError = false;

    data.forEach((row, i) => {
      const rowNum = startRow + i;
      const a = (row[0] || '').toString().trim();
      const b = (row[1] || '').toString().trim();
      const c = (row[2] || '').toString().trim();

      // Kiểm tra cột A
      if (a !== r.fixedA) {
        sheet.getRange(rowNum, 1).setBackground('#ff5252');
        errorDetails.push(`<li><strong>${r.sheet} - Dòng ${rowNum}</strong>: Cột A sai → thực tế "<em>${a}</em>", mong đợi "<em>${r.fixedA}</em>"</li>`);
        sheetHasError = true;
        totalErrors++;
      }

      // Kiểm tra cột B
      let bValid = false;
      if (Array.isArray(r.codesB)) {
        bValid = r.codesB.includes(b);
      } else {
        bValid = (b === r.codesB);
      }

      if (b && !bValid) {
        sheet.getRange(rowNum, 2).setBackground('#ff5252');
        const expectedB = Array.isArray(r.codesB) ? r.codesB.join(', ') : r.codesB;
        errorDetails.push(`<li><strong>${r.sheet} - Dòng ${rowNum}</strong>: Cột B sai → thực tế "<em>${b}</em>", mong đợi một trong "<em>${expectedB}</em>"</li>`);
        sheetHasError = true;
        totalErrors++;
      }

      // Kiểm tra cột C
      if (c && !c.toUpperCase().endsWith(r.endC)) {
        sheet.getRange(rowNum, 3).setBackground('#ff5252');
        errorDetails.push(`<li><strong>${r.sheet} - Dòng ${rowNum}</strong>: Cột C sai → thực tế "<em>${c}</em>", phải kết thúc bằng "<em>${r.endC}</em>"</li>`);
        sheetHasError = true;
        totalErrors++;
      }
    });

    const status = sheetHasError ? `LỖI: ${r.sheet}` : `OK: ${r.sheet}`;
    sheetStatus.push(status);
  });

  const timeStr = Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy');

  // ================== HIỂN THỊ KẾT QUẢ ==================
  if (totalErrors === 0) {
    ui.showModalDialog(HtmlService.createHtmlOutput(`
      <div style="text-align:center;padding:80px;background:#e8f5e8;font-family:Arial;">
        <h1 style="color:#2e7d32">HOÀN HẢO!</h1>
        <div style="font-size:120px">✅</div>
        <h2>Tất cả dữ liệu đều chuẩn</h2>
        <button onclick="google.script.host.close()"
          style="padding:18px 60px;font-size:22px;background:#4caf50;color:white;border:none;border-radius:10px">
          OK
        </button>
      </div>
    `).setWidth(520).setHeight(420), 'KIỂM TRA');

    return;
  }

  // ================== GỬI EMAIL CHI TIẾT KHI CÓ LỖI ==================
  MailApp.sendEmail({
    to: 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn',
    subject: `LỖI MÃ KHUÔN & TEM NHÃN – ${totalErrors} lỗi – ${timeStr}`,
    htmlBody: `
      <h2 style="color:#c62828;">Phát hiện ${totalErrors} lỗi trong file kiểm tra nhãn</h2>
      <p><strong>Thời gian kiểm tra:</strong> ${timeStr}</p>
      <hr>
      <h3>Chi tiết lỗi:</h3>
      <ol>${errorDetails.join('')}</ol>
      <hr>
      <p>Vui lòng mở file Google Sheets và sửa các ô được tô đỏ.</p>
      <p>Trân trọng,<br>Script tự động kiểm tra</p>
    `
  });

  ui.showModalDialog(HtmlService.createHtmlOutput(`
    <div style="text-align:center;padding:80px;background:#ffebee;font-family:Arial;">
      <h1 style="color:#c62828">Có ${totalErrors} Lỗi<br>đã gửi thông tin chi tiết qua email...</h1>
      <div style="font-size:120px">❌</div>
      <button onclick="google.script.run.jumpToError()"
        style="padding:18px 60px;font-size:22px;background:#e53935;color:white;border:none;border-radius:10px">
        Xác nhận ngay !!
      </button>
    </div>
  `).setWidth(520).setHeight(420), 'CẢNH BÁO');
}
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KIỂM TRA MÃ KHUÔN & TEM NHÃN')
    .addItem('Kiểm tra toàn bộ từ dòng 7', 'checkLabelQuick')
    .addToUi();
}

// ==================== RESET HOÀN HẢO – KHÔNG MẤT DÒNG 7 NỮA ====================
function resetAllSheetsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'XÁC NHẬN RESET',
    'Xóa toàn bộ dữ liệu từ dòng 7 trở xuống trong TẤT CẢ các sheet?\n\n' +
    'Dòng dữ liệu mới nhất sẽ được kéo lên đúng dòng 7.\nKHÔNG THỂ HOÀN TÁC!',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Đã hủy', 'Không thực hiện reset.', ui.ButtonSet.OK);
    return;
  }

  ui.showModalDialog(HtmlService.createHtmlOutput(`
    <div style="text-align:center;padding:80px;background:#fff3e0;font-family:'Roboto',sans-serif;">
      <h2 style="font-size:32px;color:#ef6c00;">Đang reset dữ liệu...</h2>
      <p style="font-size:20px;color:#e65100;">Đang xử lý toàn bộ sheet</p>
      <div style="font-size:90px;">🗑️➡️📋</div>
    </div>
  `).setWidth(520).setHeight(420), 'ĐANG XỬ LÝ');

  const sheets = ss.getSheets();
  let processedCount = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === 'CheckLog') return; // Giữ lịch sử

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    // === 1. Xóa sạch nội dung + màu nền từ dòng 8 trở xuống (GIỮ NGUYÊN DÒNG 7) ===
    const maxRows = sheet.getMaxRows();
    if (maxRows > 7) {
      sheet.getRange(8, 1, maxRows - 7, lastCol)
           .clearContent()
           .setBackground(null);
    }

    // === 2. Tìm khối dữ liệu cuối cùng (dòng mới nhất) ===
    const lastRowAfterClear = sheet.getLastRow();
    if (lastRowAfterClear < 8) {
      processedCount++;
      return; // Không có dữ liệu cần di chuyển
    }

    const values = sheet.getRange(8, 1, lastRowAfterClear - 7, lastCol).getValues();

    // Tìm dòng cuối cùng có dữ liệu
    let lastDataIndex = -1;
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i].some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        lastDataIndex = i;
        break;
      }
    }
    if (lastDataIndex === -1) {
      processedCount++;
      return;
    }

    // Tìm đầu khối dữ liệu liên tục
    let firstDataIndex = lastDataIndex;
    for (let i = lastDataIndex - 1; i >= 0; i--) {
      const allEmpty = values[i].every(cell => cell === '' || cell === null || cell === undefined);
      if (allEmpty) break;
      firstDataIndex = i;
    }

    const blockHeight = lastDataIndex - firstDataIndex + 1;
    const sourceRow = 8 + firstDataIndex; // Dòng thực tế trong sheet

    // === 3. Di chuyển lên dòng 7 ===
    const sourceRange = sheet.getRange(sourceRow, 1, blockHeight, lastCol);
    const targetRange = sheet.getRange(7, 1, blockHeight, lastCol);
    sourceRange.moveTo(targetRange);

    // === 4. Xóa vùng cũ (an toàn tuyệt đối) ===
    if (sourceRow + blockHeight - 1 <= maxRows) {
      sheet.deleteRows(sourceRow, blockHeight);
    }

    processedCount++;
  });

  ui.showModalDialog(HtmlService.createHtmlOutput(`
    <div style="text-align:center;padding:80px;background:#e8f5e8;font-family:'Roboto',sans-serif;">
      <h1 style="font-size:48px;color:#2e7d32;">HOÀN TẤT!</h1>
      <p style="font-size:24px;color:#1b5e20;">
        Đã reset thành công!<br>
        <b>Dòng mới nhất đã được kéo lên dòng 7</b><br>
        Xử lý ${processedCount} sheet.
      </p>
      <div style="font-size:100px;">✅</div>
      <button onclick="google.script.host.close()" style="margin-top:30px;padding:16px 50px;font-size:22px;background:#4caf50;color:white;border:none;border-radius:12px;cursor:pointer;">
        Đóng
      </button>
    </div>
  `).setWidth(560).setHeight(520), 'RESET THÀNH CÔNG');
}
// ==================== DANH SÁCH EMAIL ĐƯỢC PHÉP DÙNG TOOL VÀ CHỈNH SỬA CODE ====================
const ALLOWED_EMAILS = [
  'sitien353@gmail.com.vn',
  // Thêm email khác nếu cần
];

// ==================== ẨN MENU VỚI NGƯỜI KHÔNG ĐƯỢC PHÉP ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail();

  if (ALLOWED_EMAILS.includes(userEmail) || userEmail === '') {
    ui.createMenu('🔧 TOOL KIỂM TRA MÃ KHUÔN')
      .addItem('🔍 Kiểm tra cấu tạo mã khuôn linh kiện', 'checkAllData')
      .addItem('🏷️ Kiểm tra mã khuôn & tem nhãn (dòng 7)', 'checkLabelQuick')
      .addSeparator()
      .addItem('♻️ Reset dữ liệu (kéo dòng mới lên dòng 7)', 'resetAllSheetsData')
      .addToUi();
  }
}

// ==================== CHẶN MỞ APPS SCRIPT EDITOR ====================
function onOpenScriptEditor() {
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!ALLOWED_EMAILS.includes(userEmail) && userEmail !== '') {
    throw new Error('🚫 Bạn không có quyền truy cập hoặc chỉnh sửa script này.\nLiên hệ admin để được hỗ trợ.');
  }
}

//update speaker 28/08/2026 + hiển thị Số xe trong lịch sử
function getSheetCheckMap() {
  return {
    'VA001-B': checkSheet8794B,
    'VA015-A': checkSheet7040GA,
    'VA015-B': checkSheet7040GB,
    'VA015-E': checkSheet7040GE,
    'VS001-A': checkSheetVS001A,
    'VS001-B': checkSheetVS001B,
    'VS001-C': checkSheetVS001C,
    'VA019-A': checkSheetVA019A,
    'VA019-B': checkSheetVA019B,
    'VA019-C': checkSheetVA019C,
    'VA003': checkSheetVA003,
    'VS051': checkSheetVS051,
    'VA016-A': checkSheet7041GA,
    'VA016-B': checkSheet7041GB
  };
}

function getLabelRules() {
  return [
    { sheet: 'LB-VA019A', fixedA: 'LB-VA019A-7289-3342-A', codesB: ['7289-3342-30', '7289-3342-40', '7289-3342-90'], endC: 'A' },
    { sheet: 'LB-VA019B', fixedA: 'LB-VA019B-7289-3342-B', codesB: '7289-3342', endC: 'B' },
    { sheet: 'LB-VA019C', fixedA: 'LB-VA019C-7289-3342-C', codesB: '7289-3342-30', endC: 'C' },
    { sheet: 'LB-VA015A', fixedA: 'LB-VA015A-7289-7040-30-A', codesB: '7289-7040-30', endC: 'A' },
    { sheet: 'LB-VA015B', fixedA: 'LB-VA015B-7289-7040-30-B', codesB: ['7289-7040', '7289-7040-30', '7289-7040-40'], endC: 'B' },
    { sheet: 'LB-VA015E', fixedA: 'LB-VA015E-7289-7040-30-E', codesB: '7289-7040-30', endC: 'E' },
    { sheet: 'LB-VA016A', fixedA: 'LB-VA016A-7289-7041-30-A', codesB: '7289-7041-30', endC: 'A' },
    { sheet: 'LB-VA016B', fixedA: 'LB-VA016B-7289-7041-30-B', codesB: '7289-7041-30', endC: 'B' },
    { sheet: 'LB-VA001B', fixedA: 'LB-VA001B-7283-8794-70-B', codesB: '7283-8794-70', endC: 'B' },
    { sheet: 'LB-VA006',  fixedA: 'LB-VA006-7283-0651-70-A', codesB: '7283-0651-70', endC: 'A' },
    { sheet: 'LB-VA018',  fixedA: 'LB-VA018-7289-8063-30-A', codesB: '7289-8063-30', endC: 'B' },
    { sheet: 'LB-VS001A', fixedA: 'LB-VS001A-7286-1896-A', codesB: '7286-1896', endC: 'A' },
    { sheet: 'LB-VS001B', fixedA: 'LB-VS001B-7286-1896-B', codesB: '7286-1896', endC: 'B' },
    { sheet: 'LB-VS001C', fixedA: 'LB-VS001C-7286-1896-C', codesB: '7286-1896', endC: 'C' },
    { sheet: 'LB-VS002',  fixedA: 'LB-VS002-7286-1902-A', codesB: '7286-1902', endC: 'A' },
    { sheet: 'LB-VS004',  fixedA: 'LB-VS004-7286-1903-A', codesB: '7286-1903', endC: 'A' },
    { sheet: 'LB-VS012C', fixedA: 'LB-VS012C-7287-8821-C', codesB: '7287-8821', endC: 'C' },
    { sheet: 'LB-VS051B', fixedA: 'LB-VS051B-7286-7859-B', codesB: '7286-7859', endC: 'B' }
  ];
}

function listSheetNames() { return Object.keys(getSheetCheckMap()); }
function listLabelSheetNames() { return getLabelRules().map(function(r){ return r.sheet; }); }

function checkLabelSheetByRule(sheet, rule) {
  var last = sheet.getLastRow();
  if (last < 7) return { valid: false };
  var allowB = Array.isArray(rule.codesB) ? rule.codesB : [rule.codesB];
  var rows = sheet.getRange(7, 1, last - 6, 3).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    var a = String(rows[i][0] || '').trim();
    var b = String(rows[i][1] || '').trim();
    var c = String(rows[i][2] || '').trim();
    if (!a && !b && !c) continue;
    if (a !== rule.fixedA) return { valid: false };
    if (allowB.indexOf(b) < 0) return { valid: false };
    if (rule.endC && (!c || c.slice(-1) !== rule.endC)) return { valid: false };
  }
  return { valid: true };
}

function getLatestCodeFromSheet(sheet) {
  var last = sheet.getLastRow();
  if (last < 7) return '';
  var raw = String(sheet.getRange(last, 1).getDisplayValue() || '').trim();
  var parts = raw.split('-');
  return parts.length >= 2 ? String(parts[1]).trim() : raw;
}

// [MỚI] Đọc "Số xe" mới nhất của sheet (chỉ đọc, không ảnh hưởng logic kiểm tra).
// Tự tìm cột có tiêu đề chứa "số xe" trong 6 dòng đầu; không có thì trả về ''.
function getXeValue_(sheet) {
  try {
    var lastCol = sheet.getLastColumn(), last = sheet.getLastRow();
    if (lastCol < 1 || last < 7) return '';
    var norm = function (s) {
      return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
    };
    var head = sheet.getRange(1, 1, 6, lastCol).getDisplayValues();
    var col = -1;
    for (var r = 0; r < head.length && col < 0; r++) {
      for (var c = 0; c < head[r].length; c++) {
        if (norm(head[r][c]).indexOf('so xe') >= 0) { col = c + 1; break; }
      }
    }
    if (col < 0) return '';
    var vals = sheet.getRange(7, col, last - 6, 1).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var v = String(vals[i][0] || '').trim();
      if (v) return v;
    }
  } catch (e) {}
  return '';
}

// [MỚI] Tự tìm cột "Tên" (theo tiêu đề trong 6 dòng đầu) và lấy giá trị mới nhất.
// Ưu tiên: "tên" > "họ tên" > "tên nhân viên" > "nhân viên" > "người"; cuối cùng là tiêu đề có từ "tên".
function getTenValue_(sheet) {
  try {
    var lastCol = sheet.getLastColumn(), last = sheet.getLastRow();
    if (lastCol < 1 || last < 7) return '';
    var norm = function (s) {
      return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
    };
    var head = sheet.getRange(1, 1, 6, lastCol).getDisplayValues();
    var tiers = [
      function (h) { return h === 'ten'; },
      function (h) { return h === 'ho ten' || h === 'ho va ten'; },
      function (h) { return h === 'ten nhan vien' || h === 'ten nv'; },
      function (h) { return h === 'nhan vien' || h === 'nv'; },
      function (h) { return h === 'nguoi' || h === 'nguoi scan' || h === 'nguoi thuc hien'; },
      function (h) { return (' ' + h + ' ').indexOf(' ten ') >= 0; }
    ];
    var col = -1;
    for (var t = 0; t < tiers.length && col < 0; t++) {
      for (var r = 0; r < head.length && col < 0; r++) {
        for (var c = 0; c < head[r].length; c++) {
          if (tiers[t](norm(head[r][c]))) { col = c + 1; break; }
        }
      }
    }
    if (col < 0) return '';
    var vals = sheet.getRange(7, col, last - 6, 1).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var v = String(vals[i][0] || '').trim();
      if (v) return v;
    }
  } catch (e) {}
  return '';
}

function ensureHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('IoT_History');
  if (!sheet) {
    sheet = ss.insertSheet('IoT_History');
    sheet.appendRow(['Thời gian', 'Loại', 'Line', 'Trạng thái', 'Nội dung', 'Số xe', 'Tên']);
  } else {
    if (!sheet.getRange(1, 6).getValue()) sheet.getRange(1, 6).setValue('Số xe');
    if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('Tên');
  }
  return sheet;
}

function writeHistory(kind, sheetName, status, text, xe, ten) {
  var sheet = ensureHistorySheet_();
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 7).setValues([[
    Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy'),
    kind || '',
    sheetName || '',
    status || '',
    text || '',
    xe || '',
    ten || ''
  ]]);
}

function writeResultForESP32(announceText, hasErrors, speak) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('IoT_Result');
  if (!sheet) {
    sheet = ss.insertSheet('IoT_Result');
    sheet.appendRow(['Thời gian', 'Trạng thái', 'Câu thông báo', 'Đã phát']);
  }
  if (speak === false) return;
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 4).setValues([[
    Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy'),
    hasErrors ? 'NG' : 'OK',
    announceText,
    'Chưa'
  ]]);
}

function createAnnounceText(hasErrors, results) {
  if (!hasErrors) return 'Tất cả mã khuôn hợp lệ. OK';
  var names = [];
  for (var name in results) {
    if (results[name] && results[name].valid === false) names.push(name);
  }
  if (!names.length) return 'Cảnh báo. Không đạt.';
  return 'Cảnh báo. Line ' + names.join(', ') + '. Không đạt.';
}

function setNextAuto_(key) {
  PropertiesService.getScriptProperties().setProperty(key, String(Date.now() + 60000));
}

function getAutoStatus() {
  var p = PropertiesService.getScriptProperties();
  var now = Date.now();
  function left(k) {
    var t = Number(p.getProperty(k) || 0);
    return t > now ? Math.ceil((t - now) / 1000) : 0;
  }
  return { moldSec: left('nextAuto_mold'), labelSec: left('nextAuto_label') };
}

function scheduleFullCheck() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAllSilent') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('checkAllSilent').timeBased().after(60 * 1000).create();
  setNextAuto_('nextAuto_mold');
}

function scheduleFullLabelCheck() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAllLabelSilent') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('checkAllLabelSilent').timeBased().after(60 * 1000).create();
  setNextAuto_('nextAuto_label');
}

// ====================== [MỚI] EMAIL CẢNH BÁO KHI NG (WEB APP) ======================
// Giống bản Sheet: khi có lỗi sẽ gửi email chi tiết cho quản lý. Không thay đổi logic kiểm tra.
var ALERT_EMAILS = 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn';

function escHtml_(t) {
  if (t === null || t === undefined) return '';
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

function sendAlertEmail_(subject, title, source, bodyHtml, colored) {
  try {
    var timeStr = Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy');
    MailApp.sendEmail({
      to: ALERT_EMAILS,
      subject: subject + ' – ' + timeStr,
      htmlBody: '<h2 style="color:#dc2626;">⚠️ ' + escHtml_(title) + '</h2>' +
        '<p><strong>Thời gian:</strong> ' + timeStr + '</p>' +
        '<p><strong>Nguồn:</strong> ' + escHtml_(source) + '</p><hr>' + bodyHtml + '<hr>' +
        '<p><strong>Vui lòng mở file Google Sheets để xem chi tiết đầy đủ và sửa lỗi' +
        (colored ? ' (các ô lỗi đã được tô đỏ)' : '') + '.</strong></p>' +
        '<p>Trân trọng,<br><em>Script kiểm tra tự động</em></p>'
    });
    return true;
  } catch (e) {
    Logger.log('Lỗi gửi email: ' + e.message);
    return false;
  }
}

// Chi tiết lỗi của 1 sheet mã khuôn (dùng result.details do các hàm checkSheetXXX trả về)
function moldErrorHtml_(name, result) {
  var d = (result && result.details) || [];
  var items = d.slice(0, 50).map(function (err) {
    if (typeof err === 'object' && err !== null) {
      var pos = err.row
        ? ('Hàng ' + err.row + (err.col ? ', Cột ' + String.fromCharCode(64 + parseInt(err.col, 10)) : ''))
        : 'Vị trí không xác định';
      return '<li style="margin:8px 0;padding:10px;background:#fff5f5;border-left:4px solid #dc2626;border-radius:4px;">' +
        '<strong>' + pos + ':</strong> ' + escHtml_(err.message || 'Lỗi không xác định') + '<br>' +
        (err.current !== undefined ? '<strong>Hiện tại:</strong> <span style="background:#fee2e2;padding:2px 6px;border-radius:4px;">' + escHtml_(err.current) + '</span><br>' : '') +
        (err.expected !== undefined ? '<strong>Mong đợi:</strong> <span style="background:#fef3c7;padding:2px 6px;border-radius:4px;">' + escHtml_(err.expected) + '</span>' : '') +
        '</li>';
    }
    return '<li style="margin:8px 0;">' + escHtml_(err) + '</li>';
  }).join('');
  if (d.length > 50) items += '<li>... và ' + (d.length - 50) + ' lỗi khác</li>';
  return '<h4 style="color:#dc2626;">📌 Sheet: ' + escHtml_(name) + ' (' + ((result && result.errorCount) || 0) + ' lỗi)</h4>' +
    (items ? '<ul style="padding-left:20px;">' + items + '</ul>'
           : '<p>' + escHtml_((result && result.message) || 'Có lỗi nhưng không có thông tin chi tiết') + '</p>');
}

// Bảng tóm tắt OK/LỖI cho kiểm tra toàn bộ
function moldSummaryHtml_(results) {
  var rows = Object.keys(results).map(function (n) {
    var r = results[n] || {};
    return '<tr style="background-color:' + (r.valid ? '#f0fdf4' : '#fee2e2') + ';">' +
      '<td style="padding:10px;"><strong>' + escHtml_(n) + '</strong></td>' +
      '<td style="padding:10px;font-weight:bold;color:' + (r.valid ? '#166534' : '#dc2626') + ';">' + (r.valid ? 'OK' : 'LỖI') + '</td>' +
      '<td style="padding:10px;text-align:center;font-weight:bold;">' + (r.errorCount || 0) + '</td></tr>';
  }).join('');
  return '<h3>📋 Tóm tắt kết quả</h3>' +
    '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;">' +
    '<tr style="background-color:#dc2626;color:white;"><th>Tên Sheet</th><th>Trạng thái</th><th>Số lỗi</th></tr>' + rows + '</table>';
}

// Chi tiết lỗi label: dùng ĐÚNG điều kiện của checkLabelSheetByRule (chỉ để mô tả, không đổi kết quả)
function getLabelErrorDetails_(sheet, rule) {
  var out = [];
  if (!sheet) return ['<li><strong>Không tìm thấy sheet:</strong> ' + escHtml_(rule.sheet) + '</li>'];
  var last = sheet.getLastRow();
  if (last < 7) return ['<li><strong>Không có dữ liệu từ hàng 7:</strong> ' + escHtml_(rule.sheet) + '</li>'];
  var allowB = Array.isArray(rule.codesB) ? rule.codesB : [rule.codesB];
  var rows = sheet.getRange(7, 1, last - 6, 3).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    var a = String(rows[i][0] || '').trim();
    var b = String(rows[i][1] || '').trim();
    var c = String(rows[i][2] || '').trim();
    if (!a && !b && !c) continue;
    var rn = 7 + i;
    if (a !== rule.fixedA) out.push('<li><strong>Dòng ' + rn + '</strong>: Cột A sai → thực tế "<em>' + escHtml_(a) + '</em>", mong đợi "<em>' + escHtml_(rule.fixedA) + '</em>"</li>');
    if (allowB.indexOf(b) < 0) out.push('<li><strong>Dòng ' + rn + '</strong>: Cột B sai → thực tế "<em>' + escHtml_(b || '(trống)') + '</em>", mong đợi một trong "<em>' + escHtml_(allowB.join(', ')) + '</em>"</li>');
    if (rule.endC && (!c || c.slice(-1) !== rule.endC)) out.push('<li><strong>Dòng ' + rn + '</strong>: Cột C sai → thực tế "<em>' + escHtml_(c || '(trống)') + '</em>", phải kết thúc bằng "<em>' + escHtml_(rule.endC) + '</em>"</li>');
  }
  return out;
}

function labelErrorHtml_(name, items) {
  var shown = items.slice(0, 50).join('');
  if (items.length > 50) shown += '<li>... và ' + (items.length - 50) + ' lỗi khác</li>';
  return '<h4 style="color:#dc2626;">📌 Sheet: ' + escHtml_(name) + ' (' + items.length + ' lỗi)</h4><ol>' + shown + '</ol>';
}

function checkOneSheetHeadless(sheetName) {
  var map = getSheetCheckMap();
  if (!map[sheetName]) {
    return { status: 'NG', text: 'Không có line ' + sheetName, sheet: sheetName };
  }
  writeResultForESP32('Đang kiểm tra mã khuôn line ' + sheetName, false, true);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var hasErrors = false;
  var code = '';
  var xe = '';
  var ten = '';
  var result = null;
  var xr = null;
  if (!sheet || sheet.getLastRow() < 7) {
    hasErrors = true;
  } else {
    result = map[sheetName](sheet);
    xr = xeMerge_(result, checkXeSequence_(sheet));   // [MỚI] thêm điều kiện thứ tự số xe
    hasErrors = !result.valid;
    code = getLatestCodeFromSheet(sheet);
    xe = getXeValue_(sheet);
    ten = getTenValue_(sheet);
  }
  var status = hasErrors ? 'NG' : 'OK';
  var text = hasErrors
    ? ('Cảnh báo. Line ' + sheetName + '. Không đạt.')
    : ('Line ' + sheetName + '. OK');
  Utilities.sleep(5000);
  writeResultForESP32(text, hasErrors, true);
  writeHistory('KHUON', sheetName, status, text, xe, ten);
  scheduleFullCheck();
  var emailSent;
  if (hasErrors) {
    var r1 = result || { valid: false, errorCount: 0, details: [],
      message: !sheet ? 'Sheet không tìm thấy' : 'Không có dữ liệu từ hàng 7' };
    emailSent = sendAlertEmail_('CẢNH BÁO: Lỗi mã khuôn line ' + sheetName,
      'Phát hiện lỗi trong kiểm tra mã khuôn – Line ' + sheetName,
      'Web App SCAN SYSTEM 2.1 – kiểm tra mã khuôn line ' + sheetName,
      moldErrorHtml_(sheetName, r1), true);
  }
  return { status: status, text: text, sheet: sheetName, code: code, emailSent: emailSent,
    xeReason: (xr && !xr.valid) ? ('Số xe: ' + xr.errors[0].message) : '' };
}

function checkOneLabelHeadless(sheetName) {
  var rules = getLabelRules();
  var rule = null;
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].sheet === sheetName) { rule = rules[i]; break; }
  }
  if (!rule) {
    return { status: 'NG', text: 'Không có label ' + sheetName, sheet: sheetName };
  }
  writeResultForESP32('Đang kiểm tra label line ' + sheetName, false, true);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var hasErrors = !sheet || sheet.getLastRow() < 7 || !checkLabelSheetByRule(sheet, rule).valid;
  var status = hasErrors ? 'NG' : 'OK';
  var text = hasErrors
    ? ('Cảnh báo. Label ' + sheetName + '. Không đạt.')
    : ('Label ' + sheetName + '. OK');
  Utilities.sleep(5000);
  writeResultForESP32(text, hasErrors, true);
  writeHistory('LABEL', sheetName, status, text, sheet ? getXeValue_(sheet) : '', sheet ? getTenValue_(sheet) : '');
  scheduleFullLabelCheck();
  var emailSent;
  if (hasErrors) {
    var det = getLabelErrorDetails_(sheet, rule);
    emailSent = sendAlertEmail_('LỖI TEM NHÃN: ' + sheetName,
      'Phát hiện lỗi trong kiểm tra label – ' + sheetName,
      'Web App SCAN SYSTEM 2.1 – kiểm tra label ' + sheetName,
      labelErrorHtml_(sheetName, det), false);
  }
  return { status: status, text: text, sheet: sheetName, emailSent: emailSent };
}

function checkAllSilent() {
  var map = getSheetCheckMap();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = {};
  var hasErrors = false;
  for (var name in map) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 7) {
      results[name] = { valid: false, errorCount: 0, details: [], message: 'Sheet không tìm thấy hoặc không có dữ liệu từ hàng 7' };
      hasErrors = true;
      continue;
    }
    var result = map[name](sheet);
    xeMerge_(result, checkXeSequence_(sheet));         // [MỚI] thêm điều kiện thứ tự số xe
    results[name] = result;
    if (!result.valid) hasErrors = true;
  }
  var text = createAnnounceText(hasErrors, results);
  writeResultForESP32(text, hasErrors, hasErrors);
  writeHistory('KHUON-ALL', '', hasErrors ? 'NG' : 'OK', text);
  PropertiesService.getScriptProperties().deleteProperty('nextAuto_mold');
  var emailSent;
  if (hasErrors) {
    var body = moldSummaryHtml_(results) + '<hr><h3>🔍 Chi tiết lỗi cụ thể</h3>' +
      Object.keys(results).filter(function (n) { return !results[n].valid; })
        .map(function (n) { return moldErrorHtml_(n, results[n]); }).join('<hr>');
    emailSent = sendAlertEmail_('CẢNH BÁO: Kết quả kiểm tra dữ liệu - Có lỗi',
      'CẢNH BÁO: Phát hiện lỗi trong kiểm tra mã khuôn',
      'Web App SCAN SYSTEM 2.1 – kiểm tra toàn bộ mã khuôn', body, true);
  }
  return { status: hasErrors ? 'NG' : 'OK', text: text, speak: hasErrors, emailSent: emailSent };
}

function checkAllLabelSilent() {
  var rules = getLabelRules();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bad = [];
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var sheet = ss.getSheetByName(rule.sheet);
    if (!sheet || sheet.getLastRow() < 7 || !checkLabelSheetByRule(sheet, rule).valid) bad.push(rule.sheet);
  }
  var hasErrors = bad.length > 0;
  var text = hasErrors
    ? ('Cảnh báo. Label ' + bad.join(', ') + '. Không đạt.')
    : 'Tất cả label hợp lệ. OK';
  writeResultForESP32(text, hasErrors, hasErrors);
  writeHistory('LABEL-ALL', '', hasErrors ? 'NG' : 'OK', text);
  PropertiesService.getScriptProperties().deleteProperty('nextAuto_label');
  var emailSent;
  if (hasErrors) {
    var total = 0, body = '';
    for (var k = 0; k < rules.length; k++) {
      if (bad.indexOf(rules[k].sheet) < 0) continue;
      var items = getLabelErrorDetails_(ss.getSheetByName(rules[k].sheet), rules[k]);
      total += items.length;
      body += labelErrorHtml_(rules[k].sheet, items);
    }
    emailSent = sendAlertEmail_('LỖI MÃ KHUÔN & TEM NHÃN – ' + total + ' lỗi',
      'Phát hiện ' + total + ' lỗi trong kiểm tra label',
      'Web App SCAN SYSTEM 2.1 – kiểm tra toàn bộ label', body, false);
  }
  return { status: hasErrors ? 'NG' : 'OK', text: text, emailSent: emailSent };
}

function checkAllDataHeadless() { return checkAllSilent(); }

function getRecentHistory(limit) {
  var sheet = ensureHistorySheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var n = Math.min(limit || 20, last - 1);
  var rows = sheet.getRange(2, 1, n, 7).getDisplayValues();
  return rows.map(function(r) {
    return { time: r[0], kind: r[1], line: r[2], status: r[3], text: r[4], xe: r[5], ten: r[6] };
  });
}

function getHistoryStats() {
  var sheet = ensureHistorySheet_();
  var last = sheet.getLastRow();
  var ok = 0, ng = 0;
  if (last >= 2) {
    var vals = sheet.getRange(2, 4, last - 1, 1).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === 'OK') ok++;
      if (String(vals[i][0]) === 'NG') ng++;
    }
  }
  return { ok: ok, ng: ng, total: ok + ng };
}

function espPing() {
  PropertiesService.getScriptProperties().setProperty('espPing', String(Date.now()));
  return { ok: true };
}

function getEspStatus() {
  var t = Number(PropertiesService.getScriptProperties().getProperty('espPing') || 0);
  return {
    online: t > 0 && (Date.now() - t) < 90000,
    last: t ? Utilities.formatDate(new Date(t), 'GMT+7', 'HH:mm:ss') : 'Chưa kết nối',
    sec: t ? Math.floor((Date.now() - t) / 1000) : -1   // [MỚI] chỉ để hiển thị cảnh báo mất kết nối lâu
  };
}

function getLatestResult() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('IoT_Result');
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: 'none', text: 'Chưa có kết quả', time: '', beep: 0, light: 'GREEN' };
  }
  var row = sheet.getRange(2, 1, 1, 4).getValues()[0];
  var status = String(row[1] || '');
  var text = String(row[2] || '');
  var beep = 0;
  if (/đang kiểm tra/i.test(text)) beep = 1;
  if (status === 'NG' || /không đạt/i.test(text)) beep = 3;
  // Tín hiệu đèn tháp cho ESP32-S3: GREEN = OK, YELLOW = đang kiểm tra, RED = phát hiện lỗi.
  // Dùng lại đúng điều kiện đã có ở trên (không thêm luật mới).
  var light = 'GREEN';
  if (beep === 1) light = 'YELLOW';
  if (beep === 3) light = 'RED';
  return { status: status, text: text, time: String(row[0]), beep: beep, light: light };
}

function getDashboard() {
  return {
    esp: getEspStatus(),
    latest: getLatestResult(),
    auto: getAutoStatus(),
    stats: getHistoryStats(),
    history: getRecentHistory(150)
  };
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ui';
  if (action === 'latest') {
    espPing();
    return ContentService.createTextOutput(JSON.stringify(getLatestResult()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'ping') {
    espPing();
    return ContentService.createTextOutput('{"ok":true}')
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>SCAN SYSTEM 2.1</title>
<style>
:root{--bg:#070b14;--card:#10182a;--line:#1e2a44;--text:#eef4ff;--mute:#8aa0c4;--blue2:#2563eb;--ok:#22c55e;--ng:#f43f5e;--warn:#eab308}
*{box-sizing:border-box}
body{margin:0;font-family:Arial,sans-serif;background:radial-gradient(900px 400px at 10% -10%,#1d4ed633,transparent 50%),var(--bg);color:var(--text);min-height:100vh}
.wrap{max-width:860px;margin:0 auto;padding:14px}
.hero,.card{background:linear-gradient(180deg,#152038,#10182a);border:1px solid var(--line);border-radius:22px;padding:16px;margin-bottom:14px}
.hero{box-shadow:0 12px 24px -12px rgba(0,0,0,.6);margin-bottom:8px}
.brand{margin:0 0 4px;font-size:12px;letter-spacing:.16em;color:#60a5fa;font-weight:800}
h1{margin:0 0 8px;font-size:26px}
.sub{margin:0;color:var(--mute);font-size:13px}
.row{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}
.pill{background:#0b1220;border:1px solid var(--line);border-radius:999px;padding:8px 12px;font-size:13px;font-weight:800}
.esp{display:flex;align-items:center;gap:8px}
.dot{width:10px;height:10px;border-radius:50%;background:#64748b}
.on .dot{background:var(--ok);box-shadow:0 0 0 6px #22c55e33}
.off .dot{background:var(--ng)}
.tabs{display:grid;grid-template-columns:1fr 1fr;gap:8px;position:sticky;top:0;z-index:40;padding:8px 0;margin-bottom:6px;background:var(--bg);box-shadow:0 10px 14px -10px rgba(0,0,0,.8)}
.tab{border:1px solid var(--line);border-radius:14px;padding:14px 8px;font-size:15px;font-weight:900;color:#cbd5e1;background:#0b1220;min-height:52px}
.tab.on-semi{background:linear-gradient(180deg,#fce7f3,#f9a8d4);color:#831843;border-color:#fbcfe8;box-shadow:0 0 14px -6px #f9a8d4}
.tab.on-auto{background:linear-gradient(180deg,#a78bfa,#6d28d9);color:#fff;border-color:#a78bfa}
.h2{margin:0 0 14px;font-size:13px;letter-spacing:.08em;font-weight:800;text-transform:uppercase}
.khuon .h2{color:#93c5fd}
.label .h2{color:#5eead4}
.khuon{border-color:#1d4ed6}
.label{border-color:#0f766e}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px 14px}
.btn{border:0;border-radius:16px;padding:16px 8px;font-size:14px;font-weight:800;color:#fff;background:linear-gradient(180deg,#4f8cff,var(--blue2));min-height:58px}
.btn.lb{background:linear-gradient(180deg,#14b8a6,#0f766e)}
.btn.okbtn{background:linear-gradient(180deg,#4ade80,#15803d)!important}
.btn.ngbtn{background:linear-gradient(180deg,#fb7185,#9f1239)!important;animation:blink 1s infinite}
.btn.checking{background:linear-gradient(180deg,#fde047,#eab308)!important;color:#1f2937!important}
@keyframes blink{50%{opacity:.55}}
.btn:disabled{opacity:.4}
.btn.active{display:flex;align-items:center;justify-content:center;gap:8px}
.spin{width:16px;height:16px;border-radius:50%;border:2px solid #fff6;border-top-color:#fff;animation:sp .7s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
#box{display:flex;align-items:center;gap:16px;padding:18px;border-radius:18px;border:1px solid var(--line);min-height:100px;transition:.25s}
#box.ok{background:linear-gradient(135deg,#052e16,#0a3d20);border-color:#15803d;color:#dcfce7;box-shadow:inset 4px 0 0 var(--ok)}
#box.ng{background:linear-gradient(135deg,#3f0b16,#2a0a12);border-color:#9f1239;color:#ffe4e6;box-shadow:inset 4px 0 0 var(--ng);animation:blink 1.4s infinite}
#box.wait{background:linear-gradient(135deg,#2a2410,#1e293b);border-color:#a16207;color:#fef3c7;box-shadow:inset 4px 0 0 var(--warn);font-size:16px;font-weight:800}
.rs-ic{flex:none;width:60px;height:60px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:30px;font-weight:900}
#box.ok .rs-ic{background:#22c55e26;color:#4ade80;box-shadow:0 0 0 6px #22c55e14}
#box.ng .rs-ic{background:#f43f5e26;color:#fb7185;box-shadow:0 0 0 6px #f43f5e14}
#box.wait .rs-ic{background:#eab30826;box-shadow:0 0 0 6px #eab30814}
.rs-ic .spin{width:28px;height:28px;border-width:3px;border-color:#eab30855;border-top-color:#facc15}
.rs-main{min-width:0;flex:1}
.rs-cap{font-size:10.5px;letter-spacing:.14em;color:var(--mute);font-weight:800}
.rs-title{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:3px 0;font-size:26px;font-weight:900;color:#fff}
.rs-badge{padding:3px 12px;border-radius:999px;font-size:12px;letter-spacing:.08em;color:#fff}
#box.ok .rs-badge{background:#15803d}#box.ng .rs-badge{background:#be123c}#box.wait .rs-badge{background:#a16207}
.rs-msg{font-size:14px;font-weight:600;opacity:.85}
.rs-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:9px}
#rsTime{margin-top:10px;font-size:12px}
.wait{background:#1e293b;color:#fdba74}
.ok{background:#052e16;color:#86efac}
.ng{background:#3f0b16;color:#fda4af}
.modal-overlay{position:fixed;inset:0;background:rgba(2,6,15,.82);display:flex;align-items:center;justify-content:center;z-index:999;padding:20px}
.modal-overlay.hidden{display:none}
.modal-card{width:100%;max-width:460px;background:linear-gradient(180deg,#152038,#0d1526);border:2px solid var(--line);border-radius:26px;padding:32px 24px;text-align:center}
.modal-card.ok{border-color:var(--ok)}
.modal-card.ng{border-color:var(--ng);animation:blink 1s infinite}
.modal-status{font-size:44px;font-weight:900;margin-bottom:10px}
.modal-card.ok .modal-status{color:var(--ok)}
.modal-card.ng .modal-status{color:var(--ng)}
.modal-text{font-size:19px;font-weight:700;margin-bottom:10px}
.modal-time{color:var(--mute);font-size:13px;margin-bottom:26px}
.modal-mail{font-size:14px;font-weight:700;color:#fbbf24;margin:-14px 0 22px}
.modal-close{width:100%;padding:18px;border-radius:16px;border:0;font-size:17px;font-weight:900;color:#fff;background:linear-gradient(180deg,#4f8cff,var(--blue2));min-height:56px}
.okt{color:#4ade80;font-weight:800}.ngt{color:#fb7185;font-weight:800}
.hist{max-height:340px;overflow:auto;padding-right:4px;scrollbar-width:thin;scrollbar-color:#334155 transparent}
.hist::-webkit-scrollbar{width:6px}.hist::-webkit-scrollbar-thumb{background:#334155;border-radius:6px}
.hh,.item{display:grid;grid-template-columns:92px 112px 1fr 64px;gap:10px;align-items:center}
.hh{padding:0 14px 8px;font-size:10px;letter-spacing:.12em;color:var(--mute);font-weight:800;position:sticky;top:0;background:#10182a;z-index:2}
.hh span:last-child{text-align:right}
.item{padding:10px 14px;margin-bottom:6px;background:#0b1220;border:1px solid var(--line);border-left:4px solid var(--ok);border-radius:12px;font-size:13px;color:#cbd5e1;transition:.15s}
.item:hover{background:#111b30;transform:translateX(2px)}
.item.bad{border-left-color:var(--ng);background:#170d18}
.tm b{display:block;font-size:13px;color:#f1f5f9;font-variant-numeric:tabular-nums}
.tm small{color:var(--mute);font-size:11px}
.kd{justify-self:start;padding:3px 9px;border-radius:8px;font-size:10.5px;font-weight:800;letter-spacing:.05em;white-space:nowrap}
.kd.k{background:#1e3a8a55;color:#93c5fd;border:1px solid #1d4ed6}
.kd.l{background:#0f766e55;color:#5eead4;border:1px solid #0f766e}
.kd.a{background:#4c1d9555;color:#c4b5fd;border:1px solid #6d28d9}
.ln{display:flex;flex-wrap:wrap;align-items:center;gap:6px;min-width:0}
.ln b{color:#f8fafc;font-size:14px}
.ln .all{color:var(--mute);font-size:12px}
.chip{display:inline-flex;align-items:center;gap:5px;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:700;border:1px solid transparent;white-space:nowrap}
.chip i{font-style:normal;font-weight:600;opacity:.7}
.chip.xe{background:#f59e0b2e;color:#fde047;border-color:#f59e0b;font-size:15px;font-weight:900;padding:4px 14px;letter-spacing:.03em}
.chip.xe i{opacity:1;color:#fbbf24;font-weight:800}
.chip.ten{background:#14532d44;color:#bbf7d0;border-color:#15803d}
.st{justify-self:end;padding:5px 12px;border-radius:999px;font-size:12px;font-weight:900;letter-spacing:.06em}
.st.p{background:#052e16;color:#4ade80;border:1px solid #15803d}
.st.f{background:#3f0b16;color:#fb7185;border:1px solid #9f1239}
.empty{padding:26px;text-align:center;color:var(--mute);font-size:13px}
@media(max-width:560px){.hh{display:none}.item{grid-template-columns:76px 1fr 54px;grid-template-areas:"tm ln st" "tm kd kd";row-gap:6px}.tm{grid-area:tm}.kd{grid-area:kd}.ln{grid-area:ln}.st{grid-area:st}}
.time{margin-top:6px;font-size:12px;color:var(--mute)}
@media(min-width:720px){.grid{grid-template-columns:repeat(3,1fr)}h1{font-size:32px}}
.hf{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
.hf input{flex:1 1 100%;background:#0b1220;border:1px solid var(--line);border-radius:12px;padding:11px 14px;color:var(--text);font-size:14px;outline:none}
.hf input:focus{border-color:#38bdf8}
.fb{border:1px solid var(--line);background:#0b1220;color:#cbd5e1;border-radius:999px;padding:7px 14px;font-size:12px;font-weight:800}
.fb.on{background:linear-gradient(180deg,#38bdf8,#0284c7);color:#fff;border-color:#38bdf8}
.fb.ngb.on{background:linear-gradient(180deg,#fb7185,#9f1239);border-color:#fb7185}
/* Cảnh báo Trạm cảnh báo mất kết nối lâu */
.dsep{margin:12px 2px 6px;font-size:11px;letter-spacing:.12em;font-weight:800;color:var(--mute);text-transform:uppercase;display:flex;align-items:center;gap:10px}
.dsep::after{content:"";flex:1;height:1px;background:var(--line)}
.lastv{display:block;margin-top:5px;font-size:11px;font-weight:700;opacity:.92;letter-spacing:.02em}
.btn.lastng{box-shadow:0 0 0 2px #f43f5e}
.h2.tg{cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:10px;user-select:none;-webkit-tap-highlight-color:transparent}
.h2.tg::after{content:'▾';font-size:15px;opacity:.7;transition:transform .2s}
.card.closed .h2.tg::after{transform:rotate(-90deg)}
.card.closed>*:not(.h2){display:none}
.card.closed .h2{margin-bottom:0}
/* Nút phóng to chữ (A+) cho điện thoại ngoài xưởng */
.hero{position:relative}
.szbtn{position:absolute;top:14px;right:14px;border:1px solid var(--line);background:#0b1220;color:#cbd5e1;border-radius:999px;padding:6px 14px;font-size:13px;font-weight:900;cursor:pointer;-webkit-tap-highlight-color:transparent}
body.big .szbtn{background:linear-gradient(180deg,#38bdf8,#0284c7);color:#fff;border-color:#38bdf8}
body.big.semi .szbtn{background:linear-gradient(180deg,#fce7f3,#f9a8d4);color:#831843;border-color:#fbcfe8}
body.big .sub{font-size:15px}
body.big .pill{font-size:15px}
body.big .time{font-size:14px}
body.big .tab{font-size:17px}
body.big .h2{font-size:15px}
body.big .btn{font-size:17px;min-height:72px}
body.big .lastv{font-size:13px}
body.big .rs-title{font-size:30px}
body.big .rs-msg{font-size:16px}
body.big .item{font-size:15px}
body.big .tm b{font-size:15px}
body.big .ln b{font-size:17px}
body.big .chip{font-size:13px}
body.big .chip.xe{font-size:17px}
body.big .st{font-size:14px}
body.big .modal-text{font-size:22px}
.espwarn{display:none;align-items:center;gap:12px;margin-bottom:14px;padding:14px 16px;border-radius:16px;background:linear-gradient(135deg,#4c0519,#2a0a12);border:1px solid #be123c;box-shadow:inset 4px 0 0 #f43f5e;color:#ffe4e6;animation:blink 1.8s infinite}
.espwarn.show{display:flex}
.espwarn .wi{flex:none;font-size:26px}
.espwarn b{display:block;font-size:14px;color:#fff;margin-bottom:2px}
.espwarn span{font-size:12.5px;opacity:.9;line-height:1.45}
/* Chế độ SEMI: tông hồng nhẹ (chỉ đổi màu) */
body.semi .brand{color:#f9a8d4}
body.semi .fb.on{background:linear-gradient(180deg,#fce7f3,#f9a8d4);border-color:#fbcfe8;color:#831843}
body.semi .fb.ngb.on{background:linear-gradient(180deg,#fb7185,#9f1239);border-color:#fb7185;color:#fff}
body.semi .fb:hover,body.semi .tab:hover{border-color:#f9a8d4}
body.semi .hf input:focus{border-color:#f9a8d4}
body.semi .khuon{border-color:#9d4b78}
body.semi .khuon .h2{color:#fbcfe8}
body.semi .label{border-color:#8b5aa8}
body.semi .label .h2{color:#e9d5ff}
body.semi .btn:not(.okbtn):not(.ngbtn):not(.checking){background:linear-gradient(180deg,#fce7f3,#f9a8d4);color:#831843}
body.semi .btn.lb:not(.okbtn):not(.ngbtn):not(.checking){background:linear-gradient(180deg,#f3e8ff,#e9b8f5);color:#6b216f}
</style>
</head>
<body>
<div class="wrap">
  <div class="espwarn" id="espWarn"><div class="wi">⚠️</div><div><b id="espWarnT"></b><span id="espWarnS"></span></div></div>
  <div class="hero">
    <button class="szbtn" id="szBtn" onclick="togBig()" title="Phóng to / thu nhỏ chữ">A+</button>
    <p class="brand">SCAN SYSTEM 2.1</p>
    <h1>Kiểm tra mã khuôn</h1>
    <p class="sub">Chọn đúng line vừa scan. Semi line = VS. Auto line = VA. Logic kiểm tra không đổi.</p>
    <div class="row">
      <div id="esp" class="pill esp off"><span class="dot"></span><span id="espText">Trạm cảnh báo: Mất kết nối</span></div>
      <div class="pill" id="speak">Đèn: chờ thao tác</div>
      <div class="pill" id="stat">OK 0 / NG 0</div>
    </div>
    <div class="time" id="espTime"></div>
    <div class="time" id="cd">Tự động kiểm tra toàn bộ: --</div>
  </div>
  <div class="tabs">
    <button class="tab on-semi" id="tabSemi" onclick="setMode('semi')">SEMI LINE</button>
    <button class="tab" id="tabAuto" onclick="setMode('auto')">AUTO LINE</button>
  </div>
  <div class="card khuon">
    <p class="h2 tg" id="hKhuon" onclick="togCard(this)" title="Bấm để thu gọn / mở rộng">Kiểm tra mã khuôn — SEMI (VS)</p>
    <div class="grid" id="grid"></div>
  </div>
  <div class="card label">
    <p class="h2 tg" id="hLabel" onclick="togCard(this)" title="Bấm để thu gọn / mở rộng">Kiểm tra label — SEMI (VS)</p>
    <div class="grid" id="gridLb"></div>
  </div>
  <div class="card">
    <p class="h2">Kết quả</p>
    <div id="box" class="wait">Chưa kiểm tra</div>
    <div class="time" id="rsTime"></div>
  </div>
  <div class="card" id="histCard">
    <p class="h2 tg" id="hHist" onclick="togCard(this)" title="Bấm để thu gọn / mở rộng">Lịch sử</p>
    <div class="hf">
      <input id="q" placeholder="Tìm line, số xe, tên..." oninput="renderHist()">
      <button class="fb on" data-k="" onclick="setKind('')">Tất cả</button>
      <button class="fb" data-k="KHUON" onclick="setKind('KHUON')">Mã khuôn</button>
      <button class="fb" data-k="LABEL" onclick="setKind('LABEL')">Label</button>
      <button class="fb ngb" data-k="NG" onclick="setKind('NG')">Chỉ NG</button>
    </div>
    <div class="hist" id="hist"></div>
  </div>
</div>
<div class="modal-overlay hidden" id="resultModal">
  <div class="modal-card" id="modalCard">
    <div class="modal-status" id="modalStatus">OK</div>
    <div class="modal-text" id="modalText"></div>
    <div class="modal-time" id="modalTime"></div>
    <div class="modal-mail" id="modalMail"></div>
    <button class="modal-close" onclick="closeResultModal()">ĐÓNG</button>
  </div>
</div>
<script>
var busy=false, mode='semi', allMold=[], allLabel=[], lastHist=[], histKind='';
function isSemi(n){ return n.indexOf('VS')>=0; }
function isAuto(n){ return n.indexOf('VA')>=0; }
var KIND={'KHUON':['MÃ KHUÔN','k'],'LABEL':['LABEL','l'],'KHUON-ALL':['TẤT CẢ · KHUÔN','a'],'LABEL-ALL':['TẤT CẢ · LABEL','a']};
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function paintResult(j,H){
  var st=j.light==='YELLOW'?'wait':(j.status==='NG'?'ng':'ok');
  var txt=String(j.text||'');
  var m=/(Line|Label)[ ]+([A-Za-z0-9-]+)([.,])/.exec(txt);
  var title,cap;
  if(m&&m[3]==='.'){title=m[2];cap=m[1]==='Label'?'KIỂM TRA LABEL':'KIỂM TRA MÃ KHUÔN';}
  else if(txt.indexOf('Cảnh báo')>=0){title='Nhiều line không đạt';cap='KIỂM TRA TỰ ĐỘNG';}
  else{title='Toàn bộ hệ thống';cap='KIỂM TRA TỰ ĐỘNG';}
  var label={ok:'OK',ng:'NG',wait:'ĐANG KIỂM TRA'}[st];
  var icon={ok:'✓',ng:'✕',wait:'<span class="spin"></span>'}[st];
  var h0=(H||[])[0], chips='';
  if(h0&&h0.text===txt){
    if(h0.xe) chips+='<span class="chip xe"><i>Xe</i>'+esc(h0.xe)+'</span>';
    if(h0.ten) chips+='<span class="chip ten"><i>👤</i>'+esc(h0.ten)+'</span>';
  }
  var b=document.getElementById('box');
  b.className=st;
  b.innerHTML='<div class="rs-ic">'+icon+'</div><div class="rs-main"><div class="rs-cap">'+cap+'</div>'+
    '<div class="rs-title">'+esc(title)+'<span class="rs-badge">'+label+'</span></div>'+
    '<div class="rs-msg">'+esc(txt)+'</div>'+(chips?'<div class="rs-chips">'+chips+'</div>':'')+'</div>';
  document.getElementById('rsTime').innerText='🕒 '+(j.time||'');
}
var ESP_WARN_SEC=300;   // quá 5 phút không có tín hiệu thì hiện cảnh báo
function paintEspWarn(e){
  var box=document.getElementById('espWarn');
  if(!e||e.online){box.classList.remove('show');return;}
  var sec=(typeof e.sec==='number')?e.sec:-1, t='', sub='Đèn tháp và loa không hoạt động. Kiểm tra nguồn điện và WiFi của trạm.';
  if(sec<0){t='Chưa nhận được tín hiệu từ Trạm cảnh báo';}
  else if(sec>=ESP_WARN_SEC){
    var m=Math.floor(sec/60), h=Math.floor(m/60);
    t='Trạm cảnh báo mất kết nối '+(h>0?(h+' giờ '+(m%60)+' phút'):(m+' phút'));
  } else {box.classList.remove('show');return;}
  document.getElementById('espWarnT').innerText=t;
  document.getElementById('espWarnS').innerText=sub;
  box.classList.add('show');
}
function paintDash(d){
  if(!d) return;
  document.getElementById('esp').className='pill esp '+(d.esp.online?'on':'off');
  document.getElementById('espText').innerText=d.esp.online?'Trạm cảnh báo: Đang hoạt động':'Trạm cảnh báo: Mất kết nối';
  document.getElementById('espTime').innerText='Tín hiệu gần nhất: '+((d.esp.last&&d.esp.last!=='Chưa kết nối')?d.esp.last:'Chưa nhận tín hiệu');
  paintEspWarn(d.esp);
  var a=d.auto||{}, parts=[];
  if(a.moldSec>0) parts.push('Mã khuôn '+a.moldSec+' giây');
  if(a.labelSec>0) parts.push('Label '+a.labelSec+' giây');
  document.getElementById('cd').innerText=parts.length?('Tự động kiểm tra toàn bộ sau: '+parts.join(' · ')):'Tự động kiểm tra toàn bộ: chờ thao tác';
  document.getElementById('stat').innerText='OK '+(d.stats.ok||0)+' / NG '+(d.stats.ng||0);
  var j=d.latest;
  if(j&&j.status!=='none'){
    paintResult(j,d.history);
    var lightKey=j.light||(j.status==='NG'?'RED':(j.status==='OK'?'GREEN':'YELLOW'));
    var lightInfo={GREEN:{label:'🟢 Đèn xanh — OK',cls:'okt'},YELLOW:{label:'🟡 Đèn vàng — Đang kiểm tra',cls:''},RED:{label:'🔴 Đèn đỏ — Không đạt',cls:'ngt'}}[lightKey]||{label:'Đèn: --',cls:''};
    var speakEl=document.getElementById('speak');
    speakEl.innerText=lightInfo.label;
    speakEl.classList.remove('okt','ngt');
    if(lightInfo.cls) speakEl.classList.add(lightInfo.cls);
  }
  lastHist=d.history||[];
  renderHist();
  updateLineBadges();
}
function setKind(k){
  histKind=k;
  document.querySelectorAll('.fb').forEach(function(b){b.classList.toggle('on',b.dataset.k===k);});
  renderHist();
}
function dayStr(d){function z(n){return (n<10?'0':'')+n;}return z(d.getDate())+'/'+z(d.getMonth()+1)+'/'+d.getFullYear();}
function dayLabel(ds){
  var now=new Date(), y=new Date(now.getTime()-86400000);
  if(ds===dayStr(now)) return 'Hôm nay · '+ds;
  if(ds===dayStr(y)) return 'Hôm qua · '+ds;
  return ds;
}
function updateLineBadges(){
  var last={}, today=dayStr(new Date());
  for(var i=0;i<lastHist.length;i++){var r=lastHist[i]; if(!r.line) continue; var key=r.kind+'|'+r.line; if(!(key in last)) last[key]=r;}
  [['grid','KHUON'],['gridLb','LABEL']].forEach(function(p){
    var g=document.getElementById(p[0]); if(!g) return;
    Array.prototype.forEach.call(g.children,function(b){
      if(b.classList.contains('active')) return;   // đang kiểm tra: giữ nguyên chữ ĐANG KIỂM TRA
      var r=last[p[1]+'|'+b.dataset.name], html=esc(b.dataset.name), ng=!!(r&&r.status==='NG');
      if(r){
        var t=String(r.time||'').split(' '), tm=String(t[0]||'').slice(0,5)+((t[1]&&t[1]!==today)?(' '+t[1].slice(0,5)):'');
        html+='<span class="lastv">'+(ng?'🔴 NG':'🟢 OK')+' · '+esc(tm)+'</span>';
      }
      if(b.innerHTML!==html) b.innerHTML=html;
      b.classList.toggle('lastng',ng);
    });
  });
}
function renderHist(){
  var box=document.getElementById('hist'), top=box.scrollTop;
  var q=(document.getElementById('q').value||'').toLowerCase().trim();
  var inMode=(mode==='semi'?isSemi:isAuto);
  var H=lastHist.filter(function(r){
    if(r.line && !inMode(r.line)) return false;
    if(histKind==='NG' && r.status!=='NG') return false;
    if((histKind==='KHUON'||histKind==='LABEL') && r.kind!==histKind) return false;
    if(q && (r.line+' '+r.xe+' '+r.ten+' '+r.text).toLowerCase().indexOf(q)<0) return false;
    return true;
  }).slice(0,40);
  var semi=mode==='semi', prevDay='';
  document.getElementById('hHist').innerText='Lịch sử — '+(semi?'SEMI (VS)':'AUTO (VA)')+' · '+H.length+' kết quả';
  document.getElementById('histCard').style.borderColor=semi?'#f9a8d4':'#6d28d9';
  box.innerHTML=H.length?
    '<div class="hh"><span>THỜI GIAN</span><span>LOẠI</span><span>LINE / CHI TIẾT</span><span>KẾT QUẢ</span></div>'+
    H.map(function(r){
      var t=String(r.time||'').split(' ');
      var dsep=''; if((t[1]||'')!==prevDay){prevDay=t[1]||''; dsep='<div class="dsep">'+esc(dayLabel(prevDay))+'</div>';}
      var k=KIND[r.kind]||[r.kind||'--','a'];
      var bad=r.status==='NG';
      var det=r.line?'<b>'+esc(r.line)+'</b>':'<span class="all">'+esc(r.text)+'</span>';
      var xe=r.xe?'<span class="chip xe"><i>🚚</i>'+esc(r.xe)+'</span>':'';
      var ten=r.ten?'<span class="chip ten"><i>👤</i>'+esc(r.ten)+'</span>':'';
      return dsep+'<div class="item'+(bad?' bad':'')+'"><div class="tm"><b>'+esc(t[0])+'</b><small>'+esc(t[1]||'')+'</small></div>'+
        '<span class="kd '+k[1]+'">'+k[0]+'</span><div class="ln">'+det+xe+ten+'</div>'+
        '<span class="st '+(bad?'f':'p')+'">'+esc(r.status)+'</span></div>';
    }).join('')
    :'<div class="empty">Không có kết quả phù hợp</div>';
  box.scrollTop=top;
}
function markBtn(btn,status){
  if(!btn) return;
  btn.classList.remove('okbtn','ngbtn','checking');
  if(btn._resetTimer){clearTimeout(btn._resetTimer);btn._resetTimer=null;}
  if(status==='OK') btn.classList.add('okbtn');
  if(status==='NG') btn.classList.add('ngbtn');
  if(status==='OK'||status==='NG'){
    btn._resetTimer=setTimeout(function(){btn.classList.remove('okbtn','ngbtn');btn._resetTimer=null;},5000);
  }
}
var busyTimer=null;
function setBusy(v,btn){
  busy=v;
  if(busyTimer){clearInterval(busyTimer);busyTimer=null;}
  document.querySelectorAll('.btn').forEach(function(b){b.disabled=v;});
  if(btn){
    if(v){
      btn.classList.remove('okbtn','ngbtn');
      btn.classList.add('active','checking');
      btn.innerHTML='<span class="spin"></span><span>ĐANG KIỂM TRA</span>';
      var t0=Date.now();   // hiển thị số giây đã chờ để người vận hành biết hệ thống vẫn đang chạy
      busyTimer=setInterval(function(){
        var sp=btn.lastChild; if(sp&&sp.nodeType===1) sp.innerText='ĐANG KIỂM TRA · '+Math.floor((Date.now()-t0)/1000)+'s';
      },500);
    } else {
      btn.classList.remove('active','checking');
      btn.innerText=btn.dataset.name;
    }
  }
}
function openResultModal(j){
  var card=document.getElementById('modalCard');
  card.classList.remove('ok','ng');
  card.classList.add(j.status==='NG'?'ng':'ok');
  document.getElementById('modalStatus').innerText=j.status||'';
  document.getElementById('modalText').innerText=(j.text||'')+(j.xeReason?' — '+j.xeReason:'');
  document.getElementById('modalTime').innerText=j.time||'';
  document.getElementById('modalMail').innerText=j.emailSent===true?'📧 Đã gửi email cảnh báo đến quản lý':(j.emailSent===false?'⚠️ Chưa gửi được email — kiểm tra quyền gửi mail':'');
  document.getElementById('resultModal').classList.remove('hidden');
}
function closeResultModal(){ document.getElementById('resultModal').classList.add('hidden'); }
function focusResult(){
  var el=document.getElementById('box');
  if(!el) return;
  var card=el.closest?(el.closest('.card')||el):el;
  try{ card.scrollIntoView({behavior:'smooth',block:'center',inline:'nearest'}); }
  catch(e){ card.scrollIntoView(true); }
}
function run(fn,name,btn){
  if(busy) return;
  setBusy(true,btn);
  document.getElementById('box').className='wait';
  document.getElementById('box').innerText='SCAN SYSTEM — '+name;
  setTimeout(focusResult,60);
  google.script.run
    .withSuccessHandler(function(j){setBusy(false,btn);markBtn(btn,j.status);focusResult();openResultModal(j);tick();})
    .withFailureHandler(function(e){document.getElementById('box').className='ng';document.getElementById('box').innerText='Lỗi: '+e.message;setBusy(false,btn);})
    [fn](name);
}
function drawGrid(id,names,fn,cls){
  var g=document.getElementById(id); g.innerHTML='';
  (names||[]).forEach(function(n){
    var b=document.createElement('button');
    b.className=cls; b.innerText=n; b.dataset.name=n;
    b.onclick=function(){run(fn,n,b);};
    g.appendChild(b);
  });
}
function renderMode(){
  var mold=allMold.filter(mode==='semi'?isSemi:isAuto);
  var lab=allLabel.filter(mode==='semi'?isSemi:isAuto);
  document.body.classList.toggle('semi',mode==='semi');
  document.getElementById('tabSemi').className='tab'+(mode==='semi'?' on-semi':'');
  document.getElementById('tabAuto').className='tab'+(mode==='auto'?' on-auto':'');
  document.getElementById('hKhuon').innerText=mode==='semi'?'Kiểm tra mã khuôn — SEMI (VS)':'Kiểm tra mã khuôn — AUTO (VA)';
  document.getElementById('hLabel').innerText=mode==='semi'?'Kiểm tra label — SEMI (VS)':'Kiểm tra label — AUTO (VA)';
  drawGrid('grid',mold,'checkOneSheetHeadless','btn');
  drawGrid('gridLb',lab,'checkOneLabelHeadless','btn lb');
  renderHist();
  updateLineBadges();
}
function setMode(m){ mode=m; renderMode(); }
function tick(){ google.script.run.withSuccessHandler(paintDash).getDashboard(); }
google.script.run.withSuccessHandler(function(n){ allMold=n||[]; renderMode(); }).listSheetNames();
google.script.run.withSuccessHandler(function(n){ allLabel=n||[]; renderMode(); }).listLabelSheetNames();
function togBig(){
  var on=document.body.classList.toggle('big');
  document.getElementById('szBtn').innerText=on?'A−':'A+';
  try{localStorage.setItem('bigFont',on?'1':'0');}catch(e){}
}
function togCard(h){
  var c=h.parentNode; if(!c) return;
  c.classList.toggle('closed');
  try{
    var st=JSON.parse(localStorage.getItem('cardClosed')||'{}');
    st[h.id]=c.classList.contains('closed')?1:0;
    localStorage.setItem('cardClosed',JSON.stringify(st));
  }catch(e){}   // trình duyệt chặn lưu trữ thì chỉ nhớ trong phiên này
}
(function(){
  try{
    if(localStorage.getItem('bigFont')==='1'){document.body.classList.add('big');document.getElementById('szBtn').innerText='A−';}
  }catch(e){}
  try{
    var st=JSON.parse(localStorage.getItem('cardClosed')||'{}');
    Object.keys(st).forEach(function(id){
      var h=document.getElementById(id);
      if(h&&st[id]&&h.parentNode) h.parentNode.classList.add('closed');
    });
  }catch(e){}
})();
tick();
setInterval(tick,1000);
</script>
</body>
</html>
  `).setTitle('SCAN SYSTEM 2.1').addMetaTag('viewport','width=device-width, initial-scale=1');
}


// ====================== [MỚI] KIỂM TRA THỨ TỰ SỐ XE (n/N) THEO TỪNG LOẠI XE + LOT BTP ======================
// File MỚI – chỉ ĐỌC dữ liệu, thêm điều kiện NG. Không thay đổi logic kiểm tra mã khuôn hiện có.
//
// Quy định (kiểm tra từ dòng 7 xuống, dữ liệu mới ghi tiếp ở dưới):
//  1. Mỗi loại xe (HSG, Spacer, Packing... lấy từ chữ đứng trước n/N trong ô Số xe) có chuỗi số xe RIÊNG,
//     và theo dõi RIÊNG từng Lot BTP của loại đó, kể cả khi các lot xen kẽ (A, B, A, B...).
//  2. Cột lot của từng loại xe được TỰ NHẬN BIẾT (các cột có mã dạng xxxx$yyyy, cột nào khớp chuỗi xe nhất).
//     Ví dụ HSG đi theo cột "Lot BTP" thứ nhất, Spacer đi theo cột lot thứ hai.
//  3. Lần đầu một Lot BTP xuất hiện phải là 1/N (thiếu lịch sử → báo lỗi để bổ sung/chỉnh dữ liệu).
//  4. Cùng Lot BTP: số xe mới phải đi tuần tự 1/N → 2/N → ... ; không nhảy số, không lùi số, không đổi tổng N.
//  5. Quét lại một số xe ĐÃ ghi nhận của đúng lot đó là hợp lệ (không tính bước tiến mới, không báo lỗi).
//  6. Xe 1/1 (chỉ có 1 xe) luôn hợp lệ.
//  7. Không tìm thấy cột "Số xe" hoặc cột lot → bỏ qua, không báo NG.
var XE_FIRST_ROW_BASELINE = false; // false: lot nào lần đầu xuất hiện cũng phải là 1/N. true: riêng dòng 7 được coi là mốc (dùng nếu muốn nút Reset không gây NG).

function xeNorm_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function xeFindCols_(sheet) {
  var lastCol = sheet.getLastColumn(), xe = -1, lot = -1;
  if (lastCol < 1) return { xe: xe, lot: lot };
  var head = sheet.getRange(1, 1, 6, lastCol).getDisplayValues();
  var has = function (h, w) { return (' ' + h + ' ').indexOf(' ' + w + ' ') >= 0; };
  var tiers = [
    function (h) { return (has(h, 'lot') && (has(h, 'btp') || h.indexOf('ban thanh pham') >= 0)) || h.indexOf('lotbtp') >= 0; }
  ];
  for (var r = 0; r < head.length; r++) {
    for (var c = 0; c < head[r].length; c++) {
      var h = xeNorm_(head[r][c]);
      if (xe < 0 && h.indexOf('so xe') >= 0) xe = c + 1;
      if (lot < 0 && tiers[0](h)) lot = c + 1;
    }
  }
  return { xe: xe, lot: lot };
}

function xeColLetter_(idx) {   // idx tính từ 1
  var s = '';
  while (idx > 0) { var m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); }
  return s;
}

// Mô phỏng thứ tự số xe của MỘT loại xe khi dùng cột lot lotIdx; trả về danh sách lỗi
function xeSimulate_(recs, lotIdx, xeCol) {
  var lots = {}, errs = [];
  for (var i = 0; i < recs.length; i++) {
    var r = recs[i];
    if (r.N === 1 && r.n === 1) continue;                    // xe 1/1 luôn hợp lệ
    var lot = String(r.cells[lotIdx] || '').trim();
    var key = (lot || '(trống)').toUpperCase().replace(/\s+/g, ' ');
    var st = lots[key];
    var fail = function (msg, expected) {
      errs.push({ row: r.row, col: xeCol, message: msg, current: r.xeStr + ' (lot ' + (lot || 'trống') + ')', expected: expected });
    };
    if (!st) {
      // Lần đầu lot này xuất hiện → phải là 1/N
      if (r.n !== 1 && !(XE_FIRST_ROW_BASELINE && r.row === 7)) fail('Lot BTP mới phải bắt đầu từ xe 1/' + r.N, '1/' + r.N);
      st = lots[key] = { N: r.N, last: r.n, seen: {} };
      st.seen[r.n] = true;
    } else if (r.N !== st.N) {
      fail('Tổng số xe đổi trong cùng lot (' + st.N + ' → ' + r.N + ')', r.n + '/' + st.N);
    } else if (st.seen[r.n]) {
      // Quét lại số xe đã ghi nhận của đúng lot này → hợp lệ, không đổi thứ tự
    } else {
      var expected = (st.last + 1) + '/' + st.N;
      if (r.n > st.last + 1) fail('Nhảy số xe (' + st.last + ' → ' + r.n + ')', expected);
      else if (r.n < st.last) fail('Lùi số xe (' + r.n + '/' + r.N + ' chưa ghi nhận nhưng đã đến ' + st.last + '/' + r.N + ')', expected);
      st.seen[r.n] = true;
      if (r.n > st.last) st.last = r.n;
    }
  }
  return errs;
}

function checkXeSequence_(sheet) {
  var res = { valid: true, applicable: false, errors: [], note: '', lotCols: {} };
  try {
    var last = sheet.getLastRow();
    if (last < 7) return res;
    var cols = xeFindCols_(sheet);
    if (cols.xe < 0) { res.note = 'Không có cột Số xe'; return res; }

    var n = last - 6;
    var width = Math.max(sheet.getLastColumn(), cols.xe, cols.lot, 3);
    var data = sheet.getRange(7, 1, n, width).getDisplayValues();

    // Cột lot ứng viên: các cột (trừ cột A và cột Số xe) có ≥80% ô chứa dấu $, cộng cột "Lot BTP" tìm theo tiêu đề
    var cand = [];
    for (var c = 1; c < width; c++) {
      if (c === cols.xe - 1) continue;
      var tot = 0, dollar = 0;
      for (var i = 0; i < n; i++) {
        var v = String(data[i][c] || '').trim();
        if (v) { tot++; if (v.indexOf('$') >= 0) dollar++; }
      }
      if (tot > 0 && dollar / tot >= 0.8) cand.push(c);
    }
    if (cols.lot > 0 && cand.indexOf(cols.lot - 1) < 0) cand.push(cols.lot - 1);
    if (!cand.length) { res.note = 'Không tìm thấy cột Lot BTP'; return res; }
    cand.sort(function (x, y) { return x - y; });
    res.applicable = true;

    // Gom theo loại xe (chữ đứng trước n/N: hsg, spacer, packing...)
    var byType = {};
    for (var r = 0; r < n; r++) {
      var d = data[r], row = 7 + r;
      if (d.join('').trim() === '') continue;
      var xeStr = String(d[cols.xe - 1] || '').trim();
      var m = /(\d+)\s*\/\s*(\d+)/.exec(xeStr);
      if (!m) { res.errors.push({ row: row, col: cols.xe, message: 'Số xe không đúng dạng n/N', current: xeStr, expected: 'ví dụ 1/4' }); continue; }
      var cn = parseInt(m[1], 10), cN = parseInt(m[2], 10);
      if (cn < 1 || cn > cN) {
        res.errors.push({ row: row, col: cols.xe, message: 'Số xe ngoài phạm vi 1/' + cN, current: xeStr, expected: '1/' + cN + ' đến ' + cN + '/' + cN });
        continue;
      }
      var type = xeNorm_(xeStr.slice(0, m.index)).replace(/ /g, '');
      (byType[type] = byType[type] || []).push({ row: row, xeStr: xeStr, n: cn, N: cN, cells: d });
    }

    // Mỗi loại xe chọn cột lot cho ít lỗi nhất (hòa → ưu tiên cột "Lot BTP" theo tiêu đề, rồi cột bên trái)
    Object.keys(byType).forEach(function (t) {
      var best = null;
      cand.forEach(function (ci) {
        var errs = xeSimulate_(byType[t], ci, cols.xe);
        if (!best || errs.length < best.errs.length ||
            (errs.length === best.errs.length && ci === cols.lot - 1 && best.ci !== cols.lot - 1)) {
          best = { ci: ci, errs: errs };
        }
      });
      res.lotCols[t || '(không tên)'] = xeColLetter_(best.ci + 1);
      res.errors = res.errors.concat(best.errs);
    });
    res.errors.sort(function (x, y) { return x.row - y.row; });

    res.valid = res.errors.length === 0;
    if (res.errors.length) {
      try {   // lỗi khi tô màu không được làm mất kết quả NG
        var cells = res.errors.map(function (e) { return sheet.getRange(e.row, cols.xe).getA1Notation(); });
        sheet.getRangeList(cells).setBackground('#ff0000');
      } catch (colorErr) { res.note = 'Không tô được màu: ' + colorErr.message; }
    }
  } catch (e) {
    res.note = 'Lỗi kiểm tra số xe: ' + e.message;   // lỗi kỹ thuật không được làm sai kết quả cũ
    if (!res.errors.length) res.valid = true;
  }
  return res;
}

// Gộp lỗi số xe vào kết quả của hàm checkSheetXXX (cùng định dạng details) – chỉ thêm, không sửa kết quả cũ
function xeMerge_(result, xr) {
  if (result && xr && !xr.valid) {
    result.valid = false;
    result.errorCount = (result.errorCount || 0) + xr.errors.length;
    result.details = (result.details || []).concat(xr.errors);
  }
  return xr;
}

// Chạy tay trong Apps Script (chọn hàm debugXe → Chạy → xem Nhật ký) để biết vì sao có/không báo lỗi số xe
function debugXe(sheetName) {
  sheetName = sheetName || 'VA015-E';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) { Logger.log('Không tìm thấy sheet ' + sheetName); return; }
  var cols = xeFindCols_(sheet);
  Logger.log('Sheet ' + sheetName + ' | cột Số xe = ' + cols.xe + ' | cột Lot BTP = ' + cols.lot + '  (-1 nghĩa là KHÔNG tìm thấy)');
  var r = checkXeSequence_(sheet);
  Logger.log('Áp dụng: ' + r.applicable + ' | Hợp lệ: ' + r.valid + ' | Số lỗi: ' + r.errors.length + ' | Ghi chú: ' + r.note);
  Logger.log('Cột lot theo loại xe: ' + JSON.stringify(r.lotCols));
  r.errors.forEach(function (e) { Logger.log('Dòng ' + e.row + ': ' + e.message + ' → ' + e.current + ' (mong đợi ' + e.expected + ')'); });
}

// ====================== [MỚI] ĐỒNG BỘ KIỂM TRA SỐ XE CHO BẢN GOOGLE SHEET ======================
// Khai báo sau cùng nên thay thế checkAllData cũ. Logic giữ nguyên 100%, chỉ thêm điều kiện số xe (n/N).
function checkAllData() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Không thể truy cập spreadsheet hiện tại.');
    }

    var ui = SpreadsheetApp.getUi();
    if (!ui) {
      throw new Error('Không thể khởi tạo giao diện người dùng.');
    }

    // Hiển thị giao diện "Đang kiểm tra"
    var checkingHtml = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">' +
      '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">' +
      '<style>' +
      'body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #e6f3ff 0%, #d1e8ff 100%); }' +
      '.progress-bar { width: 100%; background-color: #e0e0e0; height: 10px; border-radius: 5px; overflow: hidden; margin-top: 20px; }' +
      '.progress { width: 0; height: 100%; background-color: #10b981; animation: progress 2s infinite; }' +
      '@keyframes progress { 0% { width: 0; } 100% { width: 100%; } }' +
      '</style>' +
      '</head>' +
      '<body class="flex items-center justify-center min-h-screen">' +
      '<div class="bg-white p-6 rounded-lg shadow-xl max-w-md w-full text-center">' +
      '<h2 class="text-2xl font-bold text-gray-800 mb-4">Đang kiểm tra...</h2>' +
      '<h3 class="text-xl font-semibold text-blue-800 mb-4">CẤU TẠO MÃ KHUÔN LINH KIỆN</h3>' +
      '<p class="text-gray-600 mt-4">Vui lòng chờ trong khi dữ liệu được kiểm tra.</p>' +
      '<div class="progress-bar"><div class="progress"></div></div>' +
      '<audio id="startSound" autoplay><source src="https://www.soundjay.com/buttons/beep-02.mp3" type="audio/mpeg"></audio>' +
      '<script>' +
      'document.getElementById("startSound").play().catch(function(error) { console.log("Lỗi phát âm thanh: ", error); });' +
      '</script>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(400).setHeight(300);
    ui.showModalDialog(checkingHtml, 'Đang kiểm tra');

    // Đợi 1.5 giây để hiển thị giao diện "Đang kiểm tra"
    Utilities.sleep(1500);

    // Định nghĩa các sheet cần kiểm tra
    var sheetsToCheck = {
      'VA001-B': { checkFunction: checkSheet8794B },
      'VA015-A': { checkFunction: checkSheet7040GA },
      'VA015-B': { checkFunction: checkSheet7040GB },
      'VA015-E': { checkFunction: checkSheet7040GE },
      'VS001-A': { checkFunction: checkSheetVS001A },
      'VS001-B': { checkFunction: checkSheetVS001B },
      'VS001-C': { checkFunction: checkSheetVS001C },
      'VA019-A': { checkFunction: checkSheetVA019A },
      'VA019-B': { checkFunction: checkSheetVA019B },
      'VA019-C': { checkFunction: checkSheetVA019C },
      'VA003': { checkFunction: checkSheetVA003 },
      'VS051': { checkFunction: checkSheetVS051 },
      'VA016-A': { checkFunction: checkSheet7041GA },
      'VA016-B': { checkFunction: checkSheet7041GB },
    };

    // Tạo hoặc lấy sheet log và tự động điều chỉnh cột
    var logSheet = spreadsheet.getSheetByName('CheckLog') || spreadsheet.insertSheet('CheckLog');
    if (!logSheet) {
      throw new Error('Không thể tạo hoặc truy cập sheet CheckLog.');
    }
    logSheet.clear();
    logSheet.appendRow(['Thời gian', 'Tên Sheet', 'Trạng thái', 'Số lỗi', 'Thông báo']);

    // Thu thập kết quả và đánh dấu lỗi
    var results = {};
    var startTime = new Date();
    var startTimeFormatted = Utilities.formatDate(startTime, 'GMT+7', 'HH:mm:ss');
    var erroredSheet = null;
    var hasErrors = false;
    for (var sheetName in sheetsToCheck) {
      try {
        var sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          results[sheetName] = { valid: false, errorCount: 0, message: 'Sheet không tìm thấy', details: [] };
          logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Sheet không tìm thấy']);
          hasErrors = true;
          continue;
        }
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        if (lastRow < 7) {
          results[sheetName] = { valid: false, errorCount: 0, message: 'Không có dữ liệu từ hàng 7', details: [] };
          logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Không có dữ liệu từ hàng 7']);
          hasErrors = true;
          continue;
        }

        var result = sheetsToCheck[sheetName].checkFunction(sheet);
        xeMerge_(result, checkXeSequence_(sheet));   // [MỚI] thêm điều kiện thứ tự số xe (đồng bộ với Web App)
        if (!result.valid && result.message === 'Không có lỗi') result.message = 'Phát hiện lỗi số xe';
        results[sheetName] = result;
        logSheet.appendRow([new Date(), sheetName, result.valid ? 'OK' : 'LỖI', result.errorCount, result.message]);
        if (!result.valid && !erroredSheet) {
          erroredSheet = sheet;
        }
        if (!result.valid) hasErrors = true;
      } catch (e) {
        results[sheetName] = { valid: false, errorCount: 0, message: 'Lỗi xử lý sheet: ' + e.message, details: [] };
        logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Lỗi xử lý sheet: ' + e.message]);
        hasErrors = true;
      }
    }
    var endTime = new Date();
    var endTimeFormatted = Utilities.formatDate(endTime, 'GMT+7', 'HH:mm:ss');
    var duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log('Thời gian thực thi: ' + duration + ' giây');

    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.autoResizeColumns(1, 5);
      for (var i = 1; i <= lastRow; i++) {
        logSheet.setRowHeight(i, 21);
      }
    }

    var htmlOutput = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">' +
      '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">' +
      '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">' +
      '<style>' +
      'body { font-family: "Roboto", sans-serif; ' +
      (hasErrors 
        ? 'background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%); }'
        : 'background: linear-gradient(135deg, #e6f3ff 0%, #d1e8ff 100%); }') +
      '.time-badge { background-color: #e2e8f0; color: #2c3e50; padding: 8px 12px; border-radius: 20px; display: inline-block; margin: 4px; }' +
      'table { width: 100%; border-collapse: separate; border-spacing: 0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }' +
      'th { background-color: ' + (hasErrors ? '#dc2626' : '#10b981') + '; color: white; padding: 12px; text-align: left; }' +
      'td { padding: 12px; text-align: left; border: 1px solid #ddd; vertical-align: top; }' +
      'tr:nth-child(even) { background-color: #f8fafc; }' +
      'tr:hover { background-color: ' + (hasErrors ? '#fca5a5' : '#e6fffa') + '; transition: background-color 0.3s; }' +
      '.error-row { background-color: #dc2626; color: white; border: 2px solid #991b1b; }' +
      '.error-row:hover { background-color: #b91c1c; border-color: #7f1d1d; }' +
      '.status-icon { margin-right: 8px; }' +
      '.error-detail { font-size: 0.95em; line-height: 1.6; }' +
      '.error-detail ul { margin: 8px 0; padding-left: 22px; list-style-type: disc; }' +
      '.error-item { margin-bottom: 10px; padding: 10px; background-color: rgba(220, 38, 38, 0.15); border-left: 5px solid #dc2626; border-radius: 6px; }' +
      '.error-pos { font-weight: bold; color: #991b1b; }' +
      '.error-current { font-family: monospace; background: #ffffff; padding: 3px 8px; border-radius: 4px; font-weight: bold; color: #000000; border: 1px solid #999999; }' +
      '.error-expected { font-family: monospace; background: #ffffcc; padding: 3px 8px; border-radius: 4px; font-weight: bold; color: #000000; border: 1px solid #999999; }' +
      '.no-detail { color: #666; font-style: italic; }' +
      '</style>' +
      '</head>' +
      '<body class="flex items-center justify-center min-h-screen" onload="playSoundIfErrors(\'' + (hasErrors ? 'true' : 'false') + '\')">' +
      '<div class="bg-white p-8 rounded-lg shadow-2xl max-w-6xl w-full overflow-x-auto">' +
      '<h3 class="text-2xl font-semibold ' + (hasErrors ? 'text-red-700' : 'text-blue-800') + ' text-center mb-4">CẤU TẠO MÃ KHUÔN LINH KIỆN</h3>' +
      '<h2 class="text-3xl font-bold text-gray-800 text-center mb-6">Kết Quả Kiểm Tra</h2>' +
      '<div class="flex justify-center space-x-3 mb-6">' +
      '<span class="time-badge">Bắt đầu: ' + startTimeFormatted + '</span>' +
      '<span class="time-badge">Kết thúc: ' + endTimeFormatted + '</span>' +
      '<span class="time-badge">Thời gian: ' + duration.toFixed(2) + 's</span>' +
      '</div>' +
      '<div class="overflow-x-auto">' +
      '<table class="min-w-full">' +
      '<thead><tr>' +
      '<th class="px-4 py-3">Tên Sheet</th>' +
      '<th class="px-4 py-3">Trạng thái</th>' +
      '<th class="px-4 py-3">Số lỗi</th>' +
      '<th class="px-4 py-3">Thông báo tổng quát</th>' +
      '<th class="px-4 py-3" style="width: 450px;">Chi tiết lỗi cụ thể</th>' +
      '</tr></thead>' +
      '<tbody>' +
      (function() {
        var rows = [];
        function escapeHtml(text) {
          if (text === null || text === undefined) return '';
          return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
        }

        for (var sheetName in results) {
          var result = results[sheetName];
          var detailContent = '';

          if (result.valid) {
            detailContent = '<div class="text-green-600 font-medium">✓ Tất cả dữ liệu hợp lệ</div>';
          } else {
            if (result.details && Array.isArray(result.details) && result.details.length > 0) {
              detailContent = '<div class="error-detail"><ul>';
              result.details.forEach(function(err) {
                if (typeof err === 'object' && err !== null) {
                  var row = err.row || '?';
                  var col = err.col ? String.fromCharCode(64 + parseInt(err.col)) : '?';
                  var position = 'Hàng ' + row + ', Cột ' + col;

                  detailContent += '<li class="error-item">' +
                    '<span class="error-pos">' + position + ':</span> ' + escapeHtml(err.message || 'Lỗi không xác định') + '<br>';
                  if (err.current !== undefined) detailContent += '→ Hiện tại: <span class="error-current">' + escapeHtml(err.current) + '</span><br>';
                  if (err.expected !== undefined) detailContent += '→ Mong đợi: <span class="error-expected">' + escapeHtml(err.expected) + '</span>';
                  detailContent += '</li>';
                } else {
                  detailContent += '<li class="error-item">' + escapeHtml(err) + '</li>';
                }
              });
              detailContent += '</ul></div>';
            } else if (result.message) {
              detailContent = '<div class="error-item">' + escapeHtml(result.message) + '</div>';
            } else {
              detailContent = '<div class="no-detail">Có lỗi nhưng không có thông tin chi tiết</div>';
            }
          }

          rows.push('<tr' + (result.valid ? '' : ' class="error-row"') + '>' +
                    '<td class="border px-4 py-3 font-semibold">' + escapeHtml(sheetName) + '</td>' +
                    '<td class="border px-4 py-3"><i class="status-icon fa ' + (result.valid ? 'fa-check-circle text-green-600' : 'fa-times-circle text-white') + '"></i> <strong>' + (result.valid ? 'OK' : 'LỖI') + '</strong></td>' +
                    '<td class="border px-4 py-3 text-center text-lg font-bold">' + result.errorCount + '</td>' +
                    '<td class="border px-4 py-3">' + (result.message ? escapeHtml(result.message) : '-') + '</td>' +
                    '<td class="border px-4 py-3">' + detailContent + '</td>' +
                    '</tr>');
        }
        return rows.join('');
      })() +
      '</tbody></table>' +
      '</div>' +
      (hasErrors ? '<audio id="errorSound"><source src="https://www.soundjay.com/buttons/beep-01a.mp3" type="audio/mpeg"></audio>' : '') +
      '<script>' +
      'function playSoundIfErrors(hasErrors) {' +
      '  if (hasErrors === "true") {' +
      '    var audio = document.getElementById("errorSound");' +
      '    if (audio) audio.play().catch(function(error) { console.log("Lỗi phát âm thanh: ", error); });' +
      '  }' +
      '}' +
      '</script>' +
      '<div class="text-center mt-8">' +
      '<button onclick="google.script.host.close()" class="bg-' + (hasErrors ? 'red' : 'blue') + '-600 text-white px-8 py-4 rounded-lg hover:bg-' + (hasErrors ? 'red' : 'blue') + '-700 transition font-semibold text-lg shadow-lg">Đóng</button>' +
      '</div>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(950).setHeight(650);

    ui.showModalDialog(htmlOutput, hasErrors ? 'CẢNH BÁO HỆ THỐNG:  Phát hiện lỗi...Thông tin đã gửi đến quản lý để xử lý kịp thời ' : 'Tất cả dữ liệu đều OK');

    if (hasErrors) {
      try {
        var emailSubject = hasErrors ? 'CẢNH BÁO: Kết quả kiểm tra dữ liệu - Có lỗi' : 'Kết quả kiểm tra dữ liệu - Tất cả OK';
        var emailBody = `
  <h2 style="color:#dc2626;">⚠️ CẢNH BÁO: Phát hiện lỗi trong kiểm tra mã khuôn</h2>
  <p><strong>Thời gian bắt đầu:</strong> ${startTimeFormatted}</p>
  <p><strong>Thời gian kết thúc:</strong> ${endTimeFormatted}</p>
  <p><strong>Thời gian thực thi:</strong> ${duration.toFixed(2)} giây</p>
  <hr>
  <h3>📋 Tóm tắt kết quả</h3>
  <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">
    <tr style="background-color: #10b981; color: white;">
      <th style="padding: 10px;">Tên Sheet</th>
      <th style="padding: 10px;">Trạng thái</th>
      <th style="padding: 10px;">Số lỗi</th>
    </tr>
    ${Object.keys(results).map(function(sheetName) {
      var result = results[sheetName];
      return `<tr style="background-color: ${result.valid ? '#f0fdf4' : '#fee2e2'};">
                <td style="padding: 10px;"><strong>${sheetName}</strong></td>
                <td style="padding: 10px; color: ${result.valid ? '#166534' : '#dc2626'}; font-weight: bold;">${result.valid ? 'OK' : 'LỖI'}</td>
                <td style="padding: 10px; text-align: center; font-weight: bold; color: ${result.valid ? '#166534' : '#dc2626'};">${result.errorCount}</td>
              </tr>`;
    }).join('')}
  </table>
  <hr>
  <h3>🔍 Chi tiết lỗi cụ thể</h3>
  ${Object.keys(results).filter(sheetName => !results[sheetName].valid && results[sheetName].details && results[sheetName].details.length > 0).map(sheetName => {
    var result = results[sheetName];
    var detailList = result.details.map(err => {
      if (typeof err === 'object' && err !== null) {
        var position = err.row && err.col ? `Hàng ${err.row}, Cột ${String.fromCharCode(64 + parseInt(err.col))}` : 'Vị trí không xác định';
        var msg = err.message || 'Lỗi không xác định';
        var current = err.current !== undefined ? `<strong>Hiện tại:</strong> <span style="background:#fee2e2; padding:2px 6px; border-radius:4px;">${err.current}</span>` : '';
        var expected = err.expected !== undefined ? `<strong>Mong đợi:</strong> <span style="background:#fef3c7; padding:2px 6px; border-radius:4px;">${err.expected}</span>` : '';
        return `<li style="margin:8px 0; padding:10px; background:#fff5f5; border-left:4px solid #dc2626; border-radius:4px;">
                  <strong>${position}:</strong> ${msg}<br>
                  ${current ? current + '<br>' : ''}
                  ${expected}
                </li>`;
      } else {
        return `<li style="margin:8px 0;">${err}</li>`;
      }
    }).join('');
    return `<h4 style="color:#dc2626;">📌 Sheet: ${sheetName} (${result.errorCount} lỗi)</h4>
            <ul style="padding-left:20px;">${detailList}</ul>`;
  }).join('<hr>')}
  <hr>
  <p><strong>Vui lòng mở file Google Sheets để xem chi tiết đầy đủ và sửa lỗi (các ô lỗi đã được tô đỏ).</strong></p>
  <p>Trân trọng,<br><em>Script kiểm tra tự động</em></p>
`;
        var emailAddress = 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn';
        MailApp.sendEmail({
          to: emailAddress,
          subject: emailSubject,
          htmlBody: emailBody
        });
      } catch (e) {
        Logger.log('Lỗi gửi email: ' + e.message);
      }
    }

    if (erroredSheet) {
      spreadsheet.setActiveSheet(erroredSheet);
      erroredSheet.setActiveRange(erroredSheet.getRange(7, 1));
    }
  } catch (e) {
    Logger.log('Lỗi chính trong checkAllData: ' + e.message);
    ui.alert('Lỗi nghiêm trọng', 'Đã xảy ra lỗi: ' + e.message + '\nVui lòng kiểm tra log để biết chi tiết.', ui.ButtonSet.OK);
  }
}

// ====================== [MỚI v3] GIAO DIỆN WEB APP GỌN HƠN — CHỈ THÊM Ở CUỐI FILE, KHÔNG SỬA CODE CŨ ======================
// Dán khối này xuống CUỐI CÙNG của file. Vì khai báo sau cùng nên doGet này thay thế mọi doGet ở phía trên
// (kể cả bản giao diện v2 nếu bạn đã dán). Các doGet phía trên vẫn nằm nguyên, không bị đụng tới.
// QUAY LẠI BẢN TRƯỚC: xóa toàn bộ khối này (từ dòng chú thích này đến hết file) rồi Deploy lại.
// Khối này chỉ đổi giao diện (CSS / HTML / JS phía trình duyệt). action=latest, action=ping và các hàm google.script.run
// (listSheetNames, listLabelSheetNames, getDashboard, checkOneSheetHeadless, checkOneLabelHeadless) giữ nguyên.
// Lưu ý: nằm trong template literal nên KHÔNG dùng dấu huyền, ${ hay dấu gạch chéo ngược trong HTML/JS bên dưới.
function doGetUi_(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ui';
  if (action === 'latest') {
    espPing();
    return ContentService.createTextOutput(JSON.stringify(getLatestResult()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'ping') {
    espPing();
    return ContentService.createTextOutput('{"ok":true}')
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>SCAN SYSTEM 2.1</title>
<style>
/* ===== Màu theo đúng một nghĩa: xanh lá = OK, đỏ = NG, vàng = đang kiểm tra / quá hạn. Còn lại dùng màu trung tính ===== */
:root{--bg:#0a0f1c;--card:#111a2e;--card2:#0d1424;--line:#22304d;--text:#eef4ff;--mute:#93a4c4;--btn:#1b2740;--btnh:#26365a;
--ok:#22c55e;--okbg:#0c3a20;--oktx:#bbf7d0;--ng:#f43f5e;--ngbg:#431220;--ngtx:#fecdd3;--warn:#eab308;--warnbg:#3b3008;--warntx:#fef08a;
--acc:#e2e8f0;--acctx:#0b1220;--bar:rgba(10,15,28,.94)}
:root.light{--bg:#eef1f6;--card:#ffffff;--card2:#f3f5f9;--line:#b8c3d6;--text:#0b1220;--mute:#475569;--btn:#e6ebf3;--btnh:#d6deeb;
--ok:#15803d;--okbg:#dcfce7;--oktx:#14532d;--ng:#be123c;--ngbg:#ffe4e6;--ngtx:#881337;--warn:#a16207;--warnbg:#fef3c7;--warntx:#713f12;
--acc:#0b1220;--acctx:#ffffff;--bar:rgba(238,241,246,.95)}
*{box-sizing:border-box}
html{background:var(--bg)}
body{margin:0;font-family:Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:92px;--s:1}
body.big{--s:1.22}
.wrap{max-width:860px;margin:0 auto;padding:12px}
.top,.card{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:14px;margin-bottom:12px}
.top{padding:12px 14px;margin-bottom:8px}
.top1{display:flex;align-items:center;justify-content:space-between;gap:10px}
.brand{display:block;font-size:11px;letter-spacing:.14em;color:var(--mute);font-weight:800}
h1{margin:2px 0 0;font-size:calc(20px*var(--s))}
.tools{display:flex;gap:6px;flex:none}
.tbtn{border:1px solid var(--line);background:var(--card2);color:var(--text);border-radius:999px;min-width:44px;height:38px;padding:0 12px;font-size:14px;font-weight:900;cursor:pointer;-webkit-tap-highlight-color:transparent}
/* Thanh trạng thái một dòng */
.sbar{display:flex;flex-wrap:wrap;align-items:center;gap:6px 16px;margin-top:10px;padding-top:10px;border-top:1px solid var(--line);font-size:calc(13px*var(--s));font-weight:700;color:var(--mute)}
.sp{display:inline-flex;align-items:center;gap:7px}
.dot{width:10px;height:10px;border-radius:50%;background:#64748b}
.sp.on{color:var(--text)}
.sp.on .dot{background:var(--ok);box-shadow:0 0 0 4px rgba(34,197,94,.22)}
.sp.off{color:var(--ng)}
.sp.off .dot{background:var(--ng)}
.okt{color:var(--ok);font-weight:900}
.ngt{color:var(--ng);font-weight:900}
.sep{opacity:.5;margin:0 4px}
/* Thanh chọn chế độ: nằm sát đáy trên điện thoại */
.tabs{position:fixed;left:0;right:0;bottom:0;z-index:40;display:grid;grid-template-columns:repeat(3,1fr);gap:8px;padding:8px 12px calc(8px + env(safe-area-inset-bottom));background:var(--bar);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);border-top:1px solid var(--line)}
.tab{border:1px solid var(--line);border-radius:14px;padding:12px 4px;font-size:calc(14px*var(--s));font-weight:900;color:var(--mute);background:var(--card2);min-height:54px;-webkit-tap-highlight-color:transparent}
.tab.on{background:var(--acc);color:var(--acctx);border-color:var(--acc)}
.h2{margin:0 0 12px;font-size:calc(12.5px*var(--s));letter-spacing:.08em;font-weight:800;text-transform:uppercase;color:var(--mute)}
.h2.tg{cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:10px;user-select:none;-webkit-tap-highlight-color:transparent}
.h2.tg::after{content:'▾';font-size:15px;opacity:.7;transition:transform .2s}
.card.closed .h2.tg::after{transform:rotate(-90deg)}
.card.closed>*:not(.h2){display:none}
.card.closed .h2{margin-bottom:0}
.hint{margin:-6px 0 10px;font-size:11.5px;color:var(--mute)}
/* Mỗi line một thẻ: tên line + hai nút Khuôn | Label */
.grid{display:grid;grid-template-columns:1fr;gap:10px}
.cell{display:flex;align-items:center;gap:10px;padding:8px 10px;border:1px solid var(--line);border-radius:14px;background:var(--card2);transition:opacity .2s}
.cell.lastok{opacity:.72}
.cell.lastng{border-color:var(--ng);background:var(--ngbg);box-shadow:0 0 0 1px var(--ng)}
.chead{display:flex;align-items:center;gap:2px;min-width:104px;flex:none}
.chead b{font-size:calc(15px*var(--s));color:var(--text)}
.pin{width:34px;height:34px;display:flex;align-items:center;justify-content:center;font-size:19px;line-height:1;color:var(--mute);cursor:pointer;-webkit-tap-highlight-color:transparent;user-select:none}
.cell.pinned .pin{color:var(--warn)}
.pair{display:grid;grid-template-columns:1fr 1fr;gap:8px;flex:1;min-width:0}
.btn{border:1px solid var(--line);border-radius:12px;padding:8px 6px;font-size:calc(14px*var(--s));font-weight:800;color:var(--text);background:var(--btn);min-height:calc(52px*var(--s));-webkit-user-select:none;user-select:none;-webkit-touch-callout:none;-webkit-tap-highlight-color:transparent}
.btn:active{background:var(--btnh)}
.btn .bt{display:block}
.btn .lastv{display:block;margin-top:3px;font-size:calc(11px*var(--s));font-weight:700;color:var(--mute)}
.btn .lastv.okt{color:var(--ok)}
.btn .lastv.ngt{color:var(--ng)}
.btn.okbtn{background:var(--ok)!important;border-color:var(--ok);color:#fff}
.btn.ngbtn{background:var(--ng)!important;border-color:var(--ng);color:#fff;animation:blink 1s infinite}
.btn.okbtn .lastv,.btn.ngbtn .lastv{color:#fff}
.btn.checking{background:var(--warn)!important;border-color:var(--warn);color:#1f2937!important}
.btn.checking .lastv{color:#1f2937}
.btn.active{display:flex;align-items:center;justify-content:center;gap:6px;font-size:11.5px;line-height:1.2}
.btn:disabled{opacity:.4}
.nob{display:flex;align-items:center;justify-content:center;border:1px dashed var(--line);border-radius:12px;color:var(--mute);font-size:12px;min-height:calc(52px*var(--s))}
@keyframes blink{50%{opacity:.55}}
.spin{width:16px;height:16px;border-radius:50%;border:2px solid #fff6;border-top-color:#fff;animation:sp .7s linear infinite;flex:none}
@keyframes sp{to{transform:rotate(360deg)}}
/* Kết quả */
#box{display:flex;align-items:center;gap:14px;padding:16px;border-radius:16px;border:1px solid var(--line);min-height:96px;background:var(--card2);color:var(--text);transition:.25s}
#box.ok{background:var(--okbg);border-color:var(--ok);color:var(--oktx)}
#box.ng{background:var(--ngbg);border-color:var(--ng);color:var(--ngtx);animation:blink 1.4s infinite}
#box.wait{background:var(--warnbg);border-color:var(--warn);color:var(--warntx);font-size:16px;font-weight:800}
.rs-ic{flex:none;width:56px;height:56px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:28px;font-weight:900}
#box.ok .rs-ic{background:rgba(34,197,94,.18);color:var(--ok)}
#box.ng .rs-ic{background:rgba(244,63,94,.18);color:var(--ng)}
#box.wait .rs-ic{background:rgba(234,179,8,.18)}
.rs-ic .spin{width:26px;height:26px;border-width:3px;border-color:rgba(234,179,8,.35);border-top-color:var(--warn)}
.rs-main{min-width:0;flex:1}
.rs-cap{font-size:10.5px;letter-spacing:.14em;color:var(--mute);font-weight:800}
.rs-title{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:3px 0;font-size:calc(26px*var(--s));font-weight:900;color:var(--text)}
.rs-badge{padding:3px 12px;border-radius:999px;font-size:12px;letter-spacing:.08em;color:#fff}
#box.ok .rs-badge{background:var(--ok)}
#box.ng .rs-badge{background:var(--ng)}
#box.wait .rs-badge{background:var(--warn);color:#1f2937}
.rs-msg{font-size:calc(14px*var(--s));font-weight:600;opacity:.9}
.rs-chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}
#rsTime{margin-top:10px;font-size:12px;color:var(--mute)}
/* Chip số xe và tên: trung tính, chữ lớn để dễ đọc */
.chip{display:inline-flex;align-items:center;gap:6px;padding:3px 12px;border-radius:999px;font-size:calc(12px*var(--s));font-weight:800;border:1px solid var(--line);background:var(--card);color:var(--text);white-space:nowrap}
.chip i{font-style:normal;font-weight:700;opacity:.7}
.chip.xe{font-size:calc(18px*var(--s));font-weight:900;padding:5px 16px;border-color:var(--text)}
.chip.ten{font-size:calc(15px*var(--s))}
.rs-chips .chip.xe{font-size:calc(22px*var(--s))}
.item .chip.xe{font-size:calc(14px*var(--s));padding:3px 12px;border-color:var(--line)}
.item .chip.ten{font-size:calc(12px*var(--s));padding:2px 10px}
/* Hộp kết quả (popup) */
.modal-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(2,6,15,.82);display:flex;align-items:center;justify-content:center;z-index:999;padding:20px}
.modal-overlay.hidden{display:none}
.modal-card{width:100%;max-width:460px;background:var(--card);border:2px solid var(--line);border-radius:26px;padding:28px 22px;text-align:center;color:var(--text)}
.modal-card.ok{border-color:var(--ok)}
.modal-card.ng{border-color:var(--ng);animation:blink 1s infinite}
.modal-status{font-size:calc(44px*var(--s));font-weight:900;margin-bottom:8px}
.modal-card.ok .modal-status{color:var(--ok)}
.modal-card.ng .modal-status{color:var(--ng)}
.modal-text{font-size:calc(19px*var(--s));font-weight:700;margin-bottom:12px}
.mchips{display:flex;justify-content:center;flex-wrap:wrap;gap:10px;margin:0 0 12px;min-height:8px}
.mchips .chip.xe{font-size:calc(28px*var(--s));padding:8px 22px}
.mchips .chip.ten{font-size:calc(20px*var(--s));padding:6px 16px}
.modal-time{color:var(--mute);font-size:13px;margin-bottom:14px}
.modal-mail{font-size:14px;font-weight:700;color:var(--warn);margin:0 0 14px}
.modal-auto{font-size:12px;color:var(--mute);margin:0 0 12px}
.modal-close{width:100%;padding:18px;border-radius:16px;border:0;font-size:17px;font-weight:900;color:var(--acctx);background:var(--acc);min-height:56px}
/* Lịch sử */
.hist{max-height:340px;overflow:auto;padding-right:4px;scrollbar-width:thin;scrollbar-color:#64748b transparent}
.hist::-webkit-scrollbar{width:6px}
.hist::-webkit-scrollbar-thumb{background:#64748b;border-radius:6px}
.hh,.item{display:grid;grid-template-columns:86px 100px 1fr 58px;gap:10px;align-items:center}
.hh{padding:0 14px 8px;font-size:10px;letter-spacing:.12em;color:var(--mute);font-weight:800;position:sticky;top:0;background:var(--card);z-index:2}
.hh span:last-child{text-align:right}
.item{padding:10px 14px;margin-bottom:6px;background:var(--card2);border:1px solid var(--line);border-left:4px solid var(--ok);border-radius:12px;font-size:13px;color:var(--text)}
.item.bad{border-left-color:var(--ng);background:var(--ngbg)}
.tm b{display:block;font-size:13px;font-variant-numeric:tabular-nums}
.tm small{color:var(--mute);font-size:11px}
.kd{justify-self:start;padding:2px 8px;border-radius:7px;font-size:10.5px;font-weight:800;letter-spacing:.04em;white-space:nowrap;color:var(--mute);border:1px solid var(--line);background:var(--card)}
.ln{display:flex;flex-wrap:wrap;align-items:center;gap:6px;min-width:0}
.ln b{font-size:calc(14px*var(--s))}
.ln .all{color:var(--mute);font-size:12px}
.st{justify-self:end;padding:5px 12px;border-radius:999px;font-size:calc(12px*var(--s));font-weight:900;letter-spacing:.05em}
.st.p{background:var(--okbg);color:var(--ok);border:1px solid var(--ok)}
.st.f{background:var(--ngbg);color:var(--ng);border:1px solid var(--ng)}
.empty{padding:26px;text-align:center;color:var(--mute);font-size:13px}
.dsep{margin:12px 2px 6px;font-size:11px;letter-spacing:.12em;font-weight:800;color:var(--mute);text-transform:uppercase;display:flex;align-items:center;gap:10px}
.dsep::after{content:"";flex:1;height:1px;background:var(--line)}
.hftog{display:none;width:100%;margin-bottom:8px;padding:11px;border-radius:12px;border:1px solid var(--line);background:var(--card2);color:var(--text);font-size:13px;font-weight:800}
.hf{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
.hf input{flex:1 1 100%;background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:11px 14px;color:var(--text);font-size:14px;outline:none}
.hf input:focus{border-color:var(--text)}
.fb{border:1px solid var(--line);background:var(--card2);color:var(--mute);border-radius:999px;padding:7px 14px;font-size:12px;font-weight:800}
.fb.on{background:var(--acc);color:var(--acctx);border-color:var(--acc)}
.fb.ngb.on{background:var(--ng);color:#fff;border-color:var(--ng)}
.more{display:none;width:100%;margin-top:6px;padding:12px;border-radius:12px;border:1px solid var(--line);background:var(--card2);color:var(--text);font-size:13px;font-weight:800}
/* Cảnh báo trạm mất kết nối lâu */
.espwarn{display:none;align-items:center;gap:12px;margin-bottom:10px;padding:14px 16px;border-radius:16px;background:var(--ngbg);border:1px solid var(--ng);box-shadow:inset 4px 0 0 var(--ng);color:var(--ngtx);animation:blink 1.8s infinite}
.espwarn.show{display:flex}
.espwarn .wi{flex:none;font-size:26px}
.espwarn b{display:block;font-size:14px;margin-bottom:2px}
.espwarn span{font-size:12.5px;opacity:.9;line-height:1.45}
/* Nhấp nháy viền toàn màn hình khi có kết quả */
.flash{position:fixed;top:0;left:0;right:0;bottom:0;pointer-events:none;z-index:1000;opacity:0}
.flash.fok{box-shadow:inset 0 0 0 10px #22c55e,inset 0 0 90px 24px rgba(34,197,94,.6);animation:flOk .8s ease-out 3}
.flash.fng{box-shadow:inset 0 0 0 14px #f43f5e,inset 0 0 120px 34px rgba(244,63,94,.65);animation:flNg .7s ease-in-out infinite}
@keyframes flOk{0%{opacity:0}25%{opacity:1}100%{opacity:0}}
@keyframes flNg{0%,100%{opacity:1}50%{opacity:.12}}
/* Màn hình giám sát */
#mon{display:none}
body.monmode #mon{display:block}
body.monmode .cols{display:none}
.mon-top{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px}
.mon-top .h2{margin:0}
.mon-od{font-size:12px;color:var(--mute);font-weight:700}
.mon-od select{margin-left:6px;background:var(--card2);color:var(--text);border:1px solid var(--line);border-radius:8px;padding:5px 8px;font-size:12px;font-weight:700}
.mon-sum{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:6px}
.pill{background:var(--card2);border:1px solid var(--line);border-radius:999px;padding:7px 12px;font-size:calc(13px*var(--s));font-weight:800}
.mgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:10px;margin-bottom:6px}
.tile{border-radius:14px;padding:12px 8px;text-align:center;border:1px solid var(--line);min-height:104px;display:flex;flex-direction:column;justify-content:center;gap:4px;background:var(--card2);color:var(--mute)}
.tile b{font-size:calc(17px*var(--s));color:var(--text)}
.tile .ts{font-size:calc(13px*var(--s));font-weight:900;letter-spacing:.08em}
.tile .tt{font-size:calc(12px*var(--s));opacity:.95}
.tile.tok{background:var(--okbg);border-color:var(--ok);color:var(--oktx)}
.tile.tng{background:var(--ngbg);border-color:var(--ng);color:var(--ngtx);animation:blink 1.2s infinite}
.tile.tod{background:var(--warnbg);border-color:var(--warn);color:var(--warntx)}
.time{margin-top:6px;font-size:12px;color:var(--mute)}
@media(max-width:560px){.hh{display:none}.item{grid-template-columns:76px 1fr 54px;grid-template-areas:"tm ln st" "tm kd kd";row-gap:6px}.tm{grid-area:tm}.kd{grid-area:kd}.ln{grid-area:ln}.st{grid-area:st}}
@media(max-width:700px){.hftog{display:block}.hf{display:none}.hf.open{display:flex}.more{display:block}.hist{max-height:none;overflow:visible}.hist.exp{max-height:420px;overflow:auto}}
@media(min-width:1000px){
  body{padding-bottom:0}
  .wrap{max-width:1240px}
  .tabs{position:sticky;top:0;bottom:auto;left:auto;right:auto;padding:8px 0;margin-bottom:10px;border-top:0;background:var(--bg);-webkit-backdrop-filter:none;backdrop-filter:none}
  .cols{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,1fr);gap:16px;align-items:start}
  .colR{position:sticky;top:84px}
  .hist{max-height:calc(100vh - 470px);min-height:240px}
}
@media(min-width:1500px){.grid{grid-template-columns:1fr 1fr}.mgrid{grid-template-columns:repeat(auto-fill,minmax(180px,1fr))}.tile{min-height:130px}.tile b{font-size:24px}.tile .ts{font-size:16px}.tile .tt{font-size:14px}}
</style>
</head>
<body>
<div class="flash" id="flash"></div>
<div class="wrap">
  <div class="espwarn" id="espWarn"><div class="wi">⚠️</div><div><b id="espWarnT"></b><span id="espWarnS"></span></div></div>
  <div class="top">
    <div class="top1">
      <div><span class="brand">SCAN SYSTEM 2.1</span><h1>Kiểm tra mã khuôn</h1></div>
      <div class="tools">
        <button class="tbtn" id="thBtn" onclick="togTheme()" title="Chế độ sáng / tối">☀</button>
        <button class="tbtn" id="szBtn" onclick="togBig()" title="Phóng to / thu nhỏ chữ">A+</button>
      </div>
    </div>
    <div class="sbar">
      <span class="sp off" id="esp"><span class="dot"></span><span id="espText">Trạm: mất kết nối</span></span>
      <span class="sp" id="stat"><b class="okt">OK 0</b><span class="sep">/</span><b class="ngt">NG 0</b></span>
      <span class="sp" id="cd" style="display:none" title="Đếm ngược tự động kiểm tra toàn bộ"></span>
    </div>
  </div>
  <div class="tabs">
    <button class="tab on" id="tabSemi" onclick="setMode('semi')">SEMI LINE</button>
    <button class="tab" id="tabAuto" onclick="setMode('auto')">AUTO LINE</button>
    <button class="tab" id="tabMon" onclick="setMode('mon')">GIÁM SÁT</button>
  </div>
  <div class="cols">
    <div class="colL">
      <div class="card">
        <p class="h2 tg" id="hLines" onclick="togCard(this)" title="Bấm để thu gọn / mở rộng">Chọn line — SEMI (VS)</p>
        <div class="hint">Bấm ☆ hoặc giữ nút để ghim line hay dùng lên đầu</div>
        <div class="grid" id="grid"></div>
      </div>
    </div>
    <div class="colR">
      <div class="card">
        <p class="h2">Kết quả</p>
        <div id="box" class="wait">Chưa kiểm tra</div>
        <div id="rsTime"></div>
      </div>
      <div class="card" id="histCard">
        <p class="h2 tg" id="hHist" onclick="togCard(this)" title="Bấm để thu gọn / mở rộng">Lịch sử</p>
        <button class="hftog" id="hfTog" onclick="togHf()">🔍 Tìm / lọc</button>
        <div class="hf" id="hf">
          <input id="q" placeholder="Tìm line, số xe, tên..." oninput="renderHist()">
          <button class="fb on" data-k="" onclick="setKind('')">Tất cả</button>
          <button class="fb" data-k="KHUON" onclick="setKind('KHUON')">Mã khuôn</button>
          <button class="fb" data-k="LABEL" onclick="setKind('LABEL')">Label</button>
          <button class="fb ngb" data-k="NG" onclick="setKind('NG')">Chỉ NG</button>
        </div>
        <div class="hist" id="hist"></div>
        <button class="more" id="hMore" onclick="togMore()"></button>
      </div>
    </div>
  </div>
  <div class="card" id="mon">
    <div class="mon-top">
      <p class="h2">Giám sát tất cả line</p>
      <label class="mon-od">Báo quá hạn sau<select id="odSel" onchange="setOd(this.value)">
        <option value="1">1 giờ</option><option value="2">2 giờ</option><option value="4">4 giờ</option><option value="8">8 giờ</option>
      </select></label>
    </div>
    <div class="mon-sum" id="monSum"></div>
    <div id="monBody"></div>
  </div>
</div>
<div class="modal-overlay hidden" id="resultModal">
  <div class="modal-card" id="modalCard">
    <div class="modal-status" id="modalStatus">OK</div>
    <div class="modal-text" id="modalText"></div>
    <div class="mchips" id="modalChips"></div>
    <div class="modal-time" id="modalTime"></div>
    <div class="modal-mail" id="modalMail"></div>
    <div class="modal-auto" id="modalAuto"></div>
    <button class="modal-close" onclick="closeResultModal()">ĐÓNG</button>
  </div>
</div>
<script>
var busy=false, mode='semi', allMold=[], allLabel=[], moldLoaded=false, labelLoaded=false, pairs=[];
var lastHist=[], histKind='', histOpen=false, histAll=false;
var pins={}, lastResKey=null, flashT=null, lastMonHtml='', busyTimer=null, modalT=null, modalCtx=null, OD_HOURS=2;
function isSemi(n){ return String(n).indexOf('VS')>=0; }
function isAuto(n){ return String(n).indexOf('VA')>=0; }
function isPhone(){ try{ return !!(window.matchMedia&&window.matchMedia('(max-width:700px)').matches); }catch(e){ return false; } }
var KIND={'KHUON':'KHUÔN','LABEL':'LABEL','KHUON-ALL':'TẤT CẢ · KHUÔN','LABEL-ALL':'TẤT CẢ · LABEL'};
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function dayStr(d){function z(n){return (n<10?'0':'')+n;}return z(d.getDate())+'/'+z(d.getMonth()+1)+'/'+d.getFullYear();}
function dayLabel(ds){
  var now=new Date(), y=new Date(now.getTime()-86400000);
  if(ds===dayStr(now)) return 'Hôm nay · '+ds;
  if(ds===dayStr(y)) return 'Hôm qua · '+ds;
  return ds;
}
function tmShort(r,today){
  var t=String(r.time||'').split(' ');
  return String(t[0]||'').slice(0,5)+((t[1]&&t[1]!==today)?(' '+t[1].slice(0,5)):'');
}
function paintResult(j,H){
  var st=j.light==='YELLOW'?'wait':(j.status==='NG'?'ng':'ok');
  var txt=String(j.text||'');
  var m=/(Line|Label)[ ]+([A-Za-z0-9-]+)([.,])/.exec(txt);
  var title,cap;
  if(m&&m[3]==='.'){title=m[2];cap=m[1]==='Label'?'KIỂM TRA LABEL':'KIỂM TRA MÃ KHUÔN';}
  else if(txt.indexOf('Cảnh báo')>=0){title='Nhiều line không đạt';cap='KIỂM TRA TỰ ĐỘNG';}
  else{title='Toàn bộ hệ thống';cap='KIỂM TRA TỰ ĐỘNG';}
  var label={ok:'OK',ng:'NG',wait:'ĐANG KIỂM TRA'}[st];
  var icon={ok:'✓',ng:'✕',wait:'<span class="spin"></span>'}[st];
  var h0=(H||[])[0], chips='';
  if(h0&&h0.text===txt){
    if(h0.xe) chips+='<span class="chip xe"><i>Xe</i>'+esc(h0.xe)+'</span>';
    if(h0.ten) chips+='<span class="chip ten"><i>👤</i>'+esc(h0.ten)+'</span>';
  }
  var b=document.getElementById('box');
  b.className=st;
  b.innerHTML='<div class="rs-ic">'+icon+'</div><div class="rs-main"><div class="rs-cap">'+cap+'</div>'+
    '<div class="rs-title">'+esc(title)+'<span class="rs-badge">'+label+'</span></div>'+
    '<div class="rs-msg">'+esc(txt)+'</div>'+(chips?'<div class="rs-chips">'+chips+'</div>':'')+'</div>';
  document.getElementById('rsTime').innerText='🕒 '+(j.time||'');
}
var ESP_WARN_SEC=300;   // quá 5 phút không có tín hiệu thì hiện cảnh báo
function paintEspWarn(e){
  var box=document.getElementById('espWarn');
  if(!e||e.online){box.classList.remove('show');return;}
  var sec=(typeof e.sec==='number')?e.sec:-1, t='', sub='Đèn tháp và loa không hoạt động. Kiểm tra nguồn điện và WiFi của trạm.';
  if(sec<0){t='Chưa nhận được tín hiệu từ Trạm cảnh báo';}
  else if(sec>=ESP_WARN_SEC){
    var m=Math.floor(sec/60), h=Math.floor(m/60);
    t='Trạm cảnh báo mất kết nối '+(h>0?(h+' giờ '+(m%60)+' phút'):(m+' phút'));
  } else {box.classList.remove('show');return;}
  document.getElementById('espWarnT').innerText=t;
  document.getElementById('espWarnS').innerText=sub;
  box.classList.add('show');
}
/* Nhấp nháy viền toàn màn hình: OK = xanh lóe 3 lần, NG = đỏ nhấp nháy ~12 giây (bấm ĐÓNG để tắt sớm) */
function stopFlash(){
  var f=document.getElementById('flash'); if(!f) return;
  if(flashT){clearTimeout(flashT);flashT=null;}
  f.className='flash';
}
function doFlash(st){
  var f=document.getElementById('flash'); if(!f) return;
  stopFlash();
  void f.offsetWidth;
  f.className='flash '+(st==='NG'?'fng':'fok');
  flashT=setTimeout(stopFlash,st==='NG'?12000:2800);
}
function fillModal(d){
  var m=document.getElementById('resultModal');
  if(m.classList.contains('hidden')||!modalCtx||modalCtx.done) return;
  var h0=(d.history||[])[0];
  if(h0&&h0.text===modalCtx.text&&h0.line===modalCtx.sheet){
    var c='';
    if(h0.xe) c+='<span class="chip xe"><i>🚚</i>'+esc(h0.xe)+'</span>';
    if(h0.ten) c+='<span class="chip ten"><i>👤</i>'+esc(h0.ten)+'</span>';
    document.getElementById('modalChips').innerHTML=c;
    if(h0.time) document.getElementById('modalTime').innerText=h0.time;
    modalCtx.done=true;
  }
}
function paintDash(d){
  if(!d) return;
  var on=!!(d.esp&&d.esp.online);
  var e=document.getElementById('esp');
  e.className='sp '+(on?'on':'off');
  e.title='Tín hiệu gần nhất: '+((d.esp&&d.esp.last&&d.esp.last!=='Chưa kết nối')?d.esp.last:'chưa nhận tín hiệu');
  document.getElementById('espText').innerText=on?'Trạm: hoạt động':'Trạm: mất kết nối';
  paintEspWarn(d.esp);
  var a=d.auto||{}, parts=[];
  if(a.moldSec>0) parts.push('Khuôn '+a.moldSec+'s');
  if(a.labelSec>0) parts.push('Label '+a.labelSec+'s');
  var cd=document.getElementById('cd');
  if(parts.length){ cd.innerText='⏱ '+parts.join(' · '); cd.style.display=''; } else { cd.style.display='none'; }
  var s=d.stats||{};
  document.getElementById('stat').innerHTML='<b class="okt">OK '+(s.ok||0)+'</b><span class="sep">/</span><b class="ngt">NG '+(s.ng||0)+'</b>';
  var j=d.latest;
  if(j&&j.status!=='none'){
    paintResult(j,d.history);
    var lightKey=j.light||(j.status==='NG'?'RED':(j.status==='OK'?'GREEN':'YELLOW'));
    var rkey=String(j.time||'')+'|'+String(j.text||'');
    if(lastResKey===null){ lastResKey=rkey; }
    else if(rkey!==lastResKey){
      lastResKey=rkey;
      if(lightKey!=='YELLOW') doFlash(j.status==='NG'?'NG':'OK');
    }
  }
  lastHist=d.history||[];
  fillModal(d);
  renderHist();
  updateLineBadges();
  if(mode==='mon') paintMonitor();
}
function lastMap(){
  var last={};
  for(var i=0;i<lastHist.length;i++){var r=lastHist[i]; if(!r.line) continue; var key=r.kind+'|'+r.line; if(!(key in last)) last[key]=r;}
  return last;
}
/* ===== Thẻ line: ghép mã khuôn và label cùng line (VA015-A với LB-VA015A...) ===== */
function nrm(s){ return String(s).replace(/-/g,''); }
function labKey(n){ n=String(n); if(n.indexOf('LB-')===0) n=n.slice(3); return nrm(n); }
function labTitle(n){ n=String(n); return n.indexOf('LB-')===0?n.slice(3):n; }
function buildPairs(){
  var out=[], usedL={}, matched=[];
  allMold.forEach(function(m,i){
    matched[i]=-1; var k=nrm(m);
    for(var j=0;j<allLabel.length;j++){ if(!usedL[j]&&labKey(allLabel[j])===k){ matched[i]=j; usedL[j]=1; break; } }
  });
  allMold.forEach(function(m,i){            // lượt 2: ghép theo tiền tố duy nhất (VS051 với LB-VS051B)
    if(matched[i]>=0) return;
    var k=nrm(m), c=[];
    for(var j=0;j<allLabel.length;j++){ if(!usedL[j]&&labKey(allLabel[j]).indexOf(k)===0) c.push(j); }
    if(c.length===1){ matched[i]=c[0]; usedL[c[0]]=1; }
  });
  allMold.forEach(function(m,i){ out.push({title:m,mold:m,label:matched[i]>=0?allLabel[matched[i]]:''}); });
  allLabel.forEach(function(l,j){ if(!usedL[j]) out.push({title:labTitle(l),mold:'',label:l}); });
  out.sort(function(a,b){ return a.title<b.title?-1:(a.title>b.title?1:0); });
  return out;
}
function btnInner(b,r,today){
  var h='<span class="bt">'+esc(b.dataset.lab)+'</span>';
  if(r){ var ng=r.status==='NG'; h+='<span class="lastv '+(ng?'ngt':'okt')+'">'+(ng?'NG':'OK')+' · '+esc(tmShort(r,today))+'</span>'; }
  else h+='<span class="lastv">chưa có</span>';
  return h;
}
function cellRank(c){ return (c.classList.contains('lastng')?0:2)+(pins[c.dataset.key]?0:1); }   // NG lên đầu, rồi đến line đã ghim
function orderCards(force){
  if(busy&&!force) return;
  var g=document.getElementById('grid'); if(!g) return;
  var cells=Array.prototype.slice.call(g.children);
  var sorted=cells.slice().sort(function(a,b){
    var ra=cellRank(a), rb=cellRank(b);
    return ra!==rb?ra-rb:(Number(a.dataset.idx)-Number(b.dataset.idx));
  });
  for(var i=0;i<cells.length;i++){ if(sorted[i]!==cells[i]){ sorted.forEach(function(c){ g.appendChild(c); }); return; } }
}
function updateLineBadges(){
  var g=document.getElementById('grid'); if(!g) return;
  var last=lastMap(), today=dayStr(new Date());
  Array.prototype.forEach.call(g.children,function(c){
    var ng=false, any=false;
    Array.prototype.forEach.call(c.querySelectorAll('button.btn'),function(b){
      var r=last[b.dataset.kind+'|'+b.dataset.name];
      if(r){ any=true; if(r.status==='NG') ng=true; }
      if(b.classList.contains('active')) return;   // đang kiểm tra: giữ nguyên chữ ĐANG KIỂM TRA
      var h=btnInner(b,r,today); if(b.innerHTML!==h) b.innerHTML=h;
    });
    c.classList.toggle('lastng',ng);
    c.classList.toggle('lastok',any&&!ng);
  });
  orderCards(false);
}
function savePins(){ try{localStorage.setItem('pinCards',JSON.stringify(pins));}catch(e){} }
function applyPins(){
  var g=document.getElementById('grid'); if(!g) return;
  Array.prototype.forEach.call(g.children,function(c){
    var on=!!pins[c.dataset.key];
    c.classList.toggle('pinned',on);
    var p=c.querySelector('.pin'); if(p) p.innerText=on?'★':'☆';
  });
  orderCards(true);
}
function togPin(key){
  if(pins[key]) delete pins[key]; else pins[key]=1;
  savePins(); applyPins();
}
function bindPress(b,key,tap){
  var pressT=null, longDone=false;
  function start(){
    longDone=false;
    if(pressT) clearTimeout(pressT);
    pressT=setTimeout(function(){
      pressT=null; longDone=true; togPin(key);
      try{ if(navigator.vibrate) navigator.vibrate(30); }catch(e){}
    },600);
  }
  function cancel(){ if(pressT){clearTimeout(pressT);pressT=null;} }
  b.addEventListener('touchstart',start,{passive:true});
  b.addEventListener('touchend',cancel);
  b.addEventListener('touchmove',cancel);
  b.addEventListener('touchcancel',cancel);
  b.addEventListener('mousedown',start);
  b.addEventListener('mouseup',cancel);
  b.addEventListener('mouseleave',cancel);
  b.addEventListener('contextmenu',function(ev){ev.preventDefault();});
  b.onclick=function(){
    if(longDone){longDone=false;return;}   // vừa giữ để ghim thì không chạy kiểm tra
    tap();
  };
}
function makeBtn(cell,name,fn,lab,kind,cls){
  var b=document.createElement('button');
  b.className='btn'+(cls?(' '+cls):'');
  b.dataset.name=name; b.dataset.lab=lab; b.dataset.kind=kind;
  b.innerHTML=btnInner(b,null,'');
  bindPress(b,cell.dataset.key,function(){ run(fn,name,b); });
  return b;
}
function drawCards(){
  var g=document.getElementById('grid'); if(!g) return;
  g.innerHTML='';
  var inMode=(mode==='auto')?isAuto:isSemi;
  pairs.filter(function(p){ return inMode(p.title); }).forEach(function(p,i){
    var cell=document.createElement('div');
    cell.className='cell'; cell.dataset.idx=i; cell.dataset.key='C|'+nrm(p.title);
    var head=document.createElement('div'); head.className='chead';
    var t=document.createElement('b'); t.innerText=p.title;
    var pin=document.createElement('span'); pin.className='pin'; pin.setAttribute('role','button'); pin.title='Ghim lên đầu'; pin.innerText='☆';
    pin.onclick=function(ev){ev.stopPropagation();togPin(cell.dataset.key);};
    head.appendChild(t); head.appendChild(pin);
    var pair=document.createElement('div'); pair.className='pair';
    if(p.mold) pair.appendChild(makeBtn(cell,p.mold,'checkOneSheetHeadless','Khuôn','KHUON','')); else { var n1=document.createElement('span'); n1.className='nob'; n1.innerText='Không có khuôn'; pair.appendChild(n1); }
    if(p.label) pair.appendChild(makeBtn(cell,p.label,'checkOneLabelHeadless','Label','LABEL','lb')); else { var n2=document.createElement('span'); n2.className='nob'; n2.innerText='Không có label'; pair.appendChild(n2); }
    cell.appendChild(head); cell.appendChild(pair);
    g.appendChild(cell);
  });
  applyPins();
  updateLineBadges();
}
/* ===== Lịch sử: trên điện thoại mặc định chỉ 5 dòng, bộ lọc ẩn đến khi mở ===== */
function setKind(k){
  histKind=k;
  document.querySelectorAll('.fb').forEach(function(b){b.classList.toggle('on',b.dataset.k===k);});
  renderHist();
}
function togHf(){ histOpen=!histOpen; document.getElementById('hf').classList.toggle('open',histOpen); renderHist(); }
function togMore(){ histAll=!histAll; renderHist(); }
function renderHist(){
  var box=document.getElementById('hist'), top=box.scrollTop;
  var q=(document.getElementById('q').value||'').toLowerCase().trim();
  var inMode=(mode==='auto'?isAuto:isSemi);
  var all=lastHist.filter(function(r){
    if(r.line && !inMode(r.line)) return false;
    if(histKind==='NG' && r.status!=='NG') return false;
    if((histKind==='KHUON'||histKind==='LABEL') && r.kind!==histKind) return false;
    if(q && (r.line+' '+r.xe+' '+r.ten+' '+r.text).toLowerCase().indexOf(q)<0) return false;
    return true;
  }).slice(0,40);
  var phone=isPhone(), lim=(phone&&!histAll)?5:40, H=all.slice(0,lim), prevDay='';
  var semi=mode!=='auto';
  document.getElementById('hHist').innerText='Lịch sử — '+(semi?'SEMI (VS)':'AUTO (VA)')+' · '+all.length+' kết quả';
  var active=!!(q||histKind);
  document.getElementById('hfTog').innerText=(histOpen?'▴ Ẩn bộ lọc':'🔍 Tìm / lọc')+((active&&!histOpen)?' •':'');
  box.classList.toggle('exp',phone&&histAll);
  var more=document.getElementById('hMore');
  if(phone&&all.length>5){ more.style.display='block'; more.innerText=histAll?'Thu gọn':('Xem thêm ('+(all.length-5)+')'); }
  else more.style.display='none';
  box.innerHTML=H.length?
    '<div class="hh"><span>THỜI GIAN</span><span>LOẠI</span><span>LINE / CHI TIẾT</span><span>KẾT QUẢ</span></div>'+
    H.map(function(r){
      var t=String(r.time||'').split(' ');
      var dsep=''; if((t[1]||'')!==prevDay){prevDay=t[1]||''; dsep='<div class="dsep">'+esc(dayLabel(prevDay))+'</div>';}
      var kd=KIND[r.kind]||r.kind||'--';
      var bad=r.status==='NG';
      var det=r.line?'<b>'+esc(r.line)+'</b>':'<span class="all">'+esc(r.text)+'</span>';
      var xe=r.xe?'<span class="chip xe"><i>🚚</i>'+esc(r.xe)+'</span>':'';
      var ten=r.ten?'<span class="chip ten"><i>👤</i>'+esc(r.ten)+'</span>':'';
      return dsep+'<div class="item'+(bad?' bad':'')+'"><div class="tm"><b>'+esc(t[0])+'</b><small>'+esc(t[1]||'')+'</small></div>'+
        '<span class="kd">'+esc(kd)+'</span><div class="ln">'+det+xe+ten+'</div>'+
        '<span class="st '+(bad?'f':'p')+'">'+esc(r.status)+'</span></div>';
    }).join('')
    :'<div class="empty">Không có kết quả phù hợp</div>';
  box.scrollTop=top;
}
/* ===== Chạy kiểm tra ===== */
function markBtn(btn,status){
  if(!btn) return;
  btn.classList.remove('okbtn','ngbtn','checking');
  if(btn._resetTimer){clearTimeout(btn._resetTimer);btn._resetTimer=null;}
  if(status==='OK') btn.classList.add('okbtn');
  if(status==='NG') btn.classList.add('ngbtn');
  if(status==='OK'||status==='NG'){
    btn._resetTimer=setTimeout(function(){btn.classList.remove('okbtn','ngbtn');btn._resetTimer=null;},5000);
  }
}
function setBusy(v,btn){
  busy=v;
  if(busyTimer){clearInterval(busyTimer);busyTimer=null;}
  document.querySelectorAll('.btn').forEach(function(b){b.disabled=v;});
  if(btn){
    if(v){
      btn.classList.remove('okbtn','ngbtn');
      btn.classList.add('active','checking');
      btn.innerHTML='<span class="spin"></span><span>ĐANG KIỂM TRA</span>';
      var t0=Date.now();   // hiển thị số giây đã chờ để người vận hành biết hệ thống vẫn đang chạy
      busyTimer=setInterval(function(){
        var sp=btn.lastChild; if(sp&&sp.nodeType===1) sp.innerText='ĐANG KIỂM TRA · '+Math.floor((Date.now()-t0)/1000)+'s';
      },500);
    } else {
      btn.classList.remove('active','checking');
      btn.innerHTML=btnInner(btn,null,'');
      updateLineBadges();   // điền lại kết quả gần nhất ngay, tránh nháy chữ "chưa có"
    }
  }
}
function closeResultModal(){
  if(modalT){clearTimeout(modalT);modalT=null;}
  document.getElementById('resultModal').classList.add('hidden');
  stopFlash();
}
function openResultModal(j){
  var card=document.getElementById('modalCard');
  card.classList.remove('ok','ng');
  card.classList.add(j.status==='NG'?'ng':'ok');
  modalCtx={text:j.text,sheet:j.sheet,done:false};
  document.getElementById('modalStatus').innerText=j.status||'';
  document.getElementById('modalText').innerText=(j.text||'')+(j.xeReason?' — '+j.xeReason:'');
  document.getElementById('modalChips').innerHTML='';
  document.getElementById('modalTime').innerText=j.time||'';
  document.getElementById('modalMail').innerText=j.emailSent===true?'📧 Đã gửi email cảnh báo đến quản lý':(j.emailSent===false?'⚠️ Chưa gửi được email — kiểm tra quyền gửi mail':'');
  document.getElementById('resultModal').classList.remove('hidden');
  if(modalT){clearTimeout(modalT);modalT=null;}
  var auto=document.getElementById('modalAuto');
  if(j.status!=='NG'){ auto.innerText='Tự đóng sau 3 giây'; modalT=setTimeout(closeResultModal,3000); }   // OK tự đóng, NG giữ lại đến khi bấm ĐÓNG
  else auto.innerText='';
}
function focusResult(){
  var el=document.getElementById('box');
  if(!el) return;
  var card=el.closest?(el.closest('.card')||el):el;
  try{ card.scrollIntoView({behavior:'smooth',block:'center',inline:'nearest'}); }
  catch(e){ try{ card.scrollIntoView(true); }catch(e2){} }
}
function run(fn,name,btn){
  if(busy) return;
  setBusy(true,btn);
  document.getElementById('box').className='wait';
  document.getElementById('box').innerText='SCAN SYSTEM — '+name;
  setTimeout(focusResult,60);
  google.script.run
    .withSuccessHandler(function(j){setBusy(false,btn);markBtn(btn,j.status);focusResult();openResultModal(j);tick();})
    .withFailureHandler(function(e){document.getElementById('box').className='ng';document.getElementById('box').innerText='Lỗi: '+e.message;setBusy(false,btn);})
    [fn](name);
}
/* ===== Màn hình giám sát: ô xanh / đỏ, chuyển vàng khi line quá hạn chưa kiểm tra ===== */
function parseT(s){
  var t=String(s||'').split(' '); if(t.length<2) return NaN;
  var h=t[0].split(':'), d=t[1].split('/'); if(h.length<2||d.length<3) return NaN;
  return Date.UTC(Number(d[2]),Number(d[1])-1,Number(d[0]),Number(h[0])-7,Number(h[1]),Number(h[2]||0));   // giờ trong lịch sử là GMT+7
}
function ageText(ms){
  var m=Math.floor((Date.now()-ms)/60000); if(!(m>=0)) m=0;
  if(m<1) return 'vừa xong';
  if(m<60) return m+' phút trước';
  var h=Math.floor(m/60);
  if(h<24) return h+' giờ'+(m%60?(' '+(m%60)+' phút'):'')+' trước';
  return Math.floor(h/24)+' ngày trước';
}
function setOd(v){
  OD_HOURS=Number(v)||2;
  try{localStorage.setItem('odHours',String(OD_HOURS));}catch(e){}
  lastMonHtml=''; paintMonitor();
}
function paintMonitor(){
  var body=document.getElementById('monBody'); if(!body) return;
  var last=lastMap(), today=dayStr(new Date()), c={ok:0,ng:0,od:0,none:0};
  function tiles(names,kind){
    return names.map(function(n){
      var r=last[kind+'|'+n], cls='', st='CHƯA CÓ', tt='—', tip='Không có trong lịch sử gần nhất';
      if(r){
        var bad=r.status==='NG', ms=parseT(r.time);
        var od=!bad&&!isNaN(ms)&&(Date.now()-ms)>OD_HOURS*3600000;
        cls=bad?' tng':(od?' tod':' tok'); st=bad?'NG':(od?'QUÁ HẠN':'OK');
        if(bad) c.ng++; else if(od) c.od++; else c.ok++;
        tt='🕒 '+tmShort(r,today)+(isNaN(ms)?'':(' · '+ageText(ms))); tip=r.time||'';
      } else c.none++;
      return '<div class="tile'+(cls||'')+'" title="'+esc(tip)+'"><b>'+esc(n)+'</b><span class="ts">'+st+'</span><span class="tt">'+esc(tt)+'</span></div>';
    }).join('');
  }
  var groups=[
    ['Mã khuôn · SEMI (VS)',allMold.filter(isSemi),'KHUON'],
    ['Mã khuôn · AUTO (VA)',allMold.filter(isAuto),'KHUON'],
    ['Label · SEMI (VS)',allLabel.filter(isSemi),'LABEL'],
    ['Label · AUTO (VA)',allLabel.filter(isAuto),'LABEL']
  ];
  var html=groups.map(function(g){
    if(!g[1].length) return '';
    return '<div class="dsep">'+esc(g[0])+'</div><div class="mgrid">'+tiles(g[1],g[2])+'</div>';
  }).join('');
  if(html!==lastMonHtml){ body.innerHTML=html; lastMonHtml=html; }
  document.getElementById('monSum').innerHTML=
    '<span class="pill okt">🟢 OK '+c.ok+'</span><span class="pill ngt">🔴 NG '+c.ng+'</span><span class="pill" style="color:var(--warn)">🟡 Quá hạn '+c.od+'</span><span class="pill">⚪ Chưa có '+c.none+'</span>';
}
/* ===== Chế độ, giao diện sáng / tối, cỡ chữ ===== */
function renderMode(){
  var base=(mode==='auto')?'auto':'semi';
  document.body.classList.toggle('monmode',mode==='mon');
  document.getElementById('tabSemi').className='tab'+(mode==='semi'?' on':'');
  document.getElementById('tabAuto').className='tab'+(mode==='auto'?' on':'');
  document.getElementById('tabMon').className='tab'+(mode==='mon'?' on':'');
  document.getElementById('hLines').innerText=base==='semi'?'Chọn line — SEMI (VS)':'Chọn line — AUTO (VA)';
  if(moldLoaded&&labelLoaded){ pairs=buildPairs(); drawCards(); }
  renderHist();
  if(mode==='mon'){ lastMonHtml=''; paintMonitor(); }
}
function setMode(m){
  mode=m;
  try{localStorage.setItem('scanMode',m);}catch(e){}   // máy treo TV tải lại trang vẫn ở chế độ giám sát
  renderMode();
}
function tick(){ google.script.run.withSuccessHandler(paintDash).getDashboard(); }
google.script.run.withSuccessHandler(function(n){ allMold=n||[]; moldLoaded=true; renderMode(); }).withFailureHandler(function(){ moldLoaded=true; renderMode(); }).listSheetNames();
google.script.run.withSuccessHandler(function(n){ allLabel=n||[]; labelLoaded=true; renderMode(); }).withFailureHandler(function(){ labelLoaded=true; renderMode(); }).listLabelSheetNames();
function togBig(){
  var on=document.body.classList.toggle('big');
  document.getElementById('szBtn').innerText=on?'A−':'A+';
  try{localStorage.setItem('bigFont',on?'1':'0');}catch(e){}
}
function togTheme(){
  var light=document.documentElement.classList.toggle('light');
  document.getElementById('thBtn').innerText=light?'🌙':'☀';
  try{localStorage.setItem('theme',light?'light':'dark');}catch(e){}
}
function togCard(h){
  var c=h.parentNode; if(!c) return;
  c.classList.toggle('closed');
  try{
    var st=JSON.parse(localStorage.getItem('cardClosed')||'{}');
    st[h.id]=c.classList.contains('closed')?1:0;
    localStorage.setItem('cardClosed',JSON.stringify(st));
  }catch(e){}   // trình duyệt chặn lưu trữ thì chỉ nhớ trong phiên này
}
(function(){
  try{
    if(localStorage.getItem('bigFont')==='1'){document.body.classList.add('big');document.getElementById('szBtn').innerText='A−';}
  }catch(e){}
  try{
    if(localStorage.getItem('theme')==='light'){document.documentElement.classList.add('light');document.getElementById('thBtn').innerText='🌙';}
  }catch(e){}
  try{
    var st=JSON.parse(localStorage.getItem('cardClosed')||'{}');
    Object.keys(st).forEach(function(id){
      var h=document.getElementById(id);
      if(h&&st[id]&&h.parentNode) h.parentNode.classList.add('closed');
    });
  }catch(e){}
  try{
    var sm=localStorage.getItem('scanMode');
    if(sm==='semi'||sm==='auto'||sm==='mon') mode=sm;
  }catch(e){}
  try{
    var pj=JSON.parse(localStorage.getItem('pinCards')||'{}');
    if(pj&&typeof pj==='object') pins=pj;
  }catch(e){}
  try{
    var oh=Number(localStorage.getItem('odHours'));
    if(oh===1||oh===2||oh===4||oh===8) OD_HOURS=oh;
  }catch(e){}
  document.getElementById('odSel').value=String(OD_HOURS);
})();
tick();
setInterval(tick,1000);
</script>
</body>
</html>
  `).setTitle('SCAN SYSTEM 2.1').addMetaTag('viewport','width=device-width, initial-scale=1');
}


// ====================== [MỚI] PHIẾU LỖI & KHAI BÁO CHỈNH SỬA (có mã nhân viên + mật khẩu) ======================
// Đã gộp sẵn vào file: hàm doGet của giao diện v3 được đổi tên thành doGetUi_, hàm doGet mới ở cuối file bọc bên ngoài.
// Sau khi dán: chạy tay hàm  caiDatSuaLoi  (1 lần), rồi Deploy lại Web App (New version).
//
// CÁCH HOẠT ĐỘNG
//  1. Mỗi lần kiểm tra ra NG, hệ thống TỰ tạo 1 phiếu trong sheet LOI_CHINH_SUA: giờ lỗi, line, số xe, người scan, nội dung lỗi chi tiết.
//  2. Nhân viên vào Web App → tab GIÁM SÁT → bảng "Lỗi & chỉnh sửa" → bấm "Khai báo chỉnh sửa".
//     Nhập mã nhân viên + mật khẩu → bấm các nút cách sửa (thêm nút mới tự do) → Gửi.
//  3. Hệ thống ghi giờ sửa + người sửa, TỰ GỬI EMAIL cho quản lý, và hiện lên bảng giám sát.
//  4. Lần kiểm tra OK kế tiếp của line đó sẽ tự đánh dấu "ĐÃ SỬA · OK". Sửa mà vẫn NG → phiếu cũ thành "SỬA CHƯA ĐẠT" và mở phiếu mới.
// ScanIt / ứng dụng ngoài vẫn ghi dữ liệu vào sheet bình thường: khối này KHÔNG chặn, KHÔNG sửa các sheet dữ liệu.
// Lưu ý: Google Sheet không cho phép bắt nhập mật khẩu khi sửa trực tiếp trên bảng tính. Xác thực nằm ở bước KHAI BÁO trên Web App.
// QUAY LẠI BẢN TRƯỚC: xóa khối này + đổi doGetUi_ về doGet (dòng bản v3) + Deploy lại.

var FX_TICKET = 'LOI_CHINH_SUA';
var FX_STAFF = 'NV_DANH_SACH';
var FX_OPT = 'CACH_SUA_DS';
var FX_ST_OPEN = 'CHỜ SỬA';
var FX_ST_FIXED = 'ĐÃ SỬA - CHỜ KIỂM TRA';
var FX_ST_OK = 'ĐÃ XÁC NHẬN OK';
var FX_ST_BAD = 'SỬA CHƯA ĐẠT';
var FX_MAX_FAIL = 5;          // sai mật khẩu quá số lần này thì khóa mã đó 5 phút
var FX_SESSION_SEC = 1800;    // đăng nhập có hiệu lực 30 phút
var FX_COLS = ['Mã phiếu', 'Thời gian lỗi', 'Loại', 'Line', 'Số xe', 'Người scan', 'Nội dung lỗi', 'Trạng thái',
  'Thời gian sửa', 'Mã NV sửa', 'Tên NV sửa', 'Cách sửa', 'Ghi chú', 'OK lại lúc', 'Email',
  'Dòng lỗi lúc NG (bản sao)', 'Đối chiếu dòng khi khai báo'];



// ---------- Sheet ----------
function fxNow_() { return Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy'); }

function fxSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FX_TICKET);
  if (!sh) {
    sh = ss.insertSheet(FX_TICKET);
    sh.getRange(1, 1, 1, FX_COLS.length).setValues([FX_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    var w = [120, 135, 70, 90, 80, 120, 380, 170, 135, 90, 130, 220, 220, 135, 90];
    for (var i = 0; i < w.length; i++) sh.setColumnWidth(i + 1, w[i]);
    sh.getRange(1, 7, sh.getMaxRows(), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } else if (!String(sh.getRange(1, 16).getDisplayValue() || '').trim()) {   // sheet tạo từ bản trước: thêm 2 cột mới
    sh.getRange(1, 16, 1, 2).setValues([[FX_COLS[15], FX_COLS[16]]]).setFontWeight('bold');
    sh.setColumnWidth(16, 380); sh.setColumnWidth(17, 380);
  }
  return sh;
}

function fxStaffSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FX_STAFF);
  if (!sh) {
    sh = ss.insertSheet(FX_STAFF);
    sh.getRange(1, 1, 1, 5).setValues([['Mã NV', 'Họ tên', 'Mật khẩu (nhập rồi chạy maHoaPinNhanVien)', 'Mã hóa (tự điền)', 'Trạng thái (trống = hoạt động, KHÓA = khóa)']]).setFontWeight('bold');
    sh.getRange(2, 1, 999, 5).setNumberFormat('@');   // giữ nguyên số 0 đầu của mật khẩu
    sh.setFrozenRows(1);
  }
  return sh;
}

function fxOptSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FX_OPT);
  if (!sh) {
    sh = ss.insertSheet(FX_OPT);
    sh.getRange(1, 1, 1, 3).setValues([['Cách sửa (mỗi dòng là 1 nút)', 'Thêm bởi', 'Thời gian']]).setFontWeight('bold');
    var seed = ['Sửa lại mã khuôn', 'Sửa lại số xe', 'Sửa Lot BTP', 'Quét lại', 'Sửa label', 'Xóa dòng quét nhầm', 'Báo kỹ thuật xử lý', 'Khác'];
    var rows = seed.map(function (s) { return [s, 'HỆ THỐNG', fxNow_()]; });
    sh.getRange(2, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
    sh.setColumnWidth(1, 300);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ---------- Mật khẩu: lưu dạng băm SHA-256 + muối, không lưu chữ thô ----------
function fxSalt_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('fxSalt');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('fxSalt', s); }
  return s;
}

function fxHash_(id, pin) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    fxSalt_() + '|' + String(id).toUpperCase() + '|' + pin, Utilities.Charset.UTF_8);
  return raw.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length < 2 ? '0' + v : v; }).join('');
}

// Quản trị: nhập Mã NV (cột A), Họ tên (cột B), Mật khẩu (cột C) vào sheet NV_DANH_SACH rồi chạy hàm này.
// Mật khẩu thô sẽ bị XÓA ngay, chỉ còn mã hóa ở cột D.
function maHoaPinNhanVien() {
  var sh = fxStaffSheet_(), last = sh.getLastRow(), n = 0;
  if (last < 2) { Logger.log('Chưa có nhân viên nào.'); return; }
  var v = sh.getRange(2, 1, last - 1, 5).getDisplayValues();
  for (var i = 0; i < v.length; i++) {
    var id = String(v[i][0]).trim(), pin = String(v[i][2]).trim();
    if (id && pin) {
      sh.getRange(2 + i, 4).setValue(fxHash_(id, pin));
      sh.getRange(2 + i, 3).setValue('');
      n++;
    }
  }
  Logger.log('Đã mã hóa mật khẩu cho ' + n + ' nhân viên.');
}

// Hoặc thêm / đổi mật khẩu trực tiếp bằng code: themNhanVien('NV001', 'Nguyễn Văn A', '1234')
function themNhanVien(id, ten, pin) {
  id = String(id || '').trim().toUpperCase(); pin = String(pin || '').trim();
  if (!id || !pin) { Logger.log('Thiếu mã hoặc mật khẩu'); return; }
  var sh = fxStaffSheet_(), last = sh.getLastRow(), row = 0;
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim().toUpperCase() === id) { row = 2 + i; break; }
  }
  if (!row) row = Math.max(last, 1) + 1;
  sh.getRange(row, 1, 1, 5).setNumberFormat('@').setValues([[id, ten || '', '', fxHash_(id, pin), '']]);
  Logger.log('Đã lưu nhân viên ' + id);
}

function fxSess_(tok) {
  if (!tok) return null;
  var c = CacheService.getScriptCache(), s = c.get('fxt_' + tok);
  if (!s) return null;
  c.put('fxt_' + tok, s, FX_SESSION_SEC);
  return JSON.parse(s);
}

function fxLogin(id, pin) {
  id = String(id || '').trim().toUpperCase(); pin = String(pin || '').trim();
  if (!id || !pin) return { ok: false, msg: 'Nhập mã nhân viên và mật khẩu' };
  var cache = CacheService.getScriptCache(), fk = 'fxf_' + id, fails = Number(cache.get(fk) || 0);
  if (fails >= FX_MAX_FAIL) return { ok: false, msg: 'Sai quá nhiều lần. Thử lại sau 5 phút' };
  var sh = fxStaffSheet_(), last = sh.getLastRow(), found = null;
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, 5).getDisplayValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim().toUpperCase() === id) { found = v[i]; break; }
    }
  }
  var ok = found && String(found[4]).trim().toUpperCase() !== 'KHÓA' && found[3] && String(found[3]) === fxHash_(id, pin);
  if (!ok) { cache.put(fk, String(fails + 1), 300); return { ok: false, msg: 'Sai mã nhân viên hoặc mật khẩu' }; }
  cache.remove(fk);
  var tok = Utilities.getUuid(), name = String(found[1]).trim();
  cache.put('fxt_' + tok, JSON.stringify({ id: id, name: name }), FX_SESSION_SEC);
  return { ok: true, token: tok, id: id, name: name };
}

function fxLogout(tok) {
  if (tok) CacheService.getScriptCache().remove('fxt_' + tok);
  return { ok: true };
}

// ---------- Danh sách nút "cách sửa" (thêm tự do) ----------
function fxFixOptions() {
  var sh = fxOptSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getDisplayValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return s; });
}

function fxAddOption(tok, text) {
  var u = fxSess_(tok);
  if (!u) return { ok: false, err: 'SESSION', msg: 'Hết phiên đăng nhập, vui lòng nhập lại' };
  text = String(text || '').trim().replace(/\s+/g, ' ');
  if (!text || text.length > 40) return { ok: false, msg: 'Nội dung nút từ 1 đến 40 ký tự', options: fxFixOptions() };
  var list = fxFixOptions();
  for (var i = 0; i < list.length; i++) {
    if (list[i].toLowerCase() === text.toLowerCase()) return { ok: true, options: list, added: list[i] };
  }
  var sh = fxOptSheet_();
  var r = Math.max(sh.getLastRow() + 1, 2);
  sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[text, u.id + ' - ' + u.name, fxNow_()]]);
  return { ok: true, options: fxFixOptions(), added: text };
}

// Địa chỉ file Google Sheet + mã (gid) của từng tab, để Web App mở đúng tab / đúng ô cần sửa
function fxSheetLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), g = {};
  ss.getSheets().forEach(function (sh) { g[sh.getName()] = sh.getSheetId(); });
  return { base: String(ss.getUrl()).split('#')[0], gids: g };
}

// ---------- Đọc phiếu ----------
function fxRead_(sh, n) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, Math.min(n, last - 1), FX_COLS.length).getDisplayValues();
}

function fxObj_(r, maxErr) {
  var e = String(r[6] || '');
  if (maxErr && e.length > maxErr) e = e.substring(0, maxErr) + '…';
  return { id: r[0], tLoi: r[1], kind: r[2], line: r[3], xe: r[4], scan: r[5], err: e, st: r[7],
    tSua: r[8], uid: r[9], uname: r[10], fix: r[11], note: r[12], recheck: r[13], mail: r[14],
    snap: String(r[15] || '').substring(0, 700), cmp: String(r[16] || '').substring(0, 700) };
}

// Web App gọi mỗi vài giây: 40 phiếu mới nhất + mọi phiếu còn CHỜ SỬA (kết quả được lưu đệm 10 giây)
function fxTickets() {
  var cache = CacheService.getScriptCache(), c = cache.get('fxTk');
  if (c) { try { return JSON.parse(c); } catch (x) {} }
  var rows = fxRead_(fxSheet_(), 300), list = [], open = 0;
  for (var i = 0; i < rows.length; i++) {
    var isOpen = rows[i][7] === FX_ST_OPEN;
    if (isOpen) open++;
    if (i < 40 || isOpen) list.push(fxObj_(rows[i], 500));
  }
  var out = { list: list, open: open };
  try { var js = JSON.stringify(out); if (js.length < 90000) cache.put('fxTk', js, 10); } catch (e) {}
  return out;
}

// ---------- Tự tạo phiếu khi NG ----------
function fxStrip_(s) {
  return String(s || '').replace(/<[^>]*>/g, '').replace(/&quot;/g, '"').replace(/&#039;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').trim();
}

// Trả về { bad: true/false, text: mô tả lỗi } — chỉ ĐỌC, dùng đúng các hàm kiểm tra sẵn có
function fxErrInfo_(base, line) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(line), lines = [];
  if (base === 'KHUON') {
    var fn = getSheetCheckMap()[line];
    if (!fn) return { bad: true, text: 'Không có quy tắc kiểm tra cho line ' + line };
    if (!sheet) return { bad: true, text: 'Không tìm thấy sheet ' + line, rows: [] };
    if (sheet.getLastRow() < 7) return { bad: true, text: 'Không có dữ liệu từ hàng 7', rows: [] };
    var res = fn(sheet);
    xeMerge_(res, checkXeSequence_(sheet));
    if (res.valid) return { bad: false, text: '', rows: [] };
    var d = res.details || [];
    var badRows = [];
    for (var q = 0; q < d.length; q++) { var rw = d[q] && typeof d[q] === 'object' ? parseInt(d[q].row, 10) : NaN; if (rw >= 7 && badRows.indexOf(rw) < 0) badRows.push(rw); }
    for (var i = 0; i < d.length && i < 10; i++) {
      var e = d[i];
      if (e && typeof e === 'object') {
        var pos = e.row ? ('Hàng ' + e.row + (e.col ? ', Cột ' + xeColLetter_(parseInt(e.col, 10)) : '')) : '';
        lines.push((pos ? pos + ': ' : '') + (e.message || 'Lỗi') +
          (e.current !== undefined ? ' | Hiện tại: ' + e.current : '') +
          (e.expected !== undefined ? ' | Mong đợi: ' + e.expected : ''));
      } else lines.push(String(e));
    }
    if (d.length > 10) lines.push('… và ' + (d.length - 10) + ' lỗi khác');
    if (!lines.length) lines.push(res.message || 'Có lỗi (không có chi tiết)');
    return { bad: true, text: lines.join('\n'), rows: badRows };
  }
  var rules = getLabelRules(), rule = null;
  for (var k = 0; k < rules.length; k++) if (rules[k].sheet === line) { rule = rules[k]; break; }
  if (!rule) return { bad: true, text: 'Không có quy tắc label cho ' + line };
  if (!sheet) return { bad: true, text: 'Không tìm thấy sheet ' + line, rows: [] };
  if (sheet.getLastRow() < 7) return { bad: true, text: 'Không có dữ liệu từ hàng 7', rows: [] };
  if (checkLabelSheetByRule(sheet, rule).valid) return { bad: false, text: '', rows: [] };
  var items = getLabelErrorDetails_(sheet, rule), lrows = [];
  for (var j = 0; j < items.length; j++) {
    if (j < 10) lines.push(fxStrip_(items[j]));
    var mm = /Dòng (\d+)/.exec(String(items[j])); if (mm && lrows.indexOf(+mm[1]) < 0) lrows.push(+mm[1]);
  }
  if (items.length > 10) lines.push('… và ' + (items.length - 10) + ' lỗi khác');
  return { bad: true, text: lines.join('\n'), rows: lrows };
}

// Bản sao nội dung các dòng đang lỗi (chỉ ĐỌC sheet) — để còn bằng chứng khi nhân viên xóa dòng lỗi rồi quét lại
function fxRowText_(row) {
  var a = row.map(function (v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); });
  while (a.length && a[a.length - 1] === '') a.pop();
  return a.join(' | ');
}

function fxSnapRows_(line, rows) {
  try {
    if (!rows || !rows.length) return '';
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(line);
    if (!sh) return '';
    var lc = sh.getLastColumn(), lr = sh.getLastRow(), out = [];
    rows.sort(function (a, b) { return a - b; });
    for (var i = 0; i < rows.length && i < 8; i++) {
      if (rows[i] > lr || lc < 1) continue;
      var t = fxRowText_(sh.getRange(rows[i], 1, 1, lc).getDisplayValues()[0]);
      out.push('Hàng ' + rows[i] + ': ' + (t.length > 300 ? t.substring(0, 300) + '…' : t));
    }
    if (rows.length > 8) out.push('… và ' + (rows.length - 8) + ' dòng lỗi khác');
    return out.join('\n');
  } catch (e) { return ''; }
}

// Đối chiếu: với từng dòng lỗi lúc NG, in nội dung HIỆN TẠI của đúng hàng đó theo cùng định dạng "Hàng N: a | b | c",
// kèm trạng thái (còn nguyên văn / đã thay đổi). Chỉ ĐỌC sheet.
function fxCompare_(line, snap) {
  try {
    snap = String(snap || '');
    if (!snap) return 'Không có bản sao dòng lỗi để đối chiếu';
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(line);
    if (!sh) return 'Không tìm thấy sheet ' + line;
    var lr = sh.getLastRow(), lc = sh.getLastColumn(), set = {}, cur = [];
    if (lr >= 7 && lc >= 1) {
      cur = sh.getRange(7, 1, lr - 6, lc).getDisplayValues();
      for (var i = 0; i < cur.length; i++) set[fxRowText_(cur[i])] = 7 + i;
    }
    var out = [], lines = snap.split('\n');
    for (var k = 0; k < lines.length; k++) {
      var m = /^Hàng (\d+): ([\s\S]*)$/.exec(lines[k]);
      if (!m) { out.push(lines[k]); continue; }
      var rn = parseInt(m[1], 10), txt = m[2].replace(/…$/, '');
      var nowTxt = (rn >= 7 && rn - 7 < cur.length) ? fxRowText_(cur[rn - 7]) : '';
      var shown = nowTxt ? (nowTxt.length > 300 ? nowTxt.substring(0, 300) + '…' : nowTxt) : '(trống / không còn dữ liệu)';
      var hit = set.hasOwnProperty(m[2]) ? set[m[2]] : 0;
      if (!hit && m[2] !== txt) { for (var key in set) if (key.indexOf(txt) === 0) { hit = set[key]; break; } }
      var st = hit ? ('VẪN NGUYÊN NHƯ LÚC NG (còn ở hàng ' + hit + ')') : 'ĐÃ THAY ĐỔI (dòng lỗi cũ không còn nguyên văn – đã xóa hoặc sửa)';
      out.push('Hàng ' + rn + ': ' + shown + '   → ' + st);
    }
    return out.join('\n');
  } catch (e) { return 'Không đối chiếu được: ' + e.message; }
}

// Kiểm tra lại line ngay lúc khai báo bằng ĐÚNG các hàm kiểm tra sẵn có (chỉ để báo kết quả, không ghi lịch sử / còi / đèn)
function fxRecheck_(kind, line) {
  try {
    var info = fxErrInfo_(kind, line);
    return { ok: !info.bad, text: info.bad ? info.text : '' };
  } catch (e) {
    return { ok: null, text: 'Không kiểm tra lại được: ' + e.message };
  }
}

function fxNewTicket_(kind, line, xe, scan, err, snap) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = fxSheet_(), rows = fxRead_(sh, 300), i;
    for (i = 0; i < rows.length; i++) {
      if (rows[i][2] === kind && rows[i][3] === line && rows[i][7] === FX_ST_OPEN) return;   // đã có phiếu chờ sửa
    }
    for (i = 0; i < rows.length; i++) {
      if (rows[i][2] === kind && rows[i][3] === line && rows[i][7] === FX_ST_FIXED) sh.getRange(2 + i, 8).setValue(FX_ST_BAD);   // sửa rồi mà vẫn NG
    }
    var id = 'P' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);
    sh.insertRowAfter(1);
    var rg = sh.getRange(2, 1, 1, FX_COLS.length);
    rg.setNumberFormat('@');
    rg.setValues([[id, fxNow_(), kind, line, xe || '', scan || '', err || '', FX_ST_OPEN, '', '', '', '', '', '', '', snap || '', '']]);
    CacheService.getScriptCache().remove('fxTk');
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Line OK trở lại: đánh dấu giờ OK lại; nếu đã khai báo sửa thì xác nhận luôn
function fxOnOk_(kind, line) {
  var sh = fxSheet_();
  if (sh.getLastRow() < 2) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var rows = fxRead_(sh, 300), now = fxNow_(), changed = false;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[2] !== kind || (line && r[3] !== line) || r[13]) continue;
      if (r[7] === FX_ST_OPEN) {
        sh.getRange(2 + i, 14).setNumberFormat('@').setValue(now); changed = true;
      } else if (r[7] === FX_ST_FIXED) {
        sh.getRange(2 + i, 14).setNumberFormat('@').setValue(now);
        sh.getRange(2 + i, 8).setValue(FX_ST_OK); changed = true;
      }
    }
    if (changed) CacheService.getScriptCache().remove('fxTk');
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function fxOnHistory_(kind, sheetName, status, text, xe, ten) {
  var base = String(kind).indexOf('LABEL') === 0 ? 'LABEL' : (String(kind).indexOf('KHUON') === 0 ? 'KHUON' : '');
  if (!base) return;
  var all = /-ALL$/.test(String(kind));
  if (status === 'OK') { fxOnOk_(base, all ? '' : sheetName); return; }
  if (status !== 'NG') return;
  if (!all) {
    var info = fxErrInfo_(base, sheetName);
    fxNewTicket_(base, sheetName, xe, ten, info.bad ? info.text : (text || 'Không xác định được chi tiết lỗi (có thể đã được sửa)'), info.bad ? fxSnapRows_(sheetName, info.rows) : '');
    return;
  }
  var names = base === 'KHUON' ? Object.keys(getSheetCheckMap()) : getLabelRules().map(function (r) { return r.sheet; });
  for (var i = 0; i < names.length; i++) {
    var inf = fxErrInfo_(base, names[i]);
    if (inf.bad) fxNewTicket_(base, names[i], '', '', inf.text, fxSnapRows_(names[i], inf.rows));
  }
}

// Khai báo lại writeHistory (khai báo sau cùng nên thay thế bản cũ): giữ NGUYÊN việc ghi IoT_History, chỉ thêm bước tạo / cập nhật phiếu.
function writeHistory(kind, sheetName, status, text, xe, ten) {
  var sheet = ensureHistorySheet_();
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 7).setValues([[
    Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy'),
    kind || '',
    sheetName || '',
    status || '',
    text || '',
    xe || '',
    ten || ''
  ]]);
  try { fxOnHistory_(kind, sheetName, status, text, xe, ten); } catch (err) {
    try { Logger.log('Lỗi tạo phiếu: ' + err.message); } catch (x) {}
  }
}

// ---------- Nhân viên khai báo đã chỉnh sửa ----------
function fxParse_(s) {
  var m = /(\d+):(\d+):(\d+)\s+(\d+)\/(\d+)\/(\d+)/.exec(String(s || ''));
  return m ? Date.UTC(+m[6], +m[5] - 1, +m[4], +m[1] - 7, +m[2], +m[3]) : NaN;
}

function fxMail_(t, u, opts, note, tSua, cmp, rc) {
  try {
    var ms = fxParse_(t[1]), mins = isNaN(ms) ? '' : Math.max(0, Math.round((Date.now() - ms) / 60000)) + ' phút';
    var kindName = t[2] === 'LABEL' ? 'Label' : 'Mã khuôn';
    var row = function (k, v) {
      return '<tr><td style="padding:8px;background:#f1f5f9;font-weight:bold;width:150px;">' + k + '</td><td style="padding:8px;">' + v + '</td></tr>';
    };
    var html = '<h2 style="color:#166534;">✅ Đã khai báo chỉnh sửa – ' + escHtml_(kindName) + ' line ' + escHtml_(t[3]) + '</h2>' +
      '<table border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:720px;">' +
      row('Mã phiếu', escHtml_(t[0])) +
      row('Thời gian lỗi', escHtml_(t[1])) +
      row('Số xe / Người scan', escHtml_(t[4] || '-') + ' / ' + escHtml_(t[5] || '-')) +
      row('Nội dung lỗi', '<div style="white-space:pre-wrap;background:#fee2e2;padding:8px;border-radius:4px;">' + escHtml_(t[6]) + '</div>') +
      (t[15] ? row('Dòng lỗi lúc NG', '<div style="white-space:pre-wrap;font-family:monospace;background:#f8fafc;padding:8px;border-radius:4px;">' + escHtml_(t[15]) + '</div>') : '') +
      (cmp ? row('Đối chiếu lại (cùng định dạng)', '<div style="white-space:pre-wrap;font-family:monospace;background:#fef3c7;padding:8px;border-radius:4px;">' + escHtml_(cmp) + '</div>') : '') +
      row('Kết quả kiểm tra lại', rc && rc.ok === true
        ? '<span style="display:inline-block;background:#16a34a;color:#fff;font-weight:bold;padding:6px 14px;border-radius:6px;">✅ OK – khớp, hợp lệ</span>'
        : (rc && rc.ok === false
          ? '<span style="display:inline-block;background:#dc2626;color:#fff;font-weight:bold;padding:6px 14px;border-radius:6px;">❌ NG – vẫn còn lỗi</span><div style="white-space:pre-wrap;background:#fee2e2;padding:8px;border-radius:4px;margin-top:6px;">' + escHtml_(rc.text) + '</div>'
          : '<span style="color:#92400e;font-weight:bold;">⚠️ Chưa xác định</span> ' + escHtml_((rc && rc.text) || ''))) +
      row('Thời gian sửa', escHtml_(tSua) + (mins ? ' (sau ' + mins + ')' : '')) +
      row('Người sửa', escHtml_(u.id) + ' – ' + escHtml_(u.name)) +
      row('Cách sửa', escHtml_(opts.join(', '))) +
      row('Ghi chú', escHtml_(note || '-')) +
      '</table><p>Hệ thống sẽ tự kiểm tra lại ở lần quét kế tiếp. Nếu vẫn NG, phiếu sẽ chuyển thành "SỬA CHƯA ĐẠT".</p>' +
      '<p>Trân trọng,<br><em>Script kiểm tra tự động</em></p>';
    MailApp.sendEmail({
      to: ALERT_EMAILS,
      subject: 'ĐÃ CHỈNH SỬA' + (rc && rc.ok === true ? ' [KIỂM TRA LẠI: OK]' : (rc && rc.ok === false ? ' [VẪN NG]' : '')) + ': ' + kindName + ' line ' + t[3] + ' – ' + u.name + ' – ' + tSua,
      htmlBody: html
    });
    return true;
  } catch (e) {
    Logger.log('Lỗi gửi email chỉnh sửa: ' + e.message);
    return false;
  }
}

function fxSubmitFix(tok, ticketId, opts, note) {
  var u = fxSess_(tok);
  if (!u) return { ok: false, err: 'SESSION', msg: 'Hết phiên đăng nhập, vui lòng nhập lại' };
  var valid = fxFixOptions(), chosen = [];
  (opts || []).forEach(function (o) { o = String(o).trim(); if (o && valid.indexOf(o) >= 0 && chosen.indexOf(o) < 0) chosen.push(o); });
  if (!chosen.length) return { ok: false, msg: 'Chọn ít nhất 1 cách sửa' };
  note = String(note || '').trim().substring(0, 300);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (x) { return { ok: false, msg: 'Hệ thống đang bận, thử lại sau vài giây' }; }
  try {
    var sh = fxSheet_(), last = sh.getLastRow(), row = 0;
    if (last >= 2) {
      var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = 0; i < ids.length; i++) if (ids[i][0] === ticketId) { row = 2 + i; break; }
    }
    if (!row) return { ok: false, msg: 'Không tìm thấy phiếu' };
    var t = sh.getRange(row, 1, 1, FX_COLS.length).getDisplayValues()[0];
    if (t[7] !== FX_ST_OPEN) return { ok: false, msg: 'Phiếu này đã được khai báo trước đó' };
    var now = fxNow_();
    sh.getRange(row, 8, 1, 6).setNumberFormat('@').setValues([[t[13] ? FX_ST_OK : FX_ST_FIXED, now, u.id, u.name, chosen.join(', '), note]]);
    var rc = fxRecheck_(t[2], t[3]);
    var cmp = fxCompare_(t[3], t[15]);
    var verdict = rc.ok === true ? 'KẾT QUẢ KIỂM TRA LẠI: OK – khớp, hợp lệ' : (rc.ok === false ? 'KẾT QUẢ KIỂM TRA LẠI: NG – vẫn còn lỗi' : 'KẾT QUẢ KIỂM TRA LẠI: chưa xác định');
    sh.getRange(row, 17).setNumberFormat('@').setValue(verdict + '\n' + cmp);
    var sent = fxMail_(t, u, chosen, note, now, cmp, rc);
    sh.getRange(row, 15).setValue(sent ? 'Đã gửi' : 'Chưa gửi được');
    CacheService.getScriptCache().remove('fxTk');
    return { ok: true, emailSent: sent };
  } finally {
    try { lock.releaseLock(); } catch (x2) {}
  }
}

// ---------- Cài đặt (chạy tay 1 lần) ----------
function caiDatSuaLoi() {
  fxSheet_(); fxOptSheet_();
  var st = fxStaffSheet_();
  try {   // chỉ chủ file xem / sửa được danh sách nhân viên (mật khẩu đã mã hóa)
    var p = st.protect().setDescription('Danh sách nhân viên khai báo chỉnh sửa');
    var me = Session.getEffectiveUser();
    p.addEditor(me);
    p.removeEditors(p.getEditors().filter(function (e) { return e.getEmail() !== me.getEmail(); }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
  } catch (e) { Logger.log('Chưa khóa được sheet ' + FX_STAFF + ': ' + e.message); }
  Logger.log('Xong. Bước tiếp: điền sheet ' + FX_STAFF + ' (Mã NV, Họ tên, Mật khẩu) rồi chạy maHoaPinNhanVien. Sau đó Deploy lại Web App.');
}

// ---------- Web App: chèn bảng "Lỗi & chỉnh sửa" vào tab GIÁM SÁT ----------
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ui';
  var out = doGetUi_(e);
  if (action === 'latest' || action === 'ping') return out;
  try {
    var html = out.getContent(), i = html.lastIndexOf('</body>');
    if (i < 0) return out;
    html = html.substring(0, i) + fxSnippet_() + html.substring(i);
    return HtmlService.createHtmlOutput(html).setTitle('SCAN SYSTEM 2.1')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return out;   // lỗi chèn giao diện thì vẫn hiện giao diện cũ
  }
}

// Lưu ý: nằm trong template literal nên KHÔNG dùng dấu huyền, ${ hay dấu gạch chéo ngược bên dưới.
function fxSnippet_() {
  return `
<style>
.fxc{margin-top:14px}
.fxf{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
.fxt{border:1px solid var(--line);background:var(--card2);color:var(--text);border-radius:14px;padding:12px;margin-bottom:8px}
.fxt.o{border-color:var(--ng);background:var(--ngbg)}
.fxt.k{border-left:4px solid var(--ok)}
.fxt.w{border-left:4px solid var(--warn)}
.fxh{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-bottom:6px}
.fxh b{font-size:calc(16px*var(--s))}
.fxbd{padding:2px 10px;border-radius:999px;font-size:11px;font-weight:900;border:1px solid var(--line);white-space:nowrap}
.fxbd.o{background:var(--ng);color:#fff;border-color:var(--ng)}
.fxbd.w{background:var(--warn);color:#1f2937;border-color:var(--warn)}
.fxbd.k{background:var(--ok);color:#fff;border-color:var(--ok)}
.fxe{white-space:pre-wrap;word-break:break-word;font-size:calc(12.5px*var(--s));line-height:1.5;margin:6px 0;padding:8px 10px;border-radius:8px;background:var(--card);border:1px solid var(--line);color:var(--text)}
.fxx{font-size:calc(12.5px*var(--s));line-height:1.65;margin-top:6px}
.fxx b{color:var(--text)}
.fxm0{font-size:12px;color:var(--mute)}
.fxgo{width:100%;margin-top:10px;padding:14px;border-radius:14px;border:0;font-size:15px;font-weight:900;color:var(--acctx);background:var(--acc);min-height:48px}
.fxcl{width:100%;margin-top:8px;padding:12px;border-radius:14px;border:1px solid var(--line);font-size:14px;font-weight:800;color:var(--text);background:var(--card2)}
.fxov{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(2,6,15,.82);display:none;align-items:center;justify-content:center;z-index:1001;padding:14px}
.fxov.show{display:flex}
.fxcard{width:100%;max-width:480px;max-height:92vh;overflow:auto;background:var(--card);border:2px solid var(--line);border-radius:22px;padding:20px;color:var(--text)}
.fxtt{font-size:calc(20px*var(--s));font-weight:900;margin-bottom:10px}
.fxs{padding:10px;border-radius:12px;background:var(--card2);border:1px solid var(--line);margin-bottom:10px;font-size:13px}
.fxl{display:block;margin:10px 0 4px;font-size:12px;font-weight:800;color:var(--mute)}
.fxcard input,.fxcard textarea{width:100%;background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:12px;color:var(--text);font-size:16px;outline:none;font-family:inherit}
.fxcard textarea{min-height:64px;resize:vertical}
.fxcard input:focus,.fxcard textarea:focus{border-color:var(--text)}
.fxmsg{margin-top:10px;padding:10px;border-radius:10px;background:var(--ngbg);color:var(--ngtx);border:1px solid var(--ng);font-size:13px;font-weight:700}
.fxok{margin-top:10px;padding:12px;border-radius:10px;background:var(--okbg);color:var(--oktx);border:1px solid var(--ok);font-size:14px;font-weight:700}
.fxopts{display:flex;flex-wrap:wrap;gap:8px}
.fxopts .fb{padding:10px 14px;font-size:14px}
.fxadd{display:flex;gap:8px;margin-top:10px}
.fxadd button{flex:none;padding:0 16px;border-radius:12px;border:1px solid var(--line);background:var(--card2);color:var(--text);font-weight:900;font-size:14px}
.fxwho{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:13px;font-weight:800;margin-bottom:8px}
.fxwho button{border:1px solid var(--line);background:var(--card2);color:var(--mute);border-radius:999px;padding:5px 12px;font-size:12px;font-weight:800}
</style>
<div class="fxov" id="fxM"><div class="fxcard" id="fxCard"></div></div>
<script>
var fxLinks=null, fxPopup=null, fxTok=null, fxList=[], fxOpen=0, fxFil='open', fxCur=null, fxOpts=[], fxSel={}, fxMsg='', fxNote='', fxNew='', fxLast='', fxBusy=false;
function fxEsc(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function fxLoc(t){
  var e=String(t.err||''), m=/(?:Hàng|Dòng) ([0-9]+)/.exec(e), row=m?Number(m[1]):7, col='A';
  if(m){ var c=/Cột ([A-Z]+)/.exec(e.slice(m.index)); if(c) col=c[1]; }
  return col+row;
}
function fxUrl(t){
  if(!fxLinks||!fxLinks.gids||!(t.line in fxLinks.gids)) return '';
  return fxLinks.base+'#gid='+fxLinks.gids[t.line]+'&range='+fxLoc(t);
}
function fxKind(k){ return k==='LABEL'?'Label':'Mã khuôn'; }
function fxBadge(t){
  if(t.st==='CHỜ SỬA') return t.recheck?['w','ĐÃ OK LẠI · CHƯA KHAI BÁO']:['o','CHỜ SỬA'];
  if(t.st==='ĐÃ XÁC NHẬN OK') return ['k','ĐÃ SỬA · OK'];
  if(t.st==='SỬA CHƯA ĐẠT') return ['o','SỬA CHƯA ĐẠT'];
  return ['w','ĐÃ SỬA · CHỜ KIỂM TRA LẠI'];
}
(function(){
  var mon=document.getElementById('mon'); if(!mon) return;
  var c=document.createElement('div'); c.className='fxc';
  c.innerHTML='<div class="dsep">Lỗi &amp; chỉnh sửa</div><div class="fxf" id="fxF"></div><div id="fxL"></div>';
  mon.appendChild(c);
  var F=document.getElementById('fxF');
  [['open','Chờ sửa'],['all','Tất cả']].forEach(function(p){
    var b=document.createElement('button'); b.className='fb'+(p[0]===fxFil?' on':''); b.innerText=p[1]; b.setAttribute('data-f',p[0]);
    b.onclick=function(){ fxFil=p[0]; Array.prototype.forEach.call(F.children,function(x){ x.classList.toggle('on',x.getAttribute('data-f')===fxFil); }); fxLast=''; fxRender(); };
    F.appendChild(b);
  });
  var L=document.getElementById('fxL');
  L.onclick=function(ev){
    var t=ev.target;
    while(t&&t!==L){ if(t.getAttribute&&t.getAttribute('data-id')){ fxOpenFix(t.getAttribute('data-id')); return; } t=t.parentNode; }
  };
})();
function fxRender(){
  var L=document.getElementById('fxL'); if(!L) return;
  var rows=fxList.filter(function(t){ return fxFil==='all'||t.st==='CHỜ SỬA'; });
  var html=rows.map(function(t){
    var b=fxBadge(t), cls=(b[0]==='o'&&t.st==='CHỜ SỬA')?'o':(b[0]==='k'?'k':(b[0]==='w'?'w':'o'));
    var h='<div class="fxt '+cls+'"><div class="fxh"><span class="fxbd '+b[0]+'">'+b[1]+'</span><b>'+fxEsc(t.line)+'</b><span class="fxm0">'+fxKind(t.kind)+' · 🕒 '+fxEsc(t.tLoi)+'</span>';
    if(t.xe) h+='<span class="chip xe"><i>🚚</i>'+fxEsc(t.xe)+'</span>';
    if(t.scan) h+='<span class="chip ten"><i>👤</i>'+fxEsc(t.scan)+'</span>';
    h+='</div><div class="fxm0">Nội dung lỗi</div><div class="fxe">'+fxEsc(t.err)+'</div>';
    if(t.snap) h+='<div class="fxm0">Dòng lỗi lúc NG (bản sao)</div><div class="fxe">'+fxEsc(t.snap)+'</div>';
    if(t.cmp) h+='<div class="fxm0">Đối chiếu khi khai báo</div><div class="fxe">'+fxEsc(t.cmp)+'</div>';
    if(t.uid){
      h+='<div class="fxx">🛠 <b>Đã chỉnh sửa</b> lúc '+fxEsc(t.tSua)+'<br>👤 <b>Người sửa:</b> '+fxEsc(t.uid)+' – '+fxEsc(t.uname)+'<br><b>Cách sửa:</b> '+fxEsc(t.fix);
      if(t.note) h+='<br><b>Ghi chú:</b> '+fxEsc(t.note);
      h+='</div>';
    }
    if(t.recheck) h+='<div class="fxx">✅ OK lại lúc '+fxEsc(t.recheck)+'</div>';
    var lk=fxUrl(t);
    if(lk&&(t.st==='CHỜ SỬA'||t.st==='SỬA CHƯA ĐẠT')) h+='<a class="fxcl" style="display:block;text-align:center;text-decoration:none" target="_blank" rel="noopener" href="'+fxEsc(lk)+'">📄 Mở Google Sheet tại ô '+fxEsc(fxLoc(t))+'</a>';
    if(t.st==='CHỜ SỬA') h+='<button class="fxgo" data-id="'+fxEsc(t.id)+'">Khai báo chỉnh sửa</button>';
    return h+'</div>';
  }).join('');
  if(!html) html='<div class="empty">'+(fxFil==='all'?'Chưa có phiếu lỗi nào':'Không có lỗi nào đang chờ sửa 🎉')+'</div>';
  if(html!==fxLast){ L.innerHTML=html; fxLast=html; }
}
function fxPoll(){
  google.script.run.withSuccessHandler(function(d){
    if(!d) return;
    fxList=d.list||[]; fxOpen=d.open||0;
    fxRender();
    var t=document.getElementById('tabMon'); if(t) t.innerText='GIÁM SÁT'+(fxOpen>0?' ('+fxOpen+')':'');
  }).fxTickets();
}
/* ===== Hộp thoại khai báo ===== */
function fxShow(h){
  document.getElementById('fxCard').innerHTML=h;
  document.getElementById('fxM').classList.add('show');
}
function fxClose(){ document.getElementById('fxM').classList.remove('show'); fxMsg=''; fxBusy=false; }
function fxSummary(){
  var t=fxCur; if(!t) return '';
  return '<div class="fxs"><b>'+fxEsc(t.line)+'</b> · '+fxKind(t.kind)+' · 🕒 '+fxEsc(t.tLoi)+'<div class="fxe">'+fxEsc(t.err)+'</div></div>';
}
function fxKeep(){
  var n=document.getElementById('fxNote'); if(n) fxNote=n.value;
  var w=document.getElementById('fxNew'); if(w) fxNew=w.value;
}
function fxStep1(){
  fxShow('<div class="fxtt">Khai báo chỉnh sửa</div>'+fxSummary()+
    '<label class="fxl">Mã nhân viên</label><input id="fxId" type="text" autocomplete="off" autocapitalize="characters">'+
    '<label class="fxl">Mật khẩu</label><input id="fxPin" type="password" autocomplete="off">'+
    (fxMsg?'<div class="fxmsg">'+fxEsc(fxMsg)+'</div>':'')+
    '<button class="fxgo" data-act="login">XÁC NHẬN</button><button class="fxcl" data-act="close">Hủy</button>');
  var i=document.getElementById('fxId'); if(i) setTimeout(function(){ try{i.focus();}catch(e){} },80);
}
function fxStep2(){
  var me=fxWho||{};
  fxShow('<div class="fxwho"><span>👤 '+fxEsc(me.id)+' – '+fxEsc(me.name)+'</span><button data-act="logout">Đổi người</button></div>'+fxSummary()+
    '<label class="fxl">Cách sửa (chọn một hoặc nhiều)</label><div class="fxopts">'+
    fxOpts.map(function(o,i){ return '<button class="fb'+(fxSel[o]?' on':'')+'" data-act="opt" data-i="'+i+'">'+fxEsc(o)+'</button>'; }).join('')+
    '</div><div class="fxadd"><input id="fxNew" type="text" maxlength="40" placeholder="Thêm nút mới..." value="'+fxEsc(fxNew)+'"><button data-act="add">+ Thêm</button></div>'+
    '<label class="fxl">Ghi chú (không bắt buộc)</label><textarea id="fxNote" maxlength="300">'+fxEsc(fxNote)+'</textarea>'+
    (fxMsg?'<div class="fxmsg">'+fxEsc(fxMsg)+'</div>':'')+
    '<button class="fxgo" data-act="submit">GỬI KHAI BÁO</button><button class="fxcl" data-act="close">Hủy</button>');
}
var fxWho=null;
function fxOpenFix(id){
  fxCur=null;
  for(var i=0;i<fxList.length;i++) if(fxList[i].id===id){ fxCur=fxList[i]; break; }
  if(!fxCur) return;
  fxMsg=''; fxSel={}; fxNote=''; fxNew='';
  if(fxTok){ fxLoadOpts(); } else fxStep1();
}
function fxLoadOpts(){
  fxShow('<div class="fxtt">Đang tải...</div>');
  google.script.run.withSuccessHandler(function(l){ fxOpts=l||[]; fxStep2(); })
    .withFailureHandler(function(e){ fxMsg='Lỗi: '+e.message; fxStep2(); }).fxFixOptions();
}
function fxDoLogin(){
  if(fxBusy) return;
  var id=document.getElementById('fxId').value, pin=document.getElementById('fxPin').value;
  fxBusy=true;
  google.script.run.withSuccessHandler(function(r){
    fxBusy=false;
    if(r&&r.ok){ fxTok=r.token; fxWho={id:r.id,name:r.name}; fxMsg=''; fxLoadOpts(); }
    else { fxMsg=(r&&r.msg)||'Đăng nhập thất bại'; fxStep1(); }
  }).withFailureHandler(function(e){ fxBusy=false; fxMsg='Lỗi: '+e.message; fxStep1(); }).fxLogin(id,pin);
}
function fxDoAdd(){
  fxKeep();
  var txt=String(fxNew||'').trim(); if(!txt||fxBusy) return;
  fxBusy=true;
  google.script.run.withSuccessHandler(function(r){
    fxBusy=false;
    if(r&&r.err==='SESSION'){ fxTok=null; fxMsg=r.msg; fxStep1(); return; }
    if(r&&r.options) fxOpts=r.options;
    if(r&&r.ok){ if(r.added) fxSel[r.added]=true; fxNew=''; fxMsg=''; } else fxMsg=(r&&r.msg)||'Không thêm được';
    fxStep2();
  }).withFailureHandler(function(e){ fxBusy=false; fxMsg='Lỗi: '+e.message; fxStep2(); }).fxAddOption(fxTok,txt);
}
function fxDoSubmit(){
  fxKeep();
  if(fxBusy) return;
  var chosen=fxOpts.filter(function(o){ return fxSel[o]; });
  if(!chosen.length){ fxMsg='Chọn ít nhất 1 cách sửa'; fxStep2(); return; }
  var tk=fxCur, url=fxUrl(tk), pc=true;
  try{ pc=!isPhone(); }catch(e){}
  fxPopup=null;
  if(pc&&url){ try{ fxPopup=window.open('','fxSheetTab'); }catch(e){ fxPopup=null; } }   // mở sẵn trong lúc bấm để trình duyệt không chặn
  fxBusy=true; fxMsg='';
  fxShow('<div class="fxtt">Đang gửi...</div><div class="fxm0">Đang ghi nhận và gửi email, vui lòng chờ.</div>');
  function dropPopup(){ try{ if(fxPopup) fxPopup.close(); }catch(e){} fxPopup=null; }
  google.script.run.withSuccessHandler(function(r){
    fxBusy=false;
    if(r&&r.ok){
      var opened=false;
      if(fxPopup&&url){ try{ fxPopup.location.href=url; opened=true; }catch(e){ opened=false; } }
      fxPopup=null;
      var link=url?('<a class="fxgo" style="display:block;text-align:center;text-decoration:none" target="_blank" rel="noopener" href="'+fxEsc(url)+'">📄 Mở Google Sheet tại ô '+fxEsc(fxLoc(tk))+'</a>'):'';
      fxShow('<div class="fxtt">Đã ghi nhận ✓</div><div class="fxok">'+(r.emailSent?'📧 Đã gửi email cho quản lý':'⚠️ Chưa gửi được email')+'</div>'+
        (opened?'<div class="fxm0" style="margin-top:8px">Google Sheet đã mở ở tab mới, đúng ô '+fxEsc(fxLoc(tk))+'.</div>':link)+
        '<button class="fxcl" data-act="close">ĐÓNG</button>');
      fxLast=''; fxPoll();
      if(opened) setTimeout(function(){ if(!fxBusy) fxClose(); },3500);
    } else if(r&&r.err==='SESSION'){ dropPopup(); fxTok=null; fxMsg=r.msg; fxStep1(); }
    else { dropPopup(); fxMsg=(r&&r.msg)||'Gửi thất bại'; fxStep2(); fxPoll(); }
  }).withFailureHandler(function(e){ dropPopup(); fxBusy=false; fxMsg='Lỗi: '+e.message; fxStep2(); }).fxSubmitFix(fxTok,tk.id,chosen,fxNote);
}
document.getElementById('fxCard').onclick=function(ev){
  var t=ev.target;
  while(t&&t!==this){ if(t.getAttribute&&t.getAttribute('data-act')) break; t=t.parentNode; }
  if(!t||t===this) return;
  var a=t.getAttribute('data-act');
  if(a==='close') fxClose();
  else if(a==='login') fxDoLogin();
  else if(a==='logout'){ var tk=fxTok; fxTok=null; fxWho=null; fxMsg=''; if(tk) google.script.run.fxLogout(tk); fxStep1(); }
  else if(a==='opt'){ fxKeep(); var o=fxOpts[Number(t.getAttribute('data-i'))]; fxSel[o]=!fxSel[o]; fxStep2(); }
  else if(a==='add') fxDoAdd();
  else if(a==='submit') fxDoSubmit();
};
document.getElementById('fxCard').onkeydown=function(ev){
  if(ev.keyCode!==13||!ev.target) return;
  if(ev.target.id==='fxPin'||ev.target.id==='fxId'){ ev.preventDefault(); fxDoLogin(); }
  else if(ev.target.id==='fxNew'){ ev.preventDefault(); fxDoAdd(); }
};
google.script.run.withSuccessHandler(function(l){ fxLinks=l; fxLast=''; fxRender(); }).fxSheetLinks();
fxPoll();
setInterval(fxPoll,5000);
</script>
`;
}

// Chỉ dùng để KIỂM TRA EMAIL. Không đổi gì ở code đang chạy.
// Cách dùng: dán hàm này xuống cuối file (hoặc tạo file mới trong Apps Script),
// chọn hàm  kiemTraEmail  → Chạy → cho phép quyền nếu Google hỏi → mở "Nhật ký thực thi".
// Xong có thể xóa hàm đi. Hàm chỉ gửi 1 email thử cho CHÍNH tài khoản chạy script, không gửi cho quản lý.
function kiemTraEmail() {
  var me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  Logger.log('Tài khoản chạy script: ' + (me || '(không đọc được)'));
  try {
    Logger.log('Danh sách nhận cảnh báo (ALERT_EMAILS): ' + ALERT_EMAILS + '  → ' + String(ALERT_EMAILS).split(',').length + ' người');
  } catch (e0) { Logger.log('Không đọc được ALERT_EMAILS: ' + e0.message); }
  try {
    Logger.log('Hạn mức email còn lại hôm nay: ' + MailApp.getRemainingDailyQuota() +
      '  (mỗi người nhận tính 1 lượt; Gmail thường ~100/ngày, Workspace ~1500/ngày)');
  } catch (e1) { Logger.log('Không đọc được hạn mức: ' + e1.message); }
  if (!me) { Logger.log('Không có địa chỉ để gửi thử.'); return; }
  try {
    MailApp.sendEmail({ to: me, subject: 'Thử gửi email – SCAN SYSTEM', body: 'Nếu bạn đọc được email này thì việc gửi mail từ script đang hoạt động.' });
    Logger.log('GỬI THỬ THÀNH CÔNG tới ' + me);
  } catch (e2) {
    Logger.log('GỬI THỬ THẤT BẠI: ' + e2.message);
  }
}

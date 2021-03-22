let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シフト');
let set_date = sheet.getRange(10, 1).getValue();
let last_date = new Date(set_date.getFullYear(), set_date.getMonth()+1, 0);
let last_day = last_date.getDate()
let all_shifts = sheet.getSheetValues(12, 2, 1, last_day); 
let all_comments = sheet.getRange("B12:AF12"); 


let today = new Date(set_date.getFullYear(), set_date.getMonth(), 1);
let year = today.getFullYear();
let month = 1 + today.getMonth();

function firstCheck() {
  const result = all_shifts[0].filter(null_shift => null_shift == "");
  if (result.length) {
    nullCheck_()
  } else {
    lastCheck_()
  }
}

function nullCheck_() {
  all_shifts[0].forEach( function(shift, index) {
    if (shift == "") {
      Browser.msgBox("未入力のシフトがあります。" + month + "/" + (index + 1) + "のシフトに値を入力しましょう。");
    }
  });
}

function lastCheck_() {
  var check = Browser.msgBox("【" + month + "月】のGoogleカレンダーへ自動入力を行います。", "続行しますか？", Browser.Buttons.OK_CANCEL);
  if (check == 'ok') {
    autoShift_();
    Browser.msgBox("完了しました。");
  }
  if (check == 'cancel') {
    Browser.msgBox("処理はキャンセルされました。");
  }
}

function autoShift_() {
  all_shifts[0].forEach( function(shift, index) {
    judgeCreate(shift, index + 1)
  });

  var all_comments_shifts = all_comments.getNotes()
  all_comments_shifts.forEach( function(comments) {
    comments.forEach( function(comment, index) {
      if (comment != "") {
        judgeCreate(comment, index + 1)
      } 
    }) 
  });
}

function judgeCreate(shift, index) {
  Utilities.sleep(1000);
    if (shift == "基10-19") {
      kiso10_19_(index);
    } else if (shift == "基11-20") {
      kiso11_20_(index);
    } else if (shift == "基中11-22") {
      kiso11_22_(index);
    } else if (shift == "基14-19") {
      kiso14_19_(index);
    } else if (shift == "基14-22") {
      kiso14_22_(index);
    } else if (shift == "基10-13") {
      kiso10_13_(index);
    } else if (shift == "応10-19") {
      ouyo10_19_(index);
    } else if (shift == "応11-20") {
      ouyo11_20_(index);
    } else if (shift == "応中11-22") {
      ouyo11_22_(index);
    } else if (shift == "応14-22") {
      ouyo14_22_(index);
    } else if (shift == "応19-22") {
      ouyo19_22_(index);
    } else if (shift == "応10-13") {
      ouyo10_13_(index);
    } else if (shift == "応11-13") {
      ouyo11_13_(index);
    } else if (shift == "終10-19" || shift == "終中10-19") {
      saishu10_19_(index);
    } else if (shift == "終11-20") {
      saishu11_20_(index);
    } else if (shift == "終中11-22") {
      saishu11_22_(index);
    } else if (shift == "終14-19" || shift == "終中14-19") {
      saishu14_19_(index);
    } else if (shift == "終14-22") {
      saishu14_22_(index);
    } else if (shift == "終10-13") {
      saishu10_13_(index);
    } else if (shift == "終10-16") {
      saishu10_16_(index);
    } else if (shift == "終19-22") {
      saishu19_22_(index);
    } else if (shift == "C10-13" || shift == "C&総") {
      chat10_13_(index);
    } else if (shift == "C10-19") {
      chat10_19_(index);
    } else if (shift == "C19-22") {
      chat19_22_(index);
    } else if (shift == "C14-22") {
      chat14_22_(index);
    } else if (shift == "C18-22") {
      chat18_22_(index);
    } else if (shift == "R10-13") {
      review10_13_(index);
    } else if (shift == "R10-19" || shift == "R中10-19") {
      review10_19_(index);
    } else if (shift == "R14-19") {
      review14_19_(index);
    } else if (shift == "R18-22") {
      review18_22_(index);
    } else if (shift == "R19-22") {
      review19_22_(index);
    } else if (shift == "R14-22" || shift == "R中14-22") {
      review14_22_(index);
    } else if (shift == "F10-13") {
      follow10_13_(index);
    } else if (shift == "F10-19") {
      follow10_19_(index);
    } else if (shift == "F14-19") {
      follow14_19_(index);
    } else if (shift == "F14-22") {
      follow14_22_(index);
    } else if (shift == "外A") {
      shiftgaiA_(index);
    } else if (shift == "外B") {
      shiftgaiB_(index);
    } else if (shift == "外C") {
      shiftgaiC_(index);
    } else if (shift == "公" || shift == "有" || shift == "年") {
      holiday_(index);
    } else if (shift == "総会") {
      soukai_(index);
    } else {
      Browser.msgBox(month + "/" + index + "で既定値以外の値が検出されました。お手数ですが、こちらはご自身で手動入力お願い致します。");
  }
}

function kiso10_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('基10-19', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function kiso11_20_(i) {
  CalendarApp.getDefaultCalendar().createEvent('基11-20', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/20:00'));
}

function kiso11_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('基中11-13', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  CalendarApp.getDefaultCalendar().createEvent('外16-19', new Date(year + '/' + month + '/' + i + '/16:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
  CalendarApp.getDefaultCalendar().createEvent('基中19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function kiso14_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('基14-19', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function kiso14_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('基14-22', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function kiso10_13_(i) {
  CalendarApp.getDefaultCalendar().createEvent('基10-13', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function ouyo10_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応10-19', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function ouyo11_20_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応11-20', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/20:00'));
}

function ouyo11_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応中11-13', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  CalendarApp.getDefaultCalendar().createEvent('外16-19', new Date(year + '/' + month + '/' + i + '/16:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
  CalendarApp.getDefaultCalendar().createEvent('応中19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function ouyo14_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応14-22', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function ouyo19_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function ouyo10_13_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応10-13', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function ouyo11_13_(i) {
  CalendarApp.getDefaultCalendar().createEvent('応11-13', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function saishu10_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終10-19', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function saishu11_20_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終11-20', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/20:00'));
}

function saishu11_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終中11-13', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  CalendarApp.getDefaultCalendar().createEvent('外16-19', new Date(year + '/' + month + '/' + i + '/16:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
  CalendarApp.getDefaultCalendar().createEvent('終中19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function saishu14_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終14-22', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function saishu14_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終14-19', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function saishu10_13_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終10-13', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function saishu10_16_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終10-16', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/16:00'));
}

function saishu19_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('終19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function chat10_13_(i) {
  CalendarApp.getDefaultCalendar().createEvent('C10-13', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function chat10_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('C10-19', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function chat19_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('C19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
} 

function chat14_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('C14-22', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function chat18_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('C18-22', new Date(year + '/' + month + '/' + i + '/18:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function review10_13_(i) { 
  CalendarApp.getDefaultCalendar().createEvent('R10-13', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function review10_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('R10-19', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function review14_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('R14-19', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function review18_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('R18-22', new Date(year + '/' + month + '/' + i + '/18:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function review19_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('R19-22', new Date(year + '/' + month + '/' + i + '/19:00'), new Date(year + '/' + month + '/' + i + '/22:00'));    
}

function review14_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('R14-22', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function follow10_13_(i) {
  CalendarApp.getDefaultCalendar().createEvent('F10-13', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
}

function follow10_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('F10-19', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function follow14_19_(i) {
  CalendarApp.getDefaultCalendar().createEvent('F14-19', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function follow14_22_(i) {
  CalendarApp.getDefaultCalendar().createEvent('F14-22', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function shiftgaiA_(i) {   
  CalendarApp.getDefaultCalendar().createEvent('外A', new Date(year + '/' + month + '/' + i + '/10:00'), new Date(year + '/' + month + '/' + i + '/19:00'));
}

function shiftgaiB_(i) {   
  CalendarApp.getDefaultCalendar().createEvent('外B（前半）', new Date(year + '/' + month + '/' + i + '/11:00'), new Date(year + '/' + month + '/' + i + '/13:00'));
  CalendarApp.getDefaultCalendar().createEvent('外B（後半）', new Date(year + '/' + month + '/' + i + '/16:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function shiftgaiC_(i) {   
  CalendarApp.getDefaultCalendar().createEvent('外C', new Date(year + '/' + month + '/' + i + '/14:00'), new Date(year + '/' + month + '/' + i + '/22:00'));
}

function holiday_(i) {
  let event = CalendarApp.getDefaultCalendar().createAllDayEvent('休日', new Date(year + '/' + month + '/' + i));
  event.setColor(CalendarApp.EventColor.ORANGE);
}

function soukai_(i) {   
  CalendarApp.getDefaultCalendar().createEvent('総会', new Date(year + '/' + month + '/' + i + '/14:30'), new Date(year + '/' + month + '/' + i + '/18:00'));
}

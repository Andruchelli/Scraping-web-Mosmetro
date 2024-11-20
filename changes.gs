const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current"); // лист, откуда мы берём значения
const ws_copy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Previous"); // лист, куда мы копируем значения
const ws_compare = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Differents"); // лист, куда мы заносим данные, содержащиеся только в одном из диапазонов (новые или изменённые данные)
const err = ws_copy.getRange("E1"); // ячейка E1 на втором листе, содержащая в себе переменную, в которой будет храниться сообщение об ошибке (даты двух массивов будут одинаковыми)

// function getData() { // первый лист
  // const todayDate = Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy"); // текущая дата
  // var range = ws.getRange("A2");
  // clear = range.setValue(""); // либо range.clearContent(); очищаем ячейку A2 на первом листе, где лежит формула с ссылкой на страницу с графиком ремонтов
  // Utilities.sleep(3000)
  // range.setFormula('=IMPORTHTML("https://mosmetro.ru/passengers/information/works/"; "table"; 0)'); // добавляем ссылку, откуда будут браться данные
  // date = ws.getRange("A1").setValue(todayDate); // добавляем текущую дату в ячейку A1 на первом листе 
  // date_pars = ws.getRange("A1").getDisplayValues(); // сохраняем в переменную date_pars видимое значение ячейки с датой
  // Utilities.sleep(3000)
// }

function copyData() { // второй лист
  const data = ws.getRange(3, 1, ws.getLastRow(), 1).getDisplayValues(); // данные, которые мы получаем со страницы сайта (со второй (2) строки, без заголовков столбцов)
  const no_data = ws.getRange("A3").getValue(); // первая ячейка полученных данных для проверки, получили ли мы данные с страницы или нет
  const date_A = ws_copy.getRange("A1").getValue(); // дата первого диапазона
  const date_B = ws_copy.getRange("B1").getValue(); // дата второго диапазона
  const current_date = ws.getRange("A1").getValue(); // дата, в которую собирались данные с сайта
  console.log(current_date)
  var diff_date = (date_A.getTime()) - (date_B.getTime()); // считаем разницу между двумя датами
  console.log(diff_date)

  if (diff_date < 0) {
    copy_clear_1 = ws_copy.getRange(3, 1, ws_copy.getLastRow(), 1).setValue(""); // очищаем заполненные ячейки под новые данные, если дата слева (A1) меньше, чем дата справа (B1)
    copy_data_1 = ws_copy.getRange(3, 1, data.length, 1).setValues(data);  // переносим данные, которые сохранились в переменную data на лист "Сравнение" в необходимый диапазон, согласно имеющимся датам
    copy_date_A = ws_copy.getRange(1, 1).setValue(current_date); // меняем дату в ячейке A1 на ту, что указана на первом (1) листе
    clear_error = err.setValue(""); // очищаем ячейку с ошибкой (ячейка E1 на втором листе), так как ошибки нет
  }
  if (diff_date > 0) {
    copy_clear_2 = ws_copy.getRange(3, 2, ws_copy.getLastRow(), 1).setValue(""); // очищаем заполненные ячейки под новые данные, если дата слева (A1) больше, чем дата справа (B1)
    copy_data_2 = ws_copy.getRange(3, 2, data.length, 1).setValues(data); // переносим данные, которые сохранились в переменную data на лист "Сравнение" в необходимый диапазон, согласно имеющимся датам
    copy_date_B = ws_copy.getRange(1, 2).setValue(current_date); // меняем дату в ячейке B1 на ту, что указана на первом (1) листе
    clear_error = err.setValue(""); // очищаем ячейку с ошибкой (ячейка E1 на втором листе), так как ошибки нет
  }
  if (diff_date == 0 || no_data == "Нет данных." || no_data == "") {
    ws_copy.getRange(1, 5).setValue("Ошибка"); // в случае, если даты в обеих ячейках (A1 и B1) равны, то записываем в ячейку E1, где хранится сообщение об ошибке, слово "Ошибка"
    //Utilities.sleep(1000);
  }
}

function compare() { // третий лист
  clear_data_comp = ws_compare.getRange(3, 1, ws_compare.getLastRow(), 2).setValue(""); // перед добавлением новых данных, найденных в ходе обработки обоих диапазонов, очищаем старые данные 
  compare_date_A = ws_copy.getRange("A1").getValue();
  ws_compare.getRange(1, 1).setValue(compare_date_A); // в ячейку A1 на третьем листе переносим дату из левого диапазона на втором листе
  compare_date_B = ws_copy.getRange("B1").getValue();
  ws_compare.getRange(1, 2).setValue(compare_date_B); // в ячейку B1 на третьем листе переносим дату из правого диапазона на втором листе
  data_1 = ws_copy.getRange(3, 1, ws_copy.getLastRow(), 1).getValues(); // первый диапазон данных
  //console.log(data_1)
  data_2 = ws_copy.getRange(3, 2, ws_copy.getLastRow(), 1).getValues(); // второй диапазон данных
  //console.log(data_2)

  // values_1 = data_1.map(function(e, i) {return e[0] != data_2[i][0] ? [data_1[i][0]] : [""]}); // находим данные, которые есть в первом диапазоне и отсутствуют во втором
  // values_2 = data_2.map(function(o, i) {return o[0] != data_1[i][0] ? [data_2[i][0]] : [""]}); // находим данные, которые есть во втором диапазоне и отсутствуют в первом
  // val_1 = values_1.filter(function(value) { // собираем массив из данных первого диапазона, где хранятся отличающиеся данные без пустых строк
  //   return value != "";
  // })
  // console.log(val_1)
  // val_2 = values_2.filter(function(value) { // собираем массив из данных второго диапазона, где хранятся отличающиеся данные без пустых строк
  //   return value != "";
  // })
  // console.log(val_2)

  val_1 = [];
  val_1 = data_1.filter(x => !data_2.flat().includes(x[0])); // находим данные, которые есть в первом диапазоне и отсутствуют во втором
  //console.log(val_1)

  val_2 = [];
  val_2 = data_2.filter(y => !data_1.flat().includes(y[0])); // находим данные, которые есть во втором диапазоне и отсутствуют в первом
  //console.log(val_2)

  // колонка таблицы - это массив единичных массивов вида [ [row1], [row2], [row3] ], так как мы проверяем только вхождение и больше ничего, то контрольный массив можно сделать "плоским"

  if (val_1 != "") {
    ws_compare.getRange(3, 1, val_1.length, val_1[0].length).setValues(val_1); // выводим данные из массива отличающихся данных первого диапазона на третий лист
  }
  else {
    ws_compare.getRange(3, 1).setValue("Изменений нет.");
  }
  if (val_2 != "") {
    ws_compare.getRange(3, 2, val_2.length, val_2[0].length).setValues(val_2); // выводим данные из массива отличающихся данных второго диапазона на третий лист
  }
  else {
    ws_compare.getRange(3, 2).setValue("Изменений нет.");
  }

}
function message() { // формируем запись для отправки в Telegram
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }; // параметры для красивого вывода даты в формате: день недели, число, месяц, год
  const clientIdChat = "///";
  date_1 = ws_compare.getRange(1, 1).getValue();
  date_1_parse = date_1.toLocaleDateString('ru-RU', options);
  console.log(date_1_parse)
  date_2 = ws_compare.getRange(1, 2).getValue();
  date_2_parse = date_2.toLocaleDateString('ru-RU', options);
  console.log(date_2_parse)
  data_T1 = ws_compare.getRange(3, 1, ws_compare.getLastRow(), 1).getValues().map(row => row.join('\n')).join('\n\n'); // формируем массив данных по первому диапазону для отправки в Телеграм (все ячейки целиком)
  data_T2 = ws_compare.getRange(3, 2, ws_compare.getLastRow(), 1).getValues().map(row => row.join('\n')).join('\n\n'); // формируем массив данных по второму диапазону для отправки в Телеграм (все ячейки целиком)
  // data_T1 = ws_compare.getRange(3, 1, ws_compare.getLastRow()-1, 1).getValues().map(row => row[0].split('\n\n')[2]).join('\n\n'); // формируем массив данных по первому диапазону для отправки в Телеграм (только второй элемент из каждой ячейки)
  // data_T2 = ws_compare.getRange(3, 2, ws_compare.getLastRow()-1, 1).getValues().map(row => row[0].split('\n\n')[2]).join('\n\n'); // формируем массив данных по второму диапазону для отправки в Телеграм (только второй элемент из каждой ячейки)
  mess_1 = `Изменения за <u>${date_2_parse}</u>\n\n${data_T2}`; // формирование сообщения
  mess_2 = `Изменения за <u>${date_1_parse}</u>\n\n${data_T1}`; // формирование сообщения
  mess_err = `Ошибка. Одинаковые даты на втором листе или отсутствие данных. Просьба проверить и очистить ячейку E1 второго листа.`; // формирование сообщения об ошибке
  //mess = "Изменения за <u>${date_1_parse}</u>\n\n${data_T1}Изменения за <u>${date_2_parse}</u>\n\n${data_T2}";
  if (err.getValue() == "Ошибка") {
    console.log(err.getValue())
    sendText(clientIdChat, mess_err);
  }
  else if (date_1 - date_2 > 0) { // определяем разницу между датами, чтобы в первом сообщении выводились данные по более ранней дате, а во втором сообщении данные за следующий день
    console.log(mess_1.length)
    console.log(mess_2.length)
    sendText(clientIdChat, mess_1);
    Utilities.sleep(1000);
    sendText(clientIdChat, mess_2);
  }
  else if (date_1 - date_2 < 0) {
    console.log(mess_2.length)
    console.log(mess_1.length)
    sendText(clientIdChat, mess_2);
    Utilities.sleep(1000);
    sendText(clientIdChat, mess_1);
  }
}

function sendText(clientIdChat, text) { // функция для отправки сообщения в телеграм
  const token = "///"; // токен чат-группы в телеграме
  let data_T = {
    method: 'sendMessage',
    chat_id: String(clientIdChat),
    text: text,
    parse_mode: 'HTML'
  };
  let options = {
    method: 'post',
    payload: data_T
  };
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/', options)
}

//getData()
copyData()
compare()
message()
sendText()

// https://www.youtube.com/watch?v=GilBmW050rI - про добавление формулы в ячейку
// https://www.youtube.com/watch?v=v2Ipv572CV4 - про перенос данных с одного листа на другой
// https://www.youtube.com/watch?v=-fB1wRJh6F4 - про отправку данных из таблицы в Telegram
// https://www.youtube.com/watch?v=MR10T4WPBmc - про отправку данных из таблицы в Telegram
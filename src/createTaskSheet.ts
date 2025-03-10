// クライアントごとのタスクシートを作成し、タスクリストを更新
function createClientSheet(): void{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const original = ss.getSheetByName("0_タスク原本")!;

  const clientLists = getClients();
  if(!clientLists) return;
  const allData = getAllTasks();
  console.log(allData)  //全データ　4列目がクライアント名

  for(const list of clientLists){
    if(!ss.getSheetByName(list)){
      const newSheet = original.copyTo(ss).setName(list);
      const targetData = allData.filter((item) => item[3].toString() === list.toString());
      if(!targetData || targetData.length < 1) continue;
      newSheet.getRange(2,1,targetData.length,targetData[0].length).setValues(targetData);
    } else {
      const targetSheet = ss.getSheetByName(list);
      const targetData = allData.filter((item) => item[3].toString() === list.toString());
      if(!targetData || targetData.length < 1) continue;
      targetSheet?.getRange(2,1,targetData.length,targetData[0].length).setValues(targetData);
    }
  }
}
// クライアントリストを取得
function getClients(): string[]{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("クライアント");
  if(!targetSheet) return [];

  const lastRow = targetSheet.getLastRow();
  if(lastRow < 2) return [];

  return targetSheet.getRange(1,1,lastRow-1,1).getValues().slice(1) as [];
}

type Task = [
  taskNo: number,
  taskTitle: string, 
  isClear: string, 
  client:string, 
  project: string,
  relatedTaskNo: number,
  relatedTaskTitle: string,
  minValue: number,
  realValue: number,
  maxValue: number,
];
// 全てのタスクを取得
function getAllTasks(): Task[] {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName("タスク見積もり")!;
  
  return baseSheet.getDataRange().getValues().slice(1).filter((item) => item[1] !== "") as Task[];
}

// タスクの通し番号の中で最大値を返す
function getMaxTaskNumber(): number {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("タスク見積もり");
  if (!sheet) return 0;
  
  const lastRow = sheet.getRange(1,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
  if (lastRow < 2) return 0;

  // ヘッダー行を除外して2行目から取得
  const taskNumbers = sheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(num => typeof num === 'number' && !isNaN(num)) as number[];
  
  return taskNumbers.length > 0 ? Math.max(...taskNumbers) : 0;
}

function onEdit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeCell = activeSheet.getActiveCell();
  const activeRow = activeCell.getRow();
  const activeCol = activeCell.getColumn();
  const activeValue = activeCell.getValue();
  
  // 原本シートでの編集を防ぐ
  if (activeSheet.getName().includes("0_")) {
    console.log("原本シートでは編集できません");
    return;
  }

  // タスクタイトル列（2列目）に値が入力された場合のみ処理を実行
  if (activeCol === 2 && activeValue !== "") {
    const nowTaskNumber = getMaxTaskNumber();
    const newTaskNumber = nowTaskNumber + 1;
    
    // タスク番号が既に存在しないことを確認
    const currentTaskNumber = activeCell.offset(0, -1).getValue();
    if (!currentTaskNumber) {
      const nowDate = Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss");
      activeCell.offset(0, -1).setValue(newTaskNumber);
      activeCell.offset(0, 11).insertCheckboxes(); //完了・削除のためのチェックボックス
      activeCell.offset(0, 14).insertCheckboxes(); //カレンダー連携のためのチェックボックス
      activeCell.offset(0, 15).setValue(nowDate); //タイムスタンプをset
    }
  }

  // O列の予定セットにチェックが入ったらカレンダーにセット
  if(activeCol === 15 && activeValue === true){
    setTaskCalendar(activeRow);
  } 
}

function setTaskCalendar(row: number){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const lastColumn = activeSheet.getLastColumn();
  const data: (string | number)[][] = activeSheet.getRange(row, 1, 1, lastColumn).getValues();
  
  const calendar = data[3].toString();
  const calendarId = getCalendarId(calendar);

  const taskTitle = data[1].toString();
  const taskDetail = data[2].toString();
  const requiredTime = parseInt(data[10].toString());
  const startDateStr = data[13].toString() + " " + data[14].toString();
  
  const start = Utilities.parseDate(startDateStr, "JST", "yyyy/MM/dd HH:mm")
  const endDate = new Date(start);
  endDate.setMinutes(endDate.getMinutes() + requiredTime)

  CalendarApp.getCalendarById(calendarId).createEvent(taskTitle, start, endDate, {description: taskDetail});
}

function getCalendarId(calendarName: string): string{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName("クライアント");
  const data: string[][] = clientSheet?.getDataRange().getValue().slice(1);

  const calendarList = data.filter((rowItem) => rowItem[1] === calendarName)
  return calendarList[1].toString();
} 
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
      activeCell.offset(0, 13).insertCheckboxes(); //完了・削除のためのチェックボックス
      activeCell.offset(0, 16).insertCheckboxes(); //カレンダー連携のためのチェックボックス
      activeCell.offset(0, 17).setValue(nowDate); //タイムスタンプをset
    }
  }
}

function setTaskCalendar(row: number){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const lastColumn = activeSheet.getLastColumn();
  const data: (string | number)[][] = activeSheet.getRange(row, 1, 1, lastColumn).getValues();
  
  const calendar = data[0][3] as string;
  const calendarId = getCalendarId(calendar);

  const taskTitle = data[0][1] as string;
  const taskDetail = data[0][2] as string;
  const requiredTime = Number(data[0][10]);
  const startDate = Utilities.formatDate(new Date(data[0][13]), "JST", "yyyy/MM/dd")
  const startTime = Utilities.formatDate(new Date(data[0][14]), "JST", "HH:mm")
  const startDateStr = startDate + " " + startTime;
  
  const start = Utilities.parseDate(startDateStr, "JST", "yyyy/MM/dd HH:mm")
  const endDate = new Date(start);
  endDate.setMinutes(endDate.getMinutes() + requiredTime)

  CalendarApp.getCalendarById(calendarId).createEvent(taskTitle, start, endDate, {description: taskDetail});
  activeSheet.getRange(row, 16).setValue(false);
}

function getCalendarId(calendarName: string): string {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName("クライアント");
  if (!clientSheet) return "";
  
  const data = clientSheet.getDataRange().getValues().slice(1) as string[][];
  const calendarList = data.filter((rowItem) => rowItem[0] === calendarName);
  // console.log(calendarList);
  return calendarList[0][1];
} 

function onEditCalendar(e: any){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeCell = activeSheet.getActiveCell();
  const col = activeCell.getColumn();
  const val = activeCell.getValue();
  if(col !== 16 || val !== true) return;

  const row = activeCell.getRow();
  setTaskCalendar(row);
}
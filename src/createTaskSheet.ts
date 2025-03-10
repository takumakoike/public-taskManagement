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

  return targetSheet.getRange(1,1,lastRow-1,1).getValues() as [];
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
  const lastRow = sheet?.getRange(1,1).getLastRow();
  if(!lastRow || lastRow < 2) return 0;

  const taskNumbers = sheet?.getRange(2, 1, lastRow - 1, 1).getValues().flat() as number[];
  
  return Math.max.apply(null, taskNumbers);
}

// アクティブシートにタスクを追加し始めたら通し番号を振る
function setTaskNumber(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeCell = activeSheet.getActiveCell();
  const activeCol = activeCell.getColumn();
  const activeValue: string = activeCell.getValue();
  if(activeSheet.getName().includes("0_")) return;

  const nowTaskNumber = getMaxTaskNumber();
  if( activeCol === 2 && activeValue !== ""){
    activeCell.offset(0, -1).setValue(nowTaskNumber + 1);
  }
}
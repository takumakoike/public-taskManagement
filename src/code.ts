// クライアントごとのタスクシートを作成し、タスクリストを更新
function createClientSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const original = ss.getSheetByName("タスク原本")!;

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


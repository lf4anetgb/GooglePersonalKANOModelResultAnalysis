/** @OnlyCurrentDoc */

function start() {
//  取得控制格
  let spreadsheet = SpreadsheetApp.getActive();
//  取得Y軸終點，並用於顯示結果
  let yEnd = getStartShowRange(spreadsheet);
  

//  設定正面題目列位子，如第一題D第二題E...
  let questionsOkSeat = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
//  設定反面題目列位子，如第一題L第二題M...
  let questionsNoSeat = ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'];
  
//  判斷是否有手誤
  if(tryError(spreadsheet, yEnd, questionsOkSeat, questionsNoSeat)){
    return;
  }
  
//  設定克存儲空間
  let questionsStatus = new Array(questionsOkSeat.length); // 各題各需求存儲空間E:0 M:1 L:2 R:3 I:4 Q:5
  let questionsWinner = new Array(questionsOkSeat.length); // 各題贏家
  
//  初始化
  for(let i = 0; i < questionsStatus.length; i++){
    questionsStatus[i] = [0,0,0,0,0,0];
  }
  for(let i = 0; i < questionsWinner.length; i++){
    questionsWinner[i] = '';
  }
  
//  統計各需求型態數值
  questionStatus(spreadsheet, questionsStatus, yEnd, questionsOkSeat, questionsNoSeat);

//  判斷各題贏家
  confirmWinner(questionsStatus, questionsWinner);
  
//  顯示
  show(spreadsheet, yEnd, questionsStatus, questionsOkSeat, questionsWinner);
  
};

//取得起始顯示位子
function getStartShowRange(spreadsheet){
  
  let y = 1; // 設定初始位子
  
//  遍歷A列直到那格沒字串
  while(spreadsheet.getRange('A' + y).getValue().toString().length != 0){
    y++;
  }
  
  return y; // 回傳遍歷結果

};

//  判斷是否有手誤
function tryError(spreadsheet, yEnd, questionsOkSeat, questionsNoSeat){
  
//  檢查各自欄位數是否一樣
  if(questionsNoSeat.length != questionsOkSeat .length){
    spreadsheet.getRange('A' + yEnd).setValue('檢查是否有題目位子少Key');
    return true;
  }
  
//  檢查位子是否重疊
  for(let i = 0; i < questionsOkSeat.length; i++){
    for(let j = 0; j< questionsNoSeat.length; j++){
      if(questionsOkSeat[i].indexOf(questionsNoSeat[j]) != -1){
        spreadsheet.getRange('A' + yEnd).setValue('檢查是否有正反題目位子重疊');
        return true;
      }
    }
  }
  
  return false;
};

//判斷單次需求型態
function questionStatus(spreadsheet, questionsStatus, yEnd, questionsOkSeat, questionsNoSeat){
  let yNow = 2; // 勿動
  
  while(yNow < yEnd){ // 讀所有使用者
    for(let i = 0; i < questionsOkSeat.length; i++){ // 讀所有題目
      // 讀所有答案
      let Ok_ = spreadsheet.getRange(questionsOkSeat[i] + yNow).getValue();
      let No_ = spreadsheet.getRange(questionsNoSeat[i] + yNow).getValue();
      
      // 判斷不要的答案，再判斷要的答案
      switch(No_){
        //不要等級1
        case 1:
        case 2:
          switch(Ok_){
            case 1:
            case 2:
              questionsStatus[i][5]++; //Q
              break;
            case 6:
            case 7:
              questionsStatus[i][2]++; //L
              break;
            default:
              questionsStatus[i][1]++; //M
          }
          break;

        //不要等級2,3,4
        case 3:
        case 4:
        case 5:
          switch(Ok_){
            case 1:
            case 2:
              questionsStatus[i][3]++; //R
              break;
            case 6:
            case 7:
              questionsStatus[i][0]++; //E
              break;
            default:
              questionsStatus[i][4]++; //I
          }
          break;

        //不要等級5
        case 6:
        case 7:
          switch(Ok_){
            case 6:
            case 7:
              questionsStatus[i][5]++; //Q
              break;
            default:
              questionsStatus[i][3]++; //R
          }
      }
    }
    yNow++;
  }
};

//  判斷各題贏家
function confirmWinner(questionsStatus, questionsWinner){
  
  for(let i = 0; i < questionsStatus.length; i++){
    let winner = Math.max(...questionsStatus[i]); //找出最大值
    
    for(let j = 0; j < questionsStatus[i].length; j++){
      
      if(questionsStatus[i][j] == winner){
        
        switch(j){
          case 0:
            questionsWinner[i] += 'E ';
            break;
          case 1:
            questionsWinner[i] += 'M ';
            break;
          case 2:
            questionsWinner[i] += 'L ';
            break;
          case 3:
            questionsWinner[i] += 'R ';
            break;
          case 4:
            questionsWinner[i] += 'I ';
            break;
          case 5:
            questionsWinner[i] += 'Q ';
            break;
        }
      }
    }
  }
};

//  顯示
function show(spreadsheet, yEnd, questionsStatus, questionsOkSeat, questionsWinner){
  let total = yEnd - 2; // 讀取總人數
  
//  顯示初始化
  yEnd++;
  spreadsheet.getRange('B' + yEnd).setValue('E');
  spreadsheet.getRange('C' + yEnd).setValue('M');
  spreadsheet.getRange('D' + yEnd).setValue('L');
  spreadsheet.getRange('E' + yEnd).setValue('R');
  spreadsheet.getRange('F' + yEnd).setValue('I');
  spreadsheet.getRange('G' + yEnd).setValue('Q');
  spreadsheet.getRange('H' + yEnd).setValue('結果');
  yEnd++;
  for(let i = 0; i < questionsStatus.length; i++ , yEnd++){
    spreadsheet.getRange('A' + yEnd).setValue(spreadsheet.getRange(questionsOkSeat[i] + '1').getValue());
    
    for(let j = 0; j <= questionsStatus[i].length; j++){
      let x = String.fromCharCode(66 + j);
      
      if(j >= questionsStatus[i].length){
        spreadsheet.getRange(x + yEnd).setValue(questionsWinner[i]);
      }else{
        spreadsheet.getRange(x + yEnd).setValue((questionsStatus[i][j] / total * 100) + '%');
      }
    }
  }
};
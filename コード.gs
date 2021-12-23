function createSchedule() {
/**スプレッドシート取得 */
 const ss         = SpreadsheetApp.getActiveSpreadsheet();
 const inputSheet = ss.getSheetByName('スケジュール');
 const lastRow    = inputSheet.getLastRow(); /**inputSheetの最終行取得 */

/**アドレス取得 */
const usersheet = ss.getSheetByName("担当者");
const user_id = Array.prototype.concat.apply([],usersheet.getRange(2,2,usersheet.getLastRow(),1).getValues()); /**担当者名 */
const user_address = Array.prototype.concat.apply([],usersheet.getRange(2,3,usersheet.getLastRow(),1).getValues()); /**担当者アドレス */

/**タイトル・開始日・終了日取得 */
 const title = Array.prototype.concat.apply([],inputSheet.getRange(6, 2, lastRow, 1).getValues());/**タイトル */
 const plan_id = Array.prototype.concat.apply([],inputSheet.getRange(6, 5, lastRow, 1).getValues());/**カレンダーの予定の識別子 */
 const user = Array.prototype.concat.apply([],inputSheet.getRange(6, 9, lastRow, 1).getValues());  /**担当者名 */
 const start_date = Array.prototype.concat.apply([],inputSheet.getRange(6,6,lastRow,1).getValues());/**開始日 */
 const end_date = Array.prototype.concat.apply([],inputSheet.getRange(6,7,lastRow,1).getValues());/**終了日 */
 const Description = Array.prototype.concat.apply([],inputSheet.getRange(6,8,lastRow,1).getValues());/**備考 */
 const result = Array.prototype.concat.apply([],inputSheet.getRange(6,10,lastRow,1).getValues()); /**ステータス */

 /**カレンダー情報取得（起動者のカレンダー取得） */
 var cal = CalendarApp.getDefaultCalendar(); /**カレンダー */
 var address = Session.getActiveUser().getEmail(); /**起動者のメールアドレス */
 var cnt = 0;
 ss.toast("カレンダー作成開始","【更新ステータス】",3);
 /**各カレンダーに登録・変更 */
 for (i = 0; i < title.length; i++) {
    if(start_date[i]==""||end_date[i]==""||title[i]==""||result[i]=="済") continue;
    if(cnt==23){
      for(let j=5;j>0;j--){
        ss.toast("Googleカレンダーの制限のため、登録停止中（登録再開まで残り："+j+"秒）","【更新ステータス】",1);
        Utilities.sleep(1000);
      }
      ss.toast("登録再開!","【更新ステータス】",1);
      cnt=0;
    }
    /**カレンダーIDが空白の時、新規登録 */
    if(plan_id[i]==""){
      var sd = new Date(start_date[i]);
      var ed = new Date(end_date[i]);
      ed.setDate(ed.getDate()+1);
      /**カレンダー登録 */
      var event = cal.createAllDayEvent(title[i],sd,ed);
      event.setDescription(Description[i]);
      inputSheet.getRange(i+6,5).setValue(event.getId());
      var index = user_id.indexOf(user[i]);
      if(index==-1||user[i]=="") continue;
      var id = user_address[index];
      if(id==address||id=="") continue;
      try {
        event.addGuest(id);
      }
      catch(e){
        Browser.msgBox(user[i]+"さんのアドレスのGoogleへの登録が確認できないため、登録できません\n")
      }
      cnt++;
    }
    /**カレンダーIDがある場合 */
    else{
      /**カレンダー情報の記載 */
      var event = cal.getEventById(plan_id[i]);
      var sd = event.getAllDayStartDate();
      var ed = event.getAllDayEndDate();
      var guests = event.getGuestList();
      var desc = event.getDescription();
      var check = true;
      /**担当者変更 */
      var index = user_id.indexOf(user[i]);
      if(index!=-1){
        var id = user_address[index];
        var check2 = true;
        /**担当者の変更有無確認 */
        if(guests.length>0&&guests[0].getEmail()!=id&&id!=""&&id!=address){
          check = false;
          event.removeGuest(guest[0].getEmail());
          check2 = false;
        }
        else if(guest.length==0&&id!=address) check2 = false;
        if(check2==false){
          try {
            event.addGuest(id);
           }
          catch(e){
            Browser.msgBox(user[i]+"さんのアドレスのGoogleへの登録が確認できないため、登録できません\n")
          }
        }
      }
      var ed2 = end_date[i];
      ed2.setDate(ed2.getDate()+1);
      if(sd != start_date[i]||ed!=ed2||desc!=Description[i]) check =false;
      if(check==true) continue;
      event.setAllDayDates(start_date[i],ed2);
      event.setDescription(Description[i]);
      cnt++;
    } 
 }
 ss.toast("Googleカレンダー登録終了","【更新ステータス】",-1);
}

function onOpen() {
 SpreadsheetApp.getUi()
   .createMenu('スクリプト')
   .addItem('カレンダーに反映', 'createSchedule')
   .addToUi();
}

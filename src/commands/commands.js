Office.onReady(() => {});
Office.initialize = function (reason) {};

function action(event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/dialog.html",
    { height: 20, width: 8 },
    function (result) {
      let dialog = result.value;
      setInterval(() => {
        if (localStorage.getItem("login")) {
          localStorage.removeItem("login");
          dialog.close();

          var startDate = new Date();
          startDate.setDate(startDate.getDate() + 1);
          var endDate = new Date();
          endDate.setDate(endDate.getDate() + 1);
          Office.context.mailbox.displayNewAppointmentForm({
            location: "18楼A-05会议室",
            subject: "我的约会",
            requiredAttendees: ["libing@hiklink.com", "zhangwende@hiklink.com"],
            start: startDate,
            end: endDate,
            body: `
            ------------------------------------------
            发起人：李冰
            会议主题：关于HikLink升级事宜
            会议连接：https://www.baidu.com
            会议ID：123456
            ------------------------------------------`,
          });
        }
      }, 600);
    }
  );
  event.completed();
}

function actionSet(event) {
  Office.context.mailbox.item.subject.setAsync("测试测试!!!");
  Office.context.mailbox.item.requiredAttendees.setAsync(["libing@hiklink.com", "zhangwende@hiklink.com"]);
  var startDate = new Date();
  startDate.setDate(startDate.getDate() + 1);
  Office.context.mailbox.item.start.setAsync(startDate);
  var endDate = new Date();
  endDate.setDate(endDate.getDate() + 1);
  Office.context.mailbox.item.end.setAsync(endDate);
  Office.context.mailbox.item.location.setAsync("18楼A-05会议室");
  Office.context.mailbox.item.body.setAsync(`
  ------------------------------------------
  发起人：李冰
  会议主题：关于HikLink升级事宜
  会议连接：https://www.baidu.com
  会议ID：123456
  ------------------------------------------
  `);
  event.completed();
}

Office.actions.associate("action", action);
Office.actions.associate("actionSet", actionSet);

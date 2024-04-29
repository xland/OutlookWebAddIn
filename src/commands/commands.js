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

          Office.context.mailbox.item.subject.setAsync("测试测试!!!", function (e) {
            if (e.status != Office.AsyncResultStatus.Succeeded) {
              $("#log").text("subject error");
            }
          });
          // Office.context.mailbox.item.requiredAttendees.setAsync(["libing@hiklink.com", "zhangwende@hiklink.com"]);
          var startDate = new Date();
          startDate.setDate(startDate.getDate() + 1);
          Office.context.mailbox.item.start.setAsync(startDate);
          var endDate = new Date();
          endDate.setDate(endDate.getDate() + 1);
          Office.context.mailbox.item.end.setAsync(endDate);
          Office.context.mailbox.item.location.setAsync("18楼A-05会议室");

          Office.context.mailbox.item.addHandlerAsync(Office.EventType.AppointmentTimeChanged, (eventArgs) => {
            Office.context.ui.displayDialogAsync(
              "https://localhost:3000/dialog.html",
              { height: 20, width: 8 },
              function (result) {}
            );
          });
        }
      }, 600);
    }
  );
  event.completed();
}
Office.actions.associate("action", action);

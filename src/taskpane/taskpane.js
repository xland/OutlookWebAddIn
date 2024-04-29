Office.onReady(() => {});
Office.initialize = function (reason) {};

window.addEventListener("DOMContentLoaded", () => {
  let btn1 = document.getElementById("btn1");
  btn1.addEventListener("click", () => {
    const item = Office.context.mailbox.item;
    item.subject.getAsync((e) => {
      document.getElementById("title").innerHTML = e.value;
    });
    item.start.getAsync((e) => {
      document.getElementById("startTime").innerHTML = e.value.toString();
    });
    item.end.getAsync((e) => {
      document.getElementById("endTime").innerHTML = e.value.toString();
    });
  });
  let btn2 = document.getElementById("btn2");
  btn2.addEventListener("click", () => {
    Office.context.mailbox.item.subject.setAsync(``);
    Office.context.mailbox.item.requiredAttendees.setAsync([]);
    Office.context.mailbox.item.start.setAsync(new Date());
    Office.context.mailbox.item.end.setAsync(new Date());
    Office.context.mailbox.item.location.setAsync("");
    Office.context.mailbox.item.body.setAsync(``);
  });
});

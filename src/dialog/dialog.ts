document.getElementById("btn").addEventListener("click", () => {
  // Office.context.ui.messageParent(true.toString());
  // let message = { type: 'dialogMessage', data: 'Hello from dialog!' };
  // Office.context.dialog.messageParent(message);
  // document.getElementById("log").innerHTML = JSON.stringify(Office.context);
  localStorage.setItem("login", true);
});

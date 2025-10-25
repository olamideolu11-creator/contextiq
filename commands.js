function showTaskPane(event) {
  Office.addin.showAsTaskpane()
    .then(() => event.completed())
    .catch(error => {
      console.error("Failed to show taskpane:", error);
      event.completed();
    });
}
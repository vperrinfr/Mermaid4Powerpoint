/// <reference types="office-js" />

Office.onReady(() => {
  // Commands are registered in the manifest
  console.log('Commands initialized');
});

// Function to show the task pane
function showTaskpane(event: any) {
  const officeAddin = (Office as any).addin;
  if (officeAddin && officeAddin.showAsTaskpane) {
    officeAddin.showAsTaskpane();
  }
  event.completed();
}

// Register the function
if ((Office as any).actions) {
  (Office as any).actions.associate('showTaskpane', showTaskpane);
}

export {};

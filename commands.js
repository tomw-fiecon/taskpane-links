
function sayHello(event) {
    Office.context.document.setSelectedDataAsync(
        "Hello World!",
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            }
        }
    );
    event.completed();
}

Office.onReady(() => {
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    Office.actions.associate("openLink", openLink);
    Office.actions.associate("findAbbrevs", findAbbrevs);
});




function openCria() {
    window.open('https://www.cria.fiecon.com', '_blank');
}


// function findAbbrevs(){
    
// }


// Gets some details about the current slide and displays them in a notification.
function findAbbrevs() {
    if (Office.context.document.getSelectedDataAsync) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Some slide details are:', '"' + JSON.stringify(result.value) + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            }
        );
    } else {
        app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
    }
}


async function clearMessage(callback) {
    document.getElementById("message").innerText = "";
    await callback();
  }
  
  function setMessage(message) {
    document.getElementById("message").innerText = message;
  }
  
  // Default helper for invoking an action and handling errors.
  async function tryCatch(callback) {
    try {
      document.getElementById("message").innerText = "";
      await callback();
    } catch (error) {
      setMessage("Error: " + error.toString());
    }
  }

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
    Office.actions.associate("openCria", openCria);
    Office.actions.associate("findAbbrevs", findAbbrevs);
});


function openCria() {
    window.open('https://cria.fiecon.com/', '_blank');
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
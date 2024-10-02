
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

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Task pane initialization code
        document.getElementById("helloButton").onclick = sayHello;
    }
});
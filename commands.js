Office.onReady(() => {
    Office.actions.associate("openCria", openCria);
    Office.actions.associate("findAbbrevs", findAbbrevs);
});


function openCria(event) {
    window.open('https://cria.fiecon.com/', '_blank');
    event.completed();
}


function findAbbrevs(event) {
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
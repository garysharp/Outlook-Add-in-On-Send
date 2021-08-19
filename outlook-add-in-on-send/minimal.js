Office.initialize = function(reason) {}

function onSendHandler(event) {
    event.completed({
        allowEvent: true
    });
}
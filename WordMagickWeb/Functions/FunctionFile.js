// Функцию инициализации необходимо выполнять при каждой загрузке новой страницы.
(function () {
    "use strict"

    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("initialize");
        });
    }
})();

function shuffle(a) {
    let counter = a.length;
    while (counter > 0) {
        let i = Math.floor(Math.random() * counter);
        counter--;
        let tmp = a[counter];
        a[counter] = a[i];
        a[i] = tmp;
    }
    return a;
}

function consoleLogHandler(data) {
    console.log(data);
}

function asyncResultHandler(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        console.log("Async OK!");
    }
    else {
        console.log("Async error: " + asyncResult.error.message);
    }
}

function makeSelectedDataAsyncHandler(coercionType) {
    function asyncHandler() {
        console.log("Start handler");
        Office.context.document.getSelectedDataAsync(
            coercionType, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    // normal processing
                    console.log("Normal processing");
                    result = asyncHandler.successHandler(asyncResult.value);
                    Office.context.document.setSelectedDataAsync(result, asyncResultHandler);
                }
                else {
                    // something wrong, lets try fallback
                    console.log("Fallback processing");
                    asyncHandler.fallbackHandler(asyncResult.error.message);
                }
            });
    }
    asyncHandler.successHandler = console.log;
    asyncHandler.fallbackHandler = console.log;
    return asyncHandler;
}

function changeSelectedDataAsync(changer) {
    textHandler = makeSelectedDataAsyncHandler(Office.CoercionType.Text);
    matrixHandler = makeSelectedDataAsyncHandler(Office.CoercionType.Matrix);

    textHandler.successHandler = changer;
    matrixHandler.successHandler = changer;
    matrixHandler.fallbackHandler = textHandler;

    matrixHandler();
}

function shuffle_chars() {
    changeSelectedDataAsync(function (data) {
        console.log(data);
        return data;
    });
}

function temp() {
    Word.run(function (context) {
        var range = context.document.getSelection();
        context.load(range, 'text');
        return context.sync()
            .then(function () {
                var result = [];
                var words = range.text.split(/\s+/);
                words.forEach(function (el, index, array) {
                    word = el.slice(1, -1).split('');
                    word = shuffle(word).join('');
                    console.log("" + el + " -> " + el[0] + word + el[el.length - 1]);
                    array[index] = el[0] + word + el[el.length - 1];
                    return true;
                });
                res = words.join(' ');
                Office.context.document.setSelectedDataAsync(res, function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                });
            })
            .then(context.sync)
            .then(function () {
                //
            })
            .then(context.sync)
    })
    .catch(errorHandler);
}

//$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
function errorHandler(error) {
    // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}
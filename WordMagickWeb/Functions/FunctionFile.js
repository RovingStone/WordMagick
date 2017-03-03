(function () {
    "use strict"

    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.info("Initialize completed");
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

function asyncResultHandler(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        console.info("asyncResult.status == Succeeded");
    }
    else {
        console.error("asyncResult.status == Failed. Message is: " + asyncResult.error.message);
    }
}

function makeSelectedDataAsyncHandler(coercionType) {
    function asyncHandler() {
        Office.context.document.getSelectedDataAsync(
            coercionType, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    console.info("asyncHandler - execute success handler");
                    result = asyncHandler.successHandler(asyncResult.value);
                    Office.context.document.setSelectedDataAsync(result, asyncResultHandler);
                }
                else {
                    console.info("asyncHandler - execute fallback handler");
                    asyncHandler.fallbackHandler(asyncResult.error.message);
                }
            });
    }
    asyncHandler.successHandler = console.info;
    asyncHandler.fallbackHandler = console.warn;
    return asyncHandler;
}

function changeSelectedDataAsync(changer) {
    textHandler = makeSelectedDataAsyncHandler(Office.CoercionType.Text);
    matrixHandler = makeSelectedDataAsyncHandler(Office.CoercionType.Matrix);

    textHandler.successHandler = changer;
    matrixHandler.successHandler = function (matrix) {
        return matrix.map(function (e) {
            if (Array.isArray(e)) {
                return e.map(changer);
            }
            else {
                console.warn(e);
                return e;
            }
        });
    }
    matrixHandler.fallbackHandler = textHandler;

    // Try matrix handler first, because matrix can be converted to text, but not vice versa
    matrixHandler();
}

function word_replacer(changer, data) {
    return data.toString().replace(/[\wа-яёА-ЯЁ]+/ig, changer);
}

function shuffle_chars() {
    var shuffleChars = word_replacer.bind(null, function (m) {
        return shuffle(m.split('')).join('');
    });
    changeSelectedDataAsync(shuffleChars);
}
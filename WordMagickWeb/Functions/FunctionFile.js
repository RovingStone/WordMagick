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

function replacer(regex, changer, data) {
    return data.toString().replace(regex, changer);
}

function shuffleAll() {
    var shuffleChars = replacer.bind(null, new RegExp("[\\wа-яёА-ЯЁ]+", "ig"), function (match) {
        return shuffle(match.split('')).join('');
    });
    changeSelectedDataAsync(shuffleChars);
}

function shuffleTail() {
    var shuffleChars = replacer.bind(null, new RegExp("\\B[\\wа-яёА-ЯЁ]+", "ig"), function (match) {
        return shuffle(match.split('')).join('');
    });
    changeSelectedDataAsync(shuffleChars);
}

function shuffleInner() {
    var shuffleChars = replacer.bind(null, new RegExp("\\B[\\wа-яёА-ЯЁ]+\\B", "ig"), function (match) {
        return shuffle(match.split('')).join('');
    });
    changeSelectedDataAsync(shuffleChars);
}

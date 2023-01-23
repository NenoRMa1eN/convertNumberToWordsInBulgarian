// |=====================================================================================================================|
// |                                                                                                                     |
// |   Конвертиране на число в словом на български език за Google App Script (GAS)                                       |
// |   Copyright (c) 2023 NenoRMa1eN                                                                                     |
// |                                                                                                                     |
// | ОПИСАНИЕ                                                                                                            |
// |                                                                                                                     |
// |   Конвертира лева в думи в Google Sheets                                                                            |
// |                                                                                                                     |
// | НАЧИН НА ИНСТАЛАЦИЯ                                                                                                 |
// |                                                                                                                     |
// |   1. Маркирайте кода с Ctrl+A. Копирайте с Ctrl+C системния буфер                                                   |
// |   2. Отворете файла в който искате да го използвате.                                                                |
// |   3. От менюто "Разширения" изберете Google App Script.                                                             |
// |   4. От менюто в ляво (срещу Files) изберете "+" и "Script"                                                         |
// |   5. Дайте разбираемо име на файла, например: Конвертиране на число в думи или convertNumberToWords и потвърдете.   |
// |   6. Маркирайте с Ctrl+A всичко, което ви предлага автоматично редактора на скриптове и го изтрйте с DEL.           |
// |   7. Поставете копирания в системния буфер код с Ctrl+V.                                                            |
// |   8. Запишете с Ctrl+S.                                                                                             |
// |                                                                                                                     |
// | НАЧИН НА ИЗПОЛЗВАНЕ                                                                                                 |
// |                                                                                                                     |
// |   В клетката, в която желаете да пише сумата словом можете да използвате новосъздадената формула по следния начин:  |
// |                                                                                                                     |
// |   1. Директно с число:                                                                                              |
// |     = inWords(123) - Връща сто двадесет и три лева.                                                                 |
// |                                                                                                                     |
// |   2. Да подадете стойност от друга клетка:                                                                          |
// |                                                                                                                     |
// |     = inWords(A1)  - Връща числото от клетка A1 в словом лева в текущата клетка.                                    |
// |                                                                                                                     |
// |   Максимална стойност на параметъра inputNum e 999 999 999                                                          |
// |                                               (деветстотин деветдесет и девет милиона                               |
// |                                                деветстотин деветдесет и девет хиляди                                |
// |												деветстотин деветдесет и девет лева)                                 |
// |   Още примери:                                                                                                      |
// |     = inWords(265,14)         - Двеста шестдесет и пет лева и 14 стотинки                                           |
// |     = inWords(-12)            - Минус дванадесет лева                                                               |
// |     = inWords(1,53)           - един лев и 53 стотинки                                                              |
// |     = inWords(0)              - нула лева                                                                           |
// |                                                                                                                     |
// |   Ако някой има идеи как да се унифицира и за други мерни единици ще бъде прекрасно да сподели.                     |
// |                                                                                                                     |
// |=====================================================================================================================|


function convertNumberToWords(inputNumber, grammaticalCase) {

    const ONES_IN_WORDS = ['', 'един ', 'два ', 'три ', 'четири ', 'пет ', 'шест ', 'седем ', 'осем ', 'девет ', 'десет ',
        'единадесет ', 'дванадесет ', 'тринадесет ', 'четиринадесет ', 'петнадесет ', 'шестнадесет ',
        'седемнадесет ', 'осемнадесет ', 'деветнадесет '
    ];

    const ONES_IN_WORDS_CASE_1 = ['', 'един ', 'две ', 'три ', 'четири ', 'пет ', 'шест ', 'седем ', 'осем ', 'девет ', 'десет ',
        'единадесет ', 'дванадесет ', 'тринадесет ', 'четиринадесет ', 'петнадесет ', 'шестнадесет ',
        'седемнадесет ', 'осемнадесет ', 'деветнадесет '
    ];

    let ONES_IN_WORDS_TO_USE = (grammaticalCase == 1) ? ONES_IN_WORDS_CASE_1 : ONES_IN_WORDS

    const DECIMALS_IN_WORDS = ['', '', 'двадесет ', 'тридесет ', 'четиридесет ', 'петдесет ', 'шестдесет ', 'седемдесет ', 'осемдесет ', 'деветдесет '];
    const HUNDREDS_IN_WORDS = ['', 'сто ', 'двеста ', 'триста ', 'четиристотин ', 'петстотин ', 'шестстотин ', 'седемстотин ', 'осемстотин ', 'деветстотин '];
    var numbersInWords = ['един ', 'хиляда ', 'един милион '];
    var numbersInWordsPlural = ['', 'хиляди ', 'милиона '];
    var iterationCounter = inputNumber;
    var convertedNumber = "";

    if (iterationCounter < 1000) {
        if (iterationCounter > 1) {
            var digitGroup = Math.floor(iterationCounter / 100);
            iterationCounter -= digitGroup * 100;
            var tensPlace = Math.floor(iterationCounter / 10);
            iterationCounter -= tensPlace * 10;
            if (tensPlace == 1) {
                tensPlace = 0;
                iterationCounter = iterationCounter + 10;
            }
            if (iterationCounter > 0 && tensPlace > 0) {
                convertedNumber = HUNDREDS_IN_WORDS[digitGroup] + DECIMALS_IN_WORDS[tensPlace] + "и " + ONES_IN_WORDS_TO_USE[iterationCounter] + numbersInWordsPlural[grammaticalCase];
            } else if (digitGroup == 0 || (tensPlace == 0 && iterationCounter == 0)) {
                convertedNumber = HUNDREDS_IN_WORDS[digitGroup] + DECIMALS_IN_WORDS[tensPlace] + ONES_IN_WORDS_TO_USE[iterationCounter] + numbersInWordsPlural[grammaticalCase];
            } else {
                convertedNumber = HUNDREDS_IN_WORDS[digitGroup] + "и " + DECIMALS_IN_WORDS[tensPlace] + ONES_IN_WORDS_TO_USE[iterationCounter] + numbersInWordsPlural[grammaticalCase];
            }
        } else if (iterationCounter == 1) {
            convertedNumber = numbersInWords[grammaticalCase];
        }
    }
    return convertedNumber;
}


function inWords(inputNum) {
    if (inputNum === 0) {
        return "нула лева";
    }

    if (inputNum < 1 && inputNum > 0) {
        return "нула лева и " + Math.round(inputNum * 100) + " стотинки";
    }

    if (inputNum < 0) {
        negativeSign = "минус ";
        inputNumber = Math.abs(inputNum);
    } else {
        negativeSign = "";
        inputNumber = inputNum;
    }

    var millions, thousands, remainder, result, remainderDecimal, decimalRemainder;
    remainder = Math.floor(inputNumber);
    remainderDecimal = (inputNumber - remainder) * 100;
    isPlural = false;
    if (remainder < 1000000000) {
        if (remainder > 1) {
            millions = Math.floor(remainder / 1000000);
            remainder -= millions * 1000000;
            thousands = Math.floor(remainder / 1000);
            remainder -= thousands * 1000;
            result = convertNumberToWords(remainder, 0);
            if (thousands > 0) {
                if (!isPlural && remainder > 0) {
                    result = convertNumberToWords(thousands, 1) + "и " + result;
                } else {
                    result = convertNumberToWords(thousands, 1) + result;
                }
            }
            if (millions > 0) {
                if (!isPlural && (thousands > 0 || remainder > 0)) {
                    result = convertNumberToWords(millions, 2) + "и " + result;
                } else {
                    result = convertNumberToWords(millions, 2) + result;
                }
            }
        } else if (remainder === 1) {
            result = "един";
        }
    }
    if (remainderDecimal > 0) {
        decimalRemainder = remainderDecimal | 0;
        result = negativeSign + result + "лева и " + decimalRemainder + " стотинки";
    } else if (remainderDecimal === 0) {
        result = negativeSign + result + "лева";
    }

    return result;
}
/// <reference path="../App.js" />

(function () {
    "use strict";

    // 새 페이지가 로드될 때마다 초기화 함수를 실행해야 합니다.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            // Setup URL Data
            tryCatch(setupURL());

            // OnSingleClicked 이벤트 등록
            tryCatch(RegisterOnSingleClick);

            $('.target').each(function () {
                var elem = $(this);

                // Input sURL의 값을 임시 저장소에 저장
                elem.data('oldVal', elem.val());

                // 키입력이나 붙여넣기로 값이 변경 됐을때 이벤트 발생
                elem.bind("propertychange change click keyup input paste", function (event) {
                    // 값이 변경되었다면
                    if (elem.data('oldVal') != elem.val()) {
                        // 변경된 값을 임시 저장소에 저장
                        elem.data('oldVal', elem.val());

                        // 새로운 값이 입력되었을때 
                        var url = elem.val();
                        var videoId = parserURL(url);
                        $('input[name="videoid"]').val(videoId);
                    }
                });
            });
        });
    };

    // URL 파싱 - 유튜브 영상 키값(videoId) 알아내기
    function parserURL(url) {
        if (url == "") {
            return;
        }

        //var pattern = /v\=(\w+)/i;
        var pattern = /(\Wv=)[^&]*/
        var matches = pattern.exec(url);
        if (matches != null) {
            //var videoId = matches[0] //matches[1];
            var videoId = matches[0].substr(matches[0].lastIndexOf('=') + 1);

            return videoId;
        }
    }

    // 셀에서 URL값을 읽어서 input sURL에 넣기
    // onSingleClicked 이벤트 등록
    async function RegisterOnSingleClick() {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.onSingleClicked.add(cellSingleClickedHandler);

            await context.sync();
            console.log("Register a cellSingleClicked event handler for this worksheet.");
        });
    }
    // onSingleClicked 이벤트 
    // (TypeScript 에서는 Type인수를(Excel.EventArgs) 정의할 수 있지만 javaScript 에서는 지원하지 않음.)
    // Excel ScriptLAB에서는 지원되지만
    async function cellSingleClickedHandler(event) { //: Excel.WorksheetSingleClickedEventArgs) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Sheet1");

            let range = sheet.getRange(event.address);
            range.load("values");

            await context.sync();
            
            $('input[name="sURL"]').val(range.values[0][0]);
            $('input[name="sURL"]').click();

            console.log(`Click a Cell "${event.address}"`);
            console.log(range.values[0][0]);
        });
    }

    //Setup Youtube URL 
    async function setupURL() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Sheet1");

            const urlData = [["유튜브 URL 주소"],
                             ["https://www.youtube.com/watch?v=JpTqSzm4JOk"],
                             ["https://www.youtube.com/watch?v=nvJeJSrghOI"]];
            
            let range = sheet.getRange("A1:A3");
            range.values = urlData;
            range.format.autofitColumns();
        });
    }

    /* 함수를 호출하고 에러를 처리하는 기본 함수 */
    async function tryCatch(callback) {
        try
        {
            await callback();
        }
        catch (error)
        {
            // 참고: 프로덕션 추가 기능에서는 추가 기능UI를 통해 사용자에게 알림.
            console.log(OfficeHelpers.Utilities.host);
            console.log(OfficeHelpers.Utilities.platform);
            console.error(error);
        }
    }
})();
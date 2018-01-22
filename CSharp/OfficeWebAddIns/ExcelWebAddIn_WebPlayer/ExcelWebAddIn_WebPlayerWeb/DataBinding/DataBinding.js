(function () {
    'use strict';

    // 응용 프로그램의 각 페이지에 대해 Office.initialize 함수를 정의해야 합니다.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#bind-to-existing-data').click(bindToExistingData);

            if (dataInsertionSupported()) {
                $('#insert-sample-data').show();
                $('#insert-sample-data').click(insertSampleData);
            }
        });
    };

    // 기존 데이터에 시각화를 바인딩합니다.
    function bindToExistingData() {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Table,
            { id: app.bindingID, sampleData: visualization.generateSampleData() },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    window.location.href = '../Home/Home.html';
                } else {
                    app.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }

    // 선택한 데이터 설정을 현재 응용 프로그램에서 지원하는지 여부를 확인합니다.
    function dataInsertionSupported() {
        return Office.context.document.setSelectedDataAsync &&
            (Office.context.document.bindings &&
            Office.context.document.bindings.addFromSelectionAsync);
    }

    // 현재 선택 영역에 샘플 데이터를 삽입합니다(지원되는 경우).
    function insertSampleData() {
        Office.context.document.setSelectedDataAsync(visualization.generateSampleData(),
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: app.bindingID },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                window.location.href = '../Home/Home.html';
                            } else {
                                app.showNotification(result.error.name, result.error.message);
                            }
                        }
                    );
                } else {
                    app.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }
})();

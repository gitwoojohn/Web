var visualization = (function () {
    'use strict';

    var visualization = {};

    // 샘플 데이터를 사용하여 Office.TableData 개체를 생성하고 반환합니다.
    visualization.generateSampleData = function () {
        var sampleHeaders = [['이름', '성적']];
        var sampleRows = [
            ['김배식', 79],
            ['이민서', 95],
            ['이봉진', 86],
            ['양하늘', 93]];
        return new Office.TableData(sampleRows, sampleHeaders);
    }

    // 다음 매개 변수를 기준으로 시각화를 표시합니다.
    //        $element: 시각화가 표시될 jQuery 요소입니다.
    //        data: 데이터가 포함된 Office.TableData 개체입니다.
    //        errorHandler: 문자열 설명을 허용하는 오류 콜백입니다.
    visualization.display = function ($element, data, errorHandler) {
        if ((data.rows.length < 1) || (data.rows[0].length < 2)) {
            errorHandler('데이터 범위에 적어도 1개의 행과 2개의 열이 있어야 합니다.');
            return;
        }

        var maxBarWidthInPixels = 200;
        var $table = $('<table class="visualization" />');

        if (data.headers.length > 0) {
            var $headerRow = $('<tr />').appendTo($table);
            $('<th />').text(data.headers[0][0]).appendTo($headerRow);
            $('<th />').text(data.headers[0][1]).appendTo($headerRow);
        }

        for (var i = 0; i < data.rows.length; i++) {
            var $row = $('<tr />').appendTo($table);
            var $column1 = $('<td />').appendTo($row);
            var $column2 = $('<td />').appendTo($row);

            $column1.text(data.rows[i][0]);
            var value = data.rows[i][1];
            var width = (maxBarWidthInPixels * value / 100.0);
            var $visualizationBar = $('<div />').appendTo($column2);
            $visualizationBar.addClass('bar').width(width).text(value);
        }

        $element.html($table[0].outerHTML);
    };

    return visualization;
})();

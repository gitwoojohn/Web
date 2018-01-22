(function () {
    'use strict';

    // 응용 프로그램의 각 페이지에 대해 Office.initialize 함수를 정의해야 합니다.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

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
            //// Attach an event for when the user submits the form
            //$('form').on('submit', function (event) {

            //    // Prevent the page from reloading
            //    event.preventDefault();

            //    // Set the text-output span to the value of the first input
            //    var $input = $(this).find('input');
            //    var videoId = parserURL($input.val());

            //    rebuildElement();

            //    $('#videoPlayer').attr('src', '//www.youtube.com/embed/' + videoId + '?autoplay=1' + '&start=0' + '&end=200');
            //    //$('#text-output').text("You typed: " + input);
            //});

            // JQuery on Event
            // Attach an event for when the user submits the form
            //$('form').on('submit', function (event) {

            //    // Prevent the page from reloading
            //    event.preventDefault();

            //    // Set the text-output span to the value of the first input
            //    var $input = $(this).find('input');
            //    var videoId = parserURL($input.val());

            //    window.location.href = '../WebPlayer/Youtube.html?' + videoId;


            //    //$('#videoPlayer').attr('src', '//www.youtube.com/embed/' + videoId + '?autoplay=1' + '&start=0' + '&end=200');
            //    //$('#text-output').text("You typed: " + input);
            //});                    


            //displayDataOrRedirect();
        });
    }

    function rebuildElement() {
        $('#PlayerName').remove();  
        $('.padding').remove();
    }

    // URL 파싱 - 유튜브 영상 키값(videoId) 알아내기
    function parserURL(url) {
        if (url == "") {
            return;
        }

        //var pattern = /v\=(\w+)/i;
        var pattern =  /(\Wv=)[^&]*/
        var matches = pattern.exec(url);
        if (matches != null) {
            //var videoId = matches[0] //matches[1];
            var videoId = matches[0].substr(matches[0].lastIndexOf('=') + 1);
            return videoId;
        }        
    }

    // 바인딩이 있는지 확인하고 시각화를 표시합니다.
    //        또는 데이터 바인딩 페이지로 리디렉션합니다.
    function displayDataOrRedirect() {
        Office.context.document.bindings.getByIdAsync(
            app.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var binding = result.value;
                    displayDataForBinding(binding);
                    binding.addHandlerAsync(
                        Office.EventType.BindingDataChanged,
                        function () { displayDataForBinding(binding); }
                    );
                } else {
                    window.location.href = '../DataBinding/DataBinding.html';
                    //window.location.href = '../WebPlayer/Youtube.html';
                }
            });
    }

    // 해당 데이터에 대한 바인딩을 쿼리한 후 시각화 스크립트에 위임합니다.
    function displayDataForBinding(binding) {
        binding.getDataAsync(
            {
                coercionType: Office.CoercionType.Table,
                valueFormat: Office.ValueFormat.Unformatted,
                filterType: Office.FilterType.OnlyVisible
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    visualization.display($('#data-display'), result.value, showError);
                } else {
                    showError('데이터를 읽을 수 없습니다.');
                }
            }
        );

        function showError(message) {
            $('#data-display').html(
                '<div class="notice">' +
                '    <h3>오류</h3>' + $('<p/>', { text: message })[0].outerHTML +
                '    <a href="../DataBinding/DataBinding.html">' +
                '        <b>다른 데이터 범위에 바인딩할까요?</b>' +
                '    </a>' +
                '</div>');
        }
    }

})();

/// <reference path="../App.js" />

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            //To Do 1: Check my host version (host - Excel, Word, PowerPoint...)
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works new version ");
                return;
            };

            //displayVideoPanel();
        });
    };

    function getParameterByName(name, url) {
        if (!url) url = window.location.href;

        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"), results = regex.exec(url);

        if (!results) return null;
        if (!results[2]) return '';

        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }

    // URL 파싱 - 유튜브 영상 키값(videoId) 알아내기
    function parserURL(url) {
        if (url == "") {
            return;
        }

        //var pattern = /v\=(\w+)/i;
        var pattern = /(\Wv=)[^&]*/;
        var matches = pattern.exec(url);
        if (matches != null) {
            //var videoId = matches[0] //matches[1];
            var videoId = matches[0].substr(matches[0].lastIndexOf('=') + 1);
            return videoId;
        }
    }

    function displayVideoPanel() {
        var videoId = parserURL(decodeURIComponent(window.location.href));

        // 함든 함수로 나머지 인자 구하기
        var autoPlay = getParameterByName('autoplay');
        var startTime = getParameterByName('start');
        var endTime = getParameterByName('end');

        // 비디오 플레이 함수 호출
        displayVideo(videoId, autoPlay, startTime, endTime);
    }

    // 기존 IFrame 플레이어
    function displayVideo(videoId, autoPlay, startTime, endTime) {
        if (autoPlay == 0 && endTime == 0) {
            $('#videoPlayer').attr('src', '//www.youtube.com/embed/' + videoId + '?start=' + startTime);
        }
        else {
            $('#videoPlayer').attr('src', '//www.youtube.com/embed/' + videoId + '?autoplay=' + autoPlay + '&start=' + startTime + '&end=' + endTime);
        }
        //$('#videoLink').attr('href', '//www.youtube.com/watch?v=' + videoId);
    }
})();
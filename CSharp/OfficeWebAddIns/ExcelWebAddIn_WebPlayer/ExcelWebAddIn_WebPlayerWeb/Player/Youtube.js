/// <reference path="../App.js" />

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            displayVideoPanel();
        });
    };

    function getParameterByName(name, url) {
        if (!url) url = window.location.href;
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
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

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayVideoPanel() {
        //var item = Office.context.mailbox.item;
        
        //var entities = item.getEntities();
        //if (entities.urls == null || entities.urls.length == 0)
        //    return;

        //var pattern = /v\=(\w+)/i;

        //for (var i = 0; i < entities.urls.length; i++) {
        //    var url = entities.urls[i].toString();
        //    var matches = pattern.exec(url);
        //    if (matches != null) {
        //        var videoId = matches[1];
        //        $('#content-tabs').append('<div class="content-tab" data-videoId="' + videoId + '">Video ' + (i + 1) + '</div>');
        //    }

        //$('.content-tab').click(function () {
        //    //var videoId = $(this).data('videoid');
        //    //$('#content-tabs .selected').removeClass('selected');
        //    //$(this).addClass('selected');
        //    //displayVideo(videoId);
        //    //displayVideo();
        //});
        //$('.content-tab:first').click();

        
        // 넘긴 인자를 파싱해서 videoId로 받기
        //var videoId = location.href.substr(location.href.lastIndexOf('?') + 1);

        // Home.html 링크에서 Youtube.html로 넘어온 URL 파라메터 인자 구하기


        //var videoId = getParameterByName('videoid');

        // 자바스크립트가 제공하는 디코드URI 이용해서 인코딩된 URL을 다시 정상적인 주소로 변경
        var videoId = parserURL(decodeURIComponent(window.location.href));

        // 함든 함수로 나머지 인자 구하기
        var autoPlay = getParameterByName('autoplay'); 
        var startTime = getParameterByName('start');  
        var endTime = getParameterByName('end');  

        // 비디오 플레이 함수 호출
        displayVideo(videoId, autoPlay, startTime, endTime);
    }

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
/* 일반 응용 프로그램 기능 */

var app = (function () {
    'use strict';

    var app = {};

    app.bindingID = 'myBinding';

    // 각 페이지에서 호출할 일반 초기화 함수입니다.
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });

        // 초기화 후 일반 알림 함수를 노출합니다.
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();
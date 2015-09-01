(function () {
    'use strict';

    var angularApp = angular.module('SandBox');

    angularApp.service("scrollService", [
        function () {

            var increment = 25;
            var start = 1;
            var end = increment;
            var max = 200;
            var endOfList = false;

            return {
                init: function (callBack) {

                    $(window).scroll(function () {

                        var docElement = $(document)[0].documentElement;

                        var winElement = $(window)[0];

                        var heightCalc = docElement.scrollHeight - winElement.innerHeight;

                        var reachedBottom = heightCalc == winElement.pageYOffset;

                        if (reachedBottom && !endOfList) {

                            endOfList = end == max;

                            callBack();
                        }

                    });
                },
                reachedBottom: function () {

                    var docElement = $(document)[0].documentElement;

                    var winElement = $(window)[0];

                    var heightCalc = docElement.scrollHeight - winElement.innerHeight;

                    var reachedBottom = heightCalc == winElement.pageYOffset;

                    return reachedBottom;
                }
            }

        }]);
}());
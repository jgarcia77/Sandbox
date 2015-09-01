(function () {

    var angularApp = angular.module("SandBox");

    angularApp.controller("scrollController",
        ["$window", "$scope", "$http", "scrollService", scrollController]);
    

    function scrollController($window, $scope, $http, scrollService) {

        $scope.model = [];

        var increment = 25;
        var start = 1;
        var end = increment;
        var max = 200;

        var onResponse = function (response) {
            if (response.data) {
                $(response.data).each(function (index, item) {
                    $scope.model.push(item);
                })

                start = end + 1;
                end = end + increment;
            }
        }

        $scope.getNumbers = function () {
            $http.get("api/numbers/" + start + "/" + end).then(onResponse);
        }

        $scope.getNumbers();

        //$scope.scrollTop = 0;
        //$scope.windowHeight = 0;
        //$scope.documentHeight = 0;

        //var windowElement = $(window);
        //var documentElement = document;

        //function getDocHeight() {
        //    var height = Math.max(documentElement.body.scrollHeight, documentElement.documentElement.scrollHeight,
        //                          documentElement.body.offsetHeight, documentElement.documentElement.offsetHeight,
        //                          documentElement.body.clientHeight, documentElement.documentElement.clientHeight);

        //    return height;
        //}

        //var endOfList = false;


        scrollService.init($scope.getNumbers);


        //$(window).scroll(function () {

        //    var reachedBottom = scrollService.reachedBottom()

        //    if (reachedBottom && !endOfList) {
        //        endOfList = end == max;
        //        $scope.getNumbers();
        //    }

        //})
    }



}());
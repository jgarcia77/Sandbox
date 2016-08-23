(function () {

    var angularApp = angular.module("SandBox");

    angularApp.controller("indexController", ["$window", "$scope", "$http", indexController]);



    function indexController($window, $scope, $http) {
        
        $scope.model = [];

        var increment = 25;
        var start = 1;
        var end = increment;
        

        var onResponse = function (response) {
            if (response.data) {
                $(response.data).each(function (index, item) {
                    $scope.model.push(item);
                })

                start = end + 1;
                end = end + increment;
            }
        }

        var getNumbers = function () {
            $http.get("/api/numbers/" + start + "/" + end).then(onResponse);
        }
        
        //getNumbers();

        $('.infinite-scroll').jscroll({
            callback: getNumbers
        });
    }



}());
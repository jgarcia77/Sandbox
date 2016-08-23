(function () {

    var angularApp = angular.module("SandBox");

    angularApp.controller("jscrollController", ["$window", "$scope", "$http", "$upload", "settings", jscrollController]);



    function jscrollController($window, $scope, $http, $upload, settings) {

        $scope.model = [];

        var getNumbers = function (result) {
            $http.get("/api/getnumbers")
             .then(function (result) {

                 $(result.data).each(function (index, item) {
                     $scope.model.push(item);
                 })

             })
        }

        //$('.scroll').jscroll({
        //    loadingHtml: '<div>Loading...</div>',
        //    callback: getNumbers
        //});

        $http.get("/api/getnumbers")
             .then(getNumbers);        
        
        
        
    }



}());
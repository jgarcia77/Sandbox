(function () {

    var angularApp = angular.module("SandBox");

    angularApp.controller("indexController", ["$scope", indexController]);
        
    function indexController($scope) {

        $scope.Message = "Index Controller";

    }

    

}());
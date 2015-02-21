(function () {

    var angularApp = angular.module("SandBox");

    angularApp.controller("indexController", ["$scope", "settings", indexController]);
        
    function indexController($scope, settings) {

        $scope.Message = settings.webApiBaseUrl;

    }

    

}());
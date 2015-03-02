(function () {
    var angularApp = angular.module("SandBox", ["ngRoute", 'ngSanitize', 'ngAnimate', 'ngCookies', 'angularFileUpload']);

    angularApp.config(["$routeProvider", "$locationProvider", function ($routeProvider, $locationProvider) {
        $routeProvider
            .when("/", {
                templateUrl: "/Assets/Views/index.html",
                controller: "indexController"
            })
            .otherwise({
                redirectTo: "/"
            });
    }]);

}());
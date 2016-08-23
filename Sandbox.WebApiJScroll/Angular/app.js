(function () {
    var angularApp = angular.module("SandBox", ["ngRoute", 'ngSanitize', 'ngAnimate', 'ngCookies']);

    angularApp.config(["$routeProvider", "$locationProvider", function ($routeProvider, $locationProvider) {
        $routeProvider
            .when("/", {
                templateUrl: "/Angular/index.html",
                controller: "indexController"
            })
            .otherwise({
                redirectTo: "/"
            });
    }]);

}());
(function () {
    var angularApp = angular.module("SandBox", ["ngRoute", 'ngSanitize', 'ngAnimate', 'ngCookies', 'angularFileUpload']);

    angularApp.config(["$routeProvider", "$locationProvider", function ($routeProvider, $locationProvider) {
        $routeProvider
            .when("/index", {
                templateUrl: "/Assets/Views/index.html",
                controller: "indexController"
            })
            .when("/", {
                templateUrl: "/Assets/Views/scroll.html",
                controller: "scrollController"
            })
            .otherwise({
                redirectTo: "/"
            });
    }]);

}());
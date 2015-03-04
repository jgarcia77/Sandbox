(function () {

    var angularApp = angular.module("SandBox");

    angularApp.controller("indexController", ["$window", "$scope", "$http", "$upload", "settings", indexController]);
    
    

    function indexController($window, $scope, $http, $upload, settings) {

        $scope.scrollTop = 0;
        $scope.windowHeight = 0;
        $scope.documentHeight = 0;

        var windowElement = $(window);
        var documentElement = document;

        function getDocHeight() {
            var height = Math.max(documentElement.body.scrollHeight, documentElement.documentElement.scrollHeight,
                                  documentElement.body.offsetHeight, documentElement.documentElement.offsetHeight,
                                  documentElement.body.clientHeight, documentElement.documentElement.clientHeight);

            return height;
        }

        windowElement.on("scroll", function () {
            $scope.scrollTop = windowElement.scrollTop();
            $scope.windowHeight = windowElement.height();
            //$scope.documentHeight = documentElement.height();

            $scope.documentHeight = getDocHeight();

            var result = ($scope.scrollTop + $scope.windowHeight > $scope.documentHeight - 100)

            if (result) {
                console.log("near bottom");
            }

        })

        //$window.scroll(function () {
            
        //})

        $scope.Files = null;
        
        $scope.upload = [];
        $scope.fileUploadObj = { testString1: "Test string 1", testString2: "Test string 2" };

        $scope.onFileSelect = function ($files) {

            $scope.Files = $files;

            //$files: an array of files selected, each file has name, size, and type.
            for (var i = 0; i < $files.length; i++) {
                var $file = $files[i];
                (function (index) {
                    $scope.upload[index] = $upload.upload({
                        url: "upload/files", // webapi url
                        method: "POST",
                        data: { fileUploadObj: $scope.fileUploadObj },
                        file: $file
                    }).progress(function (evt) {
                        // get upload percentage
                        console.log('percent: ' + parseInt(100.0 * evt.loaded / evt.total));
                    }).success(function (data, status, headers, config) {
                        // file is uploaded successfully
                        console.log(data);
                    }).error(function (data, status, headers, config) {
                        // file failed to upload
                        console.log(data);
                    });
                })(i);
            }
        }

        $scope.abortUpload = function (index) {
            $scope.upload[index].abort();
        }


    }

    

}());
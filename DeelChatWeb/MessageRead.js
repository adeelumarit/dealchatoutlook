// <reference path="messageread.js" />
//var baseURL = "https://localhost:44371";
var baseURL = "https://accountingsystemapi20231101183220.azurewebsites.net"

var dialogueURL = "https://adeelumarit.github.io/dealchatoutlook/DeelChatWeb"


var app = angular.module('DealChat', ['ngMaterial', "ngRoute"], function () {


});


    
app.config(function ($routeProvider) {
    $routeProvider
        .when("/", {
            templateUrl: dialogueURL+"/Templates/mainPage.html"
        })
        .when("/Signup", {
            templateUrl: dialogueURL+"/Templates/Signup.html"
        })
        .when("/Login", {
            templateUrl: dialogueURL+"/Templates/Login.html"



        })
        .when("/blue", {
            templateUrl: "blue.htm"
        });
});

app.controller('Signupctrl', function ($scope, $mdDialog, $mdToast, $log, $location,) {


    $scope.create_User = function () {
        ProgressLinearActive();
        var userObject = {
            name: $scope.user_name,
            email: $scope.user_email,
            password: $scope.user_password
            // Other properties of the user object
        };

        var settings = {
            "url": baseURL+"/api/Home/createUser",
            "method": "POST",
            "timeout": 0,
            "headers": {
                "Content-Type": "application/json"
            },
            "data": JSON.stringify({
                name: $scope.user_name,
                email: $scope.user_email,
                password: $scope.user_password
            }),
        };

        $.ajax(settings).done(function (response) {
            console.log(response);
            loadToast("user created successfully")
            $location.path("/Login")
            ProgressLinearInActive();

        }).fail(function (error) {

            ProgressLinearInActive()
            loadToast("user creating error")

            console.log(error)
        });




    }

    $scope.GotoSingIn = function () {

        $location.path("/Login")
    }

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }


})


app.controller('loginctrl', function ($scope, $mdDialog, $mdToast, $log, $location,) {


    $scope.login = function (ev) {
        ProgressLinearActive();
        var userObject = {

            email: $scope.useremail,
            password: $scope.password
            // Other properties of the user object
        };

        console.log(userObject)

        var settings = {
            "url": baseURL+"/api/Home/Login",
            "method": "POST",
            "timeout": 0,
            "headers": {
                "Content-Type": "application/json"
            },
            "data": JSON.stringify({
                email: $scope.useremail,
                password: $scope.password
            }),
        };

        $.ajax(settings).done(function (response) {
            console.log(response);
            window.localStorage.setItem("userInfo", JSON.stringify(response))
            loadToast("Login Successful")
            window.location.reload();
            $location.path("/")
            ProgressLinearInActive()


        }).fail(function (error) {
            ProgressLinearInActive()
            if (error.status === 401) {
                loadToast(error.status + "  Unauthorized")


            }
            console.log(error)

        });


    }
    $scope.GotoSingup = function () {

        $location.path("/Signup")
    }


    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }


})
app.controller('mainpageCTRL', function ($scope, $mdDialog, $mdToast, $log, $timeout, $location) {
    Office.onReady(function () { 
        var clause_ID;
        let item = Office.context.mailbox.item;

       
        $scope.getSelectedText = function () {
            console.log($scope.contract)
            contractID = $scope.contract.id

            $.ajax({
                type: "get",
                url: baseURL+"/api/Home/getClause/" + contractID, // The URL of your controller action
                //data: JSON.stringify(newItem),
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    console.log(response)
                    if(response.length<=0){
                        loadToast("This Contract has no clauses")

                    }
                    $scope.Clause = response
                    ProgressLinearInActive()

                    // console.log($scope.Companies)
                },
                error: function (error) {
                    ProgressLinearInActive()
                    // Handle error, e.g., display error message
                    console.error("Error adding item:", error);
                }
            });
            item.subject.setAsync($scope.contract.contractName, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Subject updated successfully");
                } else {
                    console.error("Error updating subject: " + asyncResult.error.message);
                }
            });
        };
        $scope.getSelectedClauseINFO = function () {
            console.log($scope.Clauses)
            clause_ID = $scope.Clauses.id
        } 


        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var emailBodyHtml = result.value;
                console.log("Email body as HTML:", emailBodyHtml);

                // You can now use the emailBodyHtml variable as needed.
            } else {
                console.error("Error retrieving email body:", result.error.message);
            }
        });



    $scope.TagEmail_Clauses= function () {
        if (!$scope.$$phase) {
            $scope.$apply();
        }
        email_TagArray = [];

       item.body.getAsync(Office.CoercionType.Html, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var emailBodyHtml = result.value;
                item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
                    ProgressLinearActive();

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        var bodyText = asyncResult.value; // Declare and assign here
                        console.log("Email body: " + bodyText);
                        for (let i = 0; i < $scope.Clauses.length; i++) {

                            //var ClauseObject = {
                            //    "emailbody": bodyText,
                            //    "time": getDate(),
                            //    "Clauseid": $scope.Clauses[i].id,
                            //    //"Clauseid": clause_ID,
                            //};
                            //email_TagArray.push(ClauseObject);

                            // Send just the email_TagArray without wrapping it in another array
                            $.ajax({
                                type: "post",

                                url: baseURL+"/api/Home/tagClause", // The URL of your controller action
                                data: JSON.stringify(
                                    {
                                        //"id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
                                        "emailbody": bodyText,
                                        "HTMLContent": emailBodyHtml,
                                        "time": getDate(),
                                        "clauseid": $scope.Clauses[i].id,
                                    }),
                                contentType: "application/json; charset=utf-8",
                                //dataType: "json",
                                success: function (data) {
                                    // Handle success, e.g., update UI, display message, etc.
                                    console.log("Item added successfully:", data);
                                    ProgressLinearInActive()
                                    loadToast("tagged successfuly")
                                },
                                error: function (error) {
                                    // Handle error, e.g., display error message
                                    console.error("Error adding item:", error);
                                    loadToast("Clause tag error")

                                    $mdDialog.hide();


                                }


                            });
                        }

                        ////$.ajax({
                        ////    type: "PUT",

                        ////    url: "https://localhost:44371/api/Home/UpdateClause?Clauseid=" + clause_ID, // The URL of your controller action
                        ////    data: JSON.stringify(
                        ////        {
                        ////            //"id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
                        ////            "emailbody": bodyText,
                        ////            "time": getDate(),
                        ////        }),
                        ////    contentType: "application/json; charset=utf-8",
                        ////    //dataType: "json",
                        ////    success: function (data) {
                        ////        // Handle success, e.g., update UI, display message, etc.
                        ////        console.log("Item added successfully:", data);

                        ////        $mdDialog.hide();
                        ////        //getdefaultContracts()

                        ////        ProgressLinearInActive()
                        ////        loadToast("Clause added successfuly")
                        ////    },
                        ////    error: function (error) {
                        ////        // Handle error, e.g., display error message
                        ////        console.error("Error adding item:", error);
                        ////        loadToast("Error adding Clause")

                        ////        $mdDialog.hide();


                        ////    }
                        ////});


                        console.log(email_TagArray);
                    } else {
                        console.error("Error getting email body: " + asyncResult.error.message);
                    }
                });

                // You can now use the emailBodyHtml variable as needed.
            } else {
                console.error("Error retrieving email body:", result.error.message);
            }
        });
        //console.log(Email);
    

    };

    function getDate() {

        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const currentTime = new Date();

        const day = currentTime.getDate();
        const month = months[currentTime.getMonth()];
        const year = currentTime.getFullYear();

        const optionsCustom = { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true };
        const timePart = currentTime.toLocaleTimeString(undefined, optionsCustom);

        const formattedCustom = `${day}-${month}-${year} ${timePart}`;
        console.log("Custom Format:", formattedCustom);

        return formattedCustom
    }
    })

    $scope.reloadAddin = function () {
        window.location.reload();
    }
    

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }


})
app.controller('DealChatCTRL', function ($scope, $mdDialog, $mdToast, $log, $timeout,$location) {
    //<------------ globel variables -------------->
    var email_TagArray = [];
    var contractID = "";
    var Email = "";


    let userinfo = window.localStorage.getItem("userInfo")
    userinfo = JSON.parse(userinfo)


    var userid;
    if (userinfo) {
        userid = userinfo.id;
        console.log("userid :>" + userid)
        if (userinfo.id) {
            $location.path("/")
        }

    } else {
        $location.path("/Login")


    }
    ProgressLinearActive();
    getthecontract()
    function getthecontract() {


  

    $.ajax({
        type: "get",
        url: baseURL+"/api/Home/ContractsByUserId/" + userid, // The URL of your controller action
        //data: JSON.stringify(newItem),
        contentType: "application/json; charset=utf-8",
        dataType: "json", 
        success: function (response) {
            //let result = JSON.parse(response)
            $scope.Contracts = response
            ProgressLinearInActive();
            console.log($scope.Contracts)
        },
        error: function (error) {
            // Handle error, e.g., display error message
            ProgressLinearInActive();

            console.error("Error adding item:", error);
        }
    });
    }
 

   
    Office.onReady(function () {

       let item = Office.context.mailbox.item;

        $scope.getSelectedText = function () {
            console.log($scope.contract)
            contractID = $scope.contract.id
            item.subject.setAsync($scope.contract.contractName, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Subject updated successfully");
                } else {
                    console.error("Error updating subject: " + asyncResult.error.message);
                }
            });
        };


        //$scope.TagEmail_Contract = function () {
        //    ProgressLinearActive();
        //    Office.context.mailbox.item.to.getAsync(function (asyncResult) {
        //        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        //            var recipients = asyncResult.value;
        //            var emailPromises = [];

        //            recipients.forEach(function (recipient) {
        //                var email_TagArray = [];

        //                Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (bodyAsyncResult) {
        //                    if (bodyAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        //                        var bodyText = bodyAsyncResult.value;

        //                        var tagObject = {
        //                            Email: recipient.emailAddress,
        //                            EmailBody: bodyText,
        //                            Time: getDate()
        //                        };
        //                        email_TagArray.push(tagObject);

        //                        var emailPromise = new Promise(function (resolve, reject) {
        //                            $.ajax({
        //                                url: '/Home/AddEmailTags',
        //                                type: 'POST',
        //                                data: { itemId: itemId, emailTags: email_TagArray },
        //                                success: function (result) {
        //                                    if (result.success) {
        //                                        loadToast("email tagged");
        //                                        console.log(result.message);
        //                                        ProgressLinearInActive();
        //                                        resolve();
        //                                    } else {
        //                                        console.error(result.message);
        //                                        ProgressLinearInActive();
        //                                        reject();
        //                                    }
        //                                },
        //                                error: function () {
        //                                    console.error('An error occurred during the AJAX request.');
        //                                    loadToast("Error tagging email");
        //                                    ProgressLinearInActive();
        //                                    reject();
        //                                }
        //                            });
        //                        });

        //                        emailPromises.push(emailPromise);
        //                    } else {
        //                        console.error("Error getting email body: " + bodyAsyncResult.error.message);
        //                    }
        //                });
        //            });

        //            Promise.all(emailPromises)
        //                .then(function () {
        //                    // All emails have been tagged and sent successfully
        //                })
        //                .catch(function () {
        //                    // Error occurred while tagging or sending emails
        //                });
        //        } else {
        //            console.error("Error getting recipients: " + asyncResult.error.message);
        //        }
        //    });
        //};
        $scope.addCompany = function () {



            Office.context.ui.displayDialogAsync(dialogueURL+'/Templates/DealChatWebPage.html?userId=' + userid, { height: 80, width: 80 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            );
        }

        
    });
    $scope.Logout = function () {
        ProgressLinearActive();

        window.localStorage.clear("userInfo")
        $location.path("/Login")
        ProgressLinearInActive();
    }

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }
})

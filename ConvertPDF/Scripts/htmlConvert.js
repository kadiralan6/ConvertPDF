angular.module('HomeApp', []).controller('importPdfDocument', function ($scope, $http) {


    $scope.deleteIds = '';
    $scope.checkALLIds = '';
    $scope.frameUrl = "";
    $scope.lockUpdate = true;
    $scope.uploadlist = {};


    $scope.dowlandFile = function () {
        var file = "aren.xlsx";
        var url = "/files/download/" + file;
        window.open(window.location.origin + url, '_blank');
    };

    $scope.uploadExcelFile = function (type) {
        if(type===1)
            var fileUpload = document.getElementById('pdf-excel-selector');
        else
            var fileUpload = document.getElementById('pdf-excel-selector2');
        var files = fileUpload.files;

        if (files.length == 0) {
            iziToast.error({
                message: "YÜklenecek Satır bulunamadı.",
                position: 'topCenter',
            });
            return;
        }

        var f = files[0];
        {


            var reader = new FileReader();
            var name = f.name;
            reader.onload = function (e) {
                var data = e.target.result;
                var workbook = XLSX.read(data, {
                    type: 'binary',
                    raw: false,
                    dateNF: 'dd/MM/yyyy'
                });
                var jsonData;
                workbook.SheetNames.forEach(function (sheetName) {
                   
                    jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false });
                });

                jsonData.forEach(function (e) {

                    e.SerialNo = e.SIRANO;
                    e.DateTime = e.TARIH;
                    e.evrak = e.EVRAKNO;
                    e.CustomerName = e.ACIKLAMA;
                    e.text = e.STOKADI;
                    e.Quantity = e.MIKTAR;
                    e.PriceList = parseFloat(e.BIRIMFIYATTL);
                    e.TotalPrice = parseFloat(e.TUTARTL);
                    e.TCNumber = e.FATURACK;
                    e.Address = e.CARIADI;
                  
                  
                  
                    
                  

                    delete e.SIRANO;
                    delete e.TARIH;
                    delete e.EVRAKNO;
                    delete e.ACIKLAMA;
                    delete e.MIKTAR;
                    delete e.BIRIMFIYATTL;
                    delete e.TUTARTL;
                    delete e.FATURACK;
                    delete e.CARIADI;
                });
                $http({
                    method: "POST",
                    url: "/Home/UploadExcel",
                    headers: { "Content-Type": "Application/json;charset=utf-8" },
                    data: {
                        uploadList: jsonData

                    }

                }).then(function successCallback(response) {

                    $scope.uploadlist = response.data;

                });
            }
        }


        reader.readAsBinaryString(f);
    }
    $scope.clear = function () {
        $scope.uploadlist = {};
    }
    $scope.changePdf = function (type) {
        iziToast.success({
            message: "Bu İşlem Biraz Zaman Alacaktır. Lütfen Bekleyin. Tahmini Süre:" + ($scope.uploadlist.length+5)+" Saniye.",
            timeout: 1000 * ($scope.uploadlist.length+5),
            position: 'topCenter',
        });

        $http({
            method: "POST",
            url: "/Home/SavePDF",
            headers: { "Content-Type": "Application/json;charset=utf-8" },
            data: { pdfItems: $scope.uploadlist, type: type }

        }).then(function (response) {
         
            //$scope.test = response.data;
            window.open(response.data, "_blank");
            //$scope.frameUrl = response.data;//duzenleMetin($scope.test);
            //$('#mPdfShow').appendTo("body").modal('show');
        });
    };
    function duzenleMetin(metin) {
        // Başta ve sonda bulunan çift tırnakları kaldır
        return metin.replace(/^"|"$/g, '');
    }
    function spinnerLoad(cntrol) {
        var loader = document.getElementById('loader');
        if (cntrol == true)
            loader.style.display = "block";
        else
            loader.style.display = "none";
    }

    //function fireCustomLoading(cntrl) {
    //    if (cntrl === undefined || cntrl === null || cntrl === "")
    //        cntrl = false;

    //    if (cntrl) {
    //        if (!$('.custom-loading-wrapper').hasClass("active"))
    //            $('.custom-loading-wrapper').addClass("active");
    //    }
    //    else {
    //        if ($('.custom-loading-wrapper').hasClass("active"))
    //            $('.custom-loading-wrapper').removeClass("active");
    //    }
    //}


    $(document).ready(function () {

    });



});
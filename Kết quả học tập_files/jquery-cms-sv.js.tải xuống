function ajaxFormBeforeForGetDataForSubmit() { } ! function (e) {
           jQuery.extend({
            message: function (t, a, n) {
                a = a || "info", n = n || "Thông báo", toastr.clear();
                toastr.options = {
                    "closeButton": true,
                    "debug": false,
                    "positionClass": "toast-top-right",
                    "onclick": null,
                    "showDuration": "1000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                };
                var $toast = toastr[a](t, n);
                $toastlast = $toast;
            },
            info: function (t, a) {
                e.message(t, "info", a)
            },
            error: function (t, a) {
                e.message(t, "error", a)
            },
            warning: function (t, a) {
                e.message(t, "warning", a)
            },
            success: function (t, a) {
                e.message(t, "success", a)
            },
            confirm: function (t, fnAccept) {
                //e.message(t, "success", a)
                toastr.success("<div style='text-align:right;'><button type='button' class='btn grey-cascade btn-sm' style='margin-right:15px;'>Không</button><button type='button' id='confirmButtonYes' class='btn blue btn-sm'>Đồng ý</button></div><div class='clearfix'></div>", t + '<br/><br/>',
                    {
                        closeButton: false,
                        allowHtml: true,
                        timeOut: 0,
                        extendedTimeOut: 0,
                        positionClass: 'toast-top-center',
                        onShown: function (toast) {
                            $("#confirmButtonYes").click(function () {
                                $('.modal-backdrop').remove();
                                toastr.remove();
                                setTimeout(fnAccept, 0);
                            });

                        }
                    });
            },
        })
}(jQuery);

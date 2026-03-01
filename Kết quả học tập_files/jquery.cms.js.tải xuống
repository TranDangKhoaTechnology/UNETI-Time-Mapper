var gIsGridSave = false;
var gRecordIndex;
var gParsers = {
    
    "number": function (value) {       
        var kq = kendo.parseFloat(value);
        return kq;
    },

    "date": function (value, format) {
        format = format == undefined ? "yyyy/MM/dd hh:mm tt" : format;

        return kendo.parseDate(value, format);
    },
    "boolean": function (value) {
        if (typeof value === "string") {
            return value.toLowerCase() === "true";
        }
        return value != null ? !!value : value;
    },

    "string": function (value) {
        return value != null ? (value + "") : value;
    },

    "default": function (value) {
        
        return value;
    }
};
var gGridActions = {
    deleled: "deleted",
    deletedOne: "deletedOne",
    save: "save",
    visible: "visible",
    invisible: "invisible"
}
var gConts = {
    VAL_TOOLTIP: "data-val-tooltip",
    ATTR_CONTINUE_ADD: "continue-add",
    ATTR_CURRENT_SENDER_ID: "current-sender-id",
    ID_UPLOAD_SELECTOR_CONTEXT_MENU: "UploadSelectorContextMenu",
    ID_UPLOAD_SELECTOR_CONTEXT_MENU_POPUP: "UploadSelectorPopup",
    ID_UPLOAD_SELECTOR_CONTEXT_MENU_PASTE_POPUP: "UploadSelectorPastePopup",
    TXT_STANDARD_DATE_TIME_FORMAT: "g",
    TXT_STANDARD_DATE_FORMAT: "d",
    TXT_STANDARD_TIME_FORMAT: "HH:mm",
};
var gOptions = {
    NotyDismissQueue: true,
    NotyDisplayTimeout: 2000,    
    GridShowNoDataMsg: true,    
    PopupEditAutoClose: false,
    ValidatePopup: true,
    ValidatePopupType: 'error',
    ValidateTooltipOptions: null,
    AutoShowModelStateError: true
};
function ajaxFormBeforeForGetDataForSubmit() { } ! function (e) {
    jQuery.fn.extend({
        getAttr: function (t) {
            return e(this).filter("[" + t + "]").length > 0 ? e(this).attr(t) : ""
        },
        grid: function () {
            return e(this).data("kendoGrid")
        },
        gridPopupEditor: function (t, a, n, r) {
            toastr.clear();
            var i, u = e(this);
            0 == e("#" + t).length ? (i = e("<div id='" + t + "' ></div>").appendTo(document.body).kendoWindow({
                close: function () {
                    u.gridRefresh(), e("#" + t).closest(".k-window, .k-overlay").remove()
                },
                modal: !0,
                iframe: !1,
                draggable: !0,
                title: n,
                resizable: !0,
                width: 0,
                scrollable: !0,
                actions: ["Pin", "Minimize", "Maximize", "Close"],
                visible: !1
            }).data("kendoWindow"), e("#" + t).gridContent(t, a), i.center()) : i = e("#" + t).data("kendoWindow"), i.open()
        },
        popupEditor: function (t, a, n, closeFunction) {
            toastr.clear();
            var i;
            0 == e("#" + t).length ? (i = e("<div id='" + t + "' ></div>").appendTo(document.body).kendoWindow({
                close: function () {
                    if (typeof (closeFunction) == 'function') {
                        closeFunction();
                    };
                    e("#" + t).closest(".k-window, .k-overlay").remove()
                },
                modal: !0,
                iframe: !1,
                draggable: !0,
                title: n,
                resizable: !0,
                width: 0,
                scrollable: !0,
                actions: ["Pin", "Minimize", "Maximize", "Close"],
                visible: !1
            }).data("kendoWindow"), e("#" + t).gridContent(t, a), i.center()) : i = e("#" + t).data("kendoWindow"), i.open()
        },
        gridRefresh: function () {
            gRecordIndex = 0;
            try {
                e(this).data("kendoGrid").dataSource.read()
            } catch (e) { }
        },
        gridContent: function (t, a) {
            e.ajax({
                url: a,
                type: "GET",
                dataType: "html",
                async: !1,
                traditional: !0,
                success: function (a) {
                    var n = e("#" + t).data("kendoWindow");
                    n.content(a), n.center()
                }
            })
        },
        gridSelecteAll: function (t, a) {
            if (void 0 == a) a = e(t).data("field"), e(this).find("tbody [select='selected'][data-field=" + a + "]").each(function () {
                void 0 != e(this).attr("disabled") && "disabled" == e(this).attr("disabled") || (this.checked = t.checked)
            });
            else {
                var n = e(this).data("kendoGrid"),
                    r = t.checked;
                e(n.dataSource.data()).each(function () {
                    this.set(a, r)
                }), t.checked = r
            }
        },
        gridSelectedModels: function () {
            var grid = $(this).grid();
            var models = Array();

            $(this).gridSelected().each(function () {
                models.push(grid.dataItem($(this)));
            });

            return models;
        },
        gridSelectedChange: function (t) {
            var a = !0,
                n = e(t).data("field");
            a && e(this).find("tbody [select='selected'][data-field=" + n + "]").each(function () {
                e(this).is(":checked") || (a = !1)
            }), e(this).find("thead [select='selecteAll'][data-field=" + n + "]").each(function () {
                this.checked = a
            })
        },
        gridSelected: function (t) {

            return $(this).find("tbody tr").filter(function () {
                return $(this).find("[select='selected']:checked").length;
            });
        },
        gridSelectedCheckbox: function () {
            return $(this).find("tbody [select='selected']:checked");
        },
        gridCheckBoxCheckedChange: function (e) {
            var checked = $(e.sender).is(':checked');
            var grid = $(this).grid();
            var selectedAll = true;

            var dataItem = grid.dataItem($(e.sender).closest('tr'));

            dataItem.set(e.fieldName, checked);

            if (selectedAll) {
                $(grid.dataSource.data()).each(function () {
                    if (!this[e.fieldName]) {
                        selectedAll = false;
                    }
                });
            }

            $(grid.thead).find("[select='selecteAll'][data-field=" + e.fieldName + "]").each(function () {
                this.checked = selectedAll;
            });
        },
        gridDelete: function (t, a, n) {
            toastr.clear();
            e.ajax({
                url: a,
                data: {
                    ids: n
                },
                type: "GET",
                async: !1,
                traditional: !0,
                success: function (t) {
                    var a = "";
                    if (t.Errors || t.errors) {
                        var n = void 0 == t.Errors ? t.errors : t.Errors;
                        n && (e.each(n, function (t, n) {
                            "errors" in n && e.each(n.errors, function () {
                                a += this + "<br/>"
                            })
                        }), a = a.substr(0, a.lastIndexOf("<br/>")), e.error(a))
                    } else e.success("Xóa thành công")
                }
            })
        },
        gridDeleteOne: function (target, url, validateUrl, isReload = false, mesConfirm, mesSuccess) {
            if (mesConfirm == null || mesConfirm == '' || mesConfirm == undefined) {
                mesConfirm = 'Bạn có chắc chắn muốn xóa không?';
            }

            if (mesSuccess == null || mesSuccess == '' || mesSuccess == undefined) {
                mesSuccess = "Xóa thành công";
            }

            toastr.clear();
            toastr.remove();
            $('<div class="modal-backdrop fade in"></div>').appendTo(document.body);

            var grid = this;

            toastr.options.onclick = function () {
                $('.modal-backdrop').remove();
                toastr.remove();
            };
            toastr.warning("<div style='text-align:right;'><button type='button' class='btn grey-cascade btn-sm' style='margin-right:15px;'>Không</button><button type='button' id='confirmDeleteYes' class='btn blue btn-sm'>Đồng ý</button><div><div class='clearfix'></div>", mesConfirm,
                {
                    closeButton: false,
                    allowHtml: true,
                    timeOut: 0,
                    extendedTimeOut: 0,
                    positionClass: 'toast-center toast-top-center',
                    onShown: function (toast) {
                        $("#confirmDeleteYes").click(function () {
                            $('.modal-backdrop').remove();
                            toastr.remove();
                            $.ajaxPostModels(url, null, function (e) {
                                var successMsg = mesSuccess; //$(target).getAttr("data-success-msg");
                                if (successMsg.length > 0) {
                                    $.success(successMsg);
                                    if (isReload) {
                                        location.reload();
                                    }
                                }
                                grid.gridOnAction({ action: gGridActions.deletedOne, data: url });
                            });
                            $(grid).gridRefresh(target);
                        });
                    }
                });
        },
        gridDeleteList: function (target, url, mesConfirm, mesSuccess, mesIfNoCheck) {
            if (mesConfirm == null || mesConfirm == '' || mesConfirm == undefined) {
                mesConfirm = 'Bạn có chắc chắn muốn xóa không?';
            }

            if (mesSuccess == null || mesSuccess == '' || mesSuccess == undefined) {
                mesSuccess = "Xóa thành công";
            }

            if (mesIfNoCheck == null || mesIfNoCheck == '' || mesIfNoCheck == undefined) {
                mesIfNoCheck = "Vui lòng chọn dữ liệu";
            }

            toastr.clear();
            toastr.remove();
            var arr = [];
            e(this).find("tbody [select='selected']:checked").each(function () {
                arr.push(e(this).val())
            });
            if (arr.length > 0) {

                var grid = $(this);

                $('<div class="modal-backdrop fade in"></div>').appendTo(document.body);
                toastr.options.onclick = function () {
                    $('.modal-backdrop').remove();
                    toastr.remove();
                };
                toastr.warning("<div style='text-align:right;'><button type='button' class='btn grey-cascade btn-sm' style='margin-right:15px;'>Không</button><button type='button' id='confirmDeleteYes' class='btn blue btn-sm'>Đồng ý</button><div><div class='clearfix'></div>", mesConfirm,
                    {
                        closeButton: false,
                        allowHtml: true,
                        timeOut: 0,
                        extendedTimeOut: 0,
                        positionClass: 'toast-center toast-top-center',
                        onShown: function (toast) {
                            $("#confirmDeleteYes").click(function () {
                                $('.modal-backdrop').remove();
                                toastr.remove();
                                $.ajaxPostModels(url, arr, function (e) {
                                    var successMsg = mesSuccess; //$(target).getAttr("data-success-msg");
                                    if (successMsg.length > 0) {
                                        $.success(successMsg);
                                    }
                                    grid.gridOnAction({ action: gGridActions.deletedOne, data: url });
                                });
                                $(grid).gridRefresh(target);
                            });
                        }
                    });
            } else {
                e.warning(mesIfNoCheck);
            }

        },
        gridChangeList: function (target, url, mesConfirm, mesSuccess) {
            toastr.clear();
            toastr.remove();
            var arr = [];
            e(this).find("tbody [select='selected']:checked").each(function () {
                arr.push(e(this).val())
            });
            if (arr.length > 0) {

                var grid = $(this);

                $('<div class="modal-backdrop fade in"></div>').appendTo(document.body);
                toastr.options.onclick = function () {
                    $('.modal-backdrop').remove();
                    toastr.remove();
                };
                toastr.warning("<div style='text-align:right;'><button type='button' class='btn grey-cascade btn-sm' style='margin-right:15px;'>Không</button><button type='button' id='confirmDeleteYes' class='btn blue btn-sm'>Đồng ý</button><div><div class='clearfix'></div>", mesConfirm,
                    {
                        closeButton: false,
                        allowHtml: true,
                        timeOut: 0,
                        extendedTimeOut: 0,
                        positionClass: 'toast-center toast-top-center',
                        onShown: function (toast) {
                            $("#confirmDeleteYes").click(function () {
                                $('.modal-backdrop').remove();
                                toastr.remove();
                                $.ajaxPostModels(url, arr, function (e) {
                                    var successMsg = mesSuccess; 
                                    if (successMsg.length > 0) {
                                        $.success(successMsg);
                                    }
                                    grid.gridOnAction({ action: gGridActions.deletedOne, data: url });
                                });
                                $(grid).gridRefresh(target);
                            });
                        }
                    });
            } else {
                e.warning("Vui lòng chọn dữ liệu");
            }

        },
        gridChangeBitList: function (target, url, title) {
            toastr.clear();
            toastr.remove();
            var arr = [];
            e(this).find("tbody [select='selected']:checked").each(function () {
                arr.push(e(this).val())
            });
            if (arr.length > 0) {

                var grid = $(this);

                $('<div class="modal-backdrop fade in"></div>').appendTo(document.body);
                toastr.options.onclick = function () {
                    $('.modal-backdrop').remove();
                    toastr.remove();
                };
                toastr.warning("<div style='text-align:right;'><button type='button' id='confirmDeleteYes' class='btn blue btn-sm'>Đồng ý</button><div><div class='clearfix'></div>", 'Bạn có chắc chắn muốn thay đổi dữ liệu?',
                    {
                        closeButton: false,
                        allowHtml: true,
                        timeOut: 0,
                        extendedTimeOut: 0,
                        positionClass: 'toast-center toast-top-center',
                        onShown: function (toast) {
                            $("#confirmDeleteYes").click(function () {
                                $('.modal-backdrop').remove();
                                toastr.remove();
                                $.ajaxPostModels(url, arr, function (e) {
                                    var successMsg = "Cập nhật thành công"; //$(target).getAttr("data-success-msg");
                                    if (successMsg.length > 0) {
                                        $.success(successMsg);
                                    }
                                    //grid.gridOnAction({ action: gGridActions.deletedOne, data: url });
                                });
                                $(grid).gridRefresh(target);
                            });
                        }
                    });
            } else {
                e.warning("Vui lòng chọn dữ liệu");
            }

        },
        gridChangeBit: function (t, a) {
            var n = this;
            toastr.clear();
            //$.message("Bạn muốn thay đổi dữ liệu đang chọn ?")
            setTimeout(function () {
                e.confirm(e(t).getAttr("data-confirm"), function () {
                    e.ajax({
                        url: a,
                        type: "POST",
                        async: !1,
                        traditional: !0,
                        success: function (a) {
                            a ? e(n).gridRefresh(t) : e.warning("Không có dữ liệu thay đổi")
                        }
                    })
                })
            })
        },
        gridChange: function (t, a) {
            var n = this;
            toastr.clear();
            setTimeout(function () {
                e.confirm(e(t).getAttr("data-confirm-msg"), function () {
                    e.ajax({
                        url: a,
                        type: "POST",
                        async: !1,
                        traditional: !0,
                        success: function (a) {
                            a ? e(n).gridRefresh(t) : e.warning("Không có dữ liệu thay đổi")
                        }
                    })
                })
            })
        },
        gridDataSource: function () {
            return $(this).data('kendoGrid').dataSource;
        },
        gridModelItems: function () {
            /// <summary>
            /// Lấy các model item trong datasource
            /// </summary>
            var dataSource = $(this).gridDataSource();

            // Có group
            if (dataSource.group().length) {
                var items = Array();

                $(dataSource.data()).each(function () {
                    $(this.items).each(function () {
                        items.push(this);
                    });
                });

                return items;
            }

            return dataSource.data();
        },
        gridDirtyItems: function () {
            var dirtyItems = Array();

            var data = $(this).gridModelItems();

            // Lấy các item có thay đổi
            for (var i = 0; i < data.length; i++) {
                if (data[i].dirty) {
                    dirtyItems.push(data[i]);
                }
            }

            return dirtyItems;
        },
        gridIsDirty: function () {
            var data = $(this).gridModelItems();

            // Lấy các item có thay đổi
            for (var i = 0; i < data.length; i++) {
                if (data[i].dirty) {
                    return true;
                }
            }

            return false;
        },
        gridBit: function (t, a) {

            toastr.clear();
            var n = this;
            e.ajax({
                url: a,
                //data: {
                //    id: a
                //},
                type: "POST",
                async: !1,
                traditional: !0,
                success: function (e) {
                    e && n.gridRefresh()
                }
            })
        },

        gridBitMes: function (t, a) {
            toastr.clear();
            var n = this;
            e.ajax({
                url: a,
                //data: {
                //    id: a
                //},
                type: "POST",
                async: !1,
                traditional: !0,
                success: function (e) {
                    var t = "";
                    if (e.Errors || e.errors) {
                        var a = void 0 == e.Errors ? e.errors : e.Errors;
                        a && ($.each(a, function (data, a) {
                            "errors" in a && $.each(a.errors, function () {
                                t += this + "<br/>"
                            })
                        }), t = t.substr(0, t.lastIndexOf("<br/>")))
                    }

                    if (t != null && t != "") {
                        $.warning(t);
                    } else {
                        n.gridRefresh();
                        $.success("Cập nhật thành công");
                    }
                }
            })
        },


        gridOnAction: function (e) {
            var grid = $(this).grid();
            for (var i in grid._events["onaction"]) {
                setTimeout(function () { grid._events["onaction"][i](e); }, 0);
            }
        },
        gridSave: function () {
            // Update các cell edit
            $(this).gridInCellClientEditCommit();

            // Có thay đổi
            if ($(this).gridIsDirty()) {
                gIsGridSave = true;

                $(this).grid().dataSource.grid = $(this).grid();

                $(this).grid().saveChanges();
            }
            else {
                $.warning(gResources.msgNoChange);
            }
        },
        formIsValid: function (e, t) {
            return !1
        },
        formSubmit: function (t, a) {
            //try {
            //    $("body,html").animate({ scrollTop: 0 }, 400);
            //} catch (ex) {

            //}
            e(this).formIsValid(t, a);
            this.submit()
        },
        formFilter: function () {
            var t = e(this).parents("form");
            if (t.length > 0) {
                var a = t.data("grid-id"),
                    n = t.data("search-suffix");
                a && n && e("#" + a).gridFilterByForm(t, n)
            }
        },
        gridFilterByForm: function (t, a) {
            e(this).grid().dataSource.filter(e(this).gridGetFilters(t, a))
        },
        gridGetFilters: function (t, a) {
            var n = [],
                r = {},
                i = !1,
                u = e(this).grid();
            if (e.each(t.serializeArray(), function () {
                if (this.value.length > 0) {
                    var t = this.name;
                    if (void 0 != a && a.length > 0) {
                        if (!t.endsWith(a)) return;
                        t = this.name.replace(a, "")
                    }
                    t.endsWith("From") ? (t = t.replace("From", ""), r[t] = e.extend({}, {
                        from: this.value
                    }, r[t])) : t.endsWith("To") ? (t = t.replace("To", ""), r[t] = e.extend({}, {
                        to: this.value
                    }, r[t])) : r[t] = e.extend({}, {
                        value: this.value
                    }, r[t]), r[t] = e.extend({}, {
                        elementName: this.name
                    }, r[t]), i = !0
                }
            }), i) {
                var o = u.dataSource.options.schema.model,
                    d = function (t, a, i, u) {
                        if (void 0 != t || r[a].elementName.endsWith("_input") || (t = e("[name=" + r[a].elementName + "]").getField(i)), void 0 != t && void 0 != t.parse && ("string" == t.type || "number" == t.type || "date" == t.type || "boolean" == t.type)) {
                            void 0 == u && (u = "string" == t.type ? "contains" : "eq");
                            var o = t.parse(i);
                            null != o && ("string" == t.type && (o = jQuery.trim(o)), n.push({
                                field: a,
                                operator: u,
                                value: o
                            }))
                        }
                    };
                e.each(r, function (t, a) {
                    var r = u.columns[t],
                        i = o.fields[t];
                    if (void 0 != a && t != o.id)
                        if (void 0 != a.from || void 0 != a.to) void 0 != a.from && d(i, t, a.from, "gte"), void 0 != a.to && d(i, t, a.to, "lte");
                        else if (void 0 != r && void 0 != r.values) {
                            var l = e.grep(r.values, function (e) {
                                return -1 != e.text.toLocaleLowerCase().indexOf(a.value)
                            });
                            e.each(l, function (e, a) {
                                n.push({
                                    field: t,
                                    operator: "eq",
                                    value: a.value
                                })
                            })
                        } else d(i, t, a.value)
                })
            }
            return n
        },
        gridEventDataBinding: function (t) {
            var a = null;
            (a = this[0] instanceof kendo.ui.Grid ? this[0] : e(this).grid()) && (a.thead.find("[select='selecteAll']").each(function () {
                e(this).attr("checked", !1)
            }), gRecordIndex = (a.dataSource.page() - 1) * a.dataSource.pageSize(), gRecordIndex = isNaN(gRecordIndex) ? 0 : gRecordIndex)
        },
        comboBoxRefreshDataSource: function () {
            var cbo = $(this).getKendoComboBox();
            if (cbo != undefined) {
                cbo.dataSource.read();
                // cbo.refresh();
                setTimeout(function () {
                    cbo.value("");
                }, 20);
            }
        },
        childFormSubmit: function (t) {
            e(this).append(' <i class="fa fa-spinner fa-spin fa-btn-sumbit"></i>');
            var a = e(this).parents("form");
            try {
                for (instance in CKEDITOR.instances) CKEDITOR.instances[instance].updateElement()
            } catch (e) { }
            a.length > 0 && (t = t || {}, e(a).formSubmit(t))
        },
        reset: function () {
            e(this).find(":input").each(function () {
                switch (this.type) {
                    case "password":
                    case "select-multiple":
                    case "select-one":
                    case "text":
                    case "textarea":
                        e(this).val("");
                        break;
                    case "checkbox":
                    case "radio":
                        this.checked = !1
                }
            })
        },
        formReset: function () {
            e(this).reset(), e(this).find("input[data-role=dropdownlist]").each(function () {
                e("#" + e(this).attr("id")).data("kendoDropDownList").select(0)
            }), e(this).find("input[data-role=combobox]").each(function () {
                e("#" + e(this).attr("id")).data("kendoComboBox").value("")
            }), e(this).find("select[data-role=multiselect]").each(function () {
                e("#" + e(this).attr("id")).getKendoMultiSelect().value([])
            }), e(this).setFocus()
        },
        setFocus: function () {
            var t = this;
            setTimeout(function () {
                var a = e(t).find("[first-focus=1]");
                (0 == a.length || 1 == a[0].disabled || a.attr("readonly").length > 0) && (a = e(t).find(":input:not(:button,[type=submit],[type=reset],[disabled],[readonly]):visible:first")), a.focus().val(a.val())
            }, 500)
        },
        kendoWindowClose: function () {
            var t;
            void 0 != (t = e(this).filter("[data-role=window]").length > 0 ? e(this).attr("id") : e(this).parents("[data-role=window]").attr("id")) && e("#" + t).data("kendoWindow").close()
        },
        kendoWindowRefresh: function (url, title, options, autoFocus, autoOpen) {
            var id = $(this).attr("id");
            if (id != undefined) {
                // Get dữ liệu
                $.ajax({
                    type: "GET",
                    url: url,
                    async: false,
                    traditional: true,
                    contentType: 'application/json; charset=utf-8',
                    success: function (e) {
                        // Không có lỗi
                        if (!ajaxErrorHandler(e)) {
                            var that = $("#" + id);
                            var popup = that.data('kendoWindow');

                            // Có options
                            if (options != undefined &&
                                options != null) {
                                popup.setOptions(options);
                            }

                            popup.content(e);

                            // Lấy các thông tin thiết lập trong content
                            var width = that.find('[window-width]').length > 0 ?
                                that.find('[window-width]').attr('window-width') :
                                that.find('[data-window-width]').attr('data-window-width');

                            var height = that.find('[window-height]').length > 0 ?
                                that.find('[window-height]').attr('window-height') :
                                that.find('[data-window-height]').attr('data-window-height');

                            var windowTitle = that.find('[window-title]').length > 0 ?
                                that.find('[window-title]').attr('window-title') :
                                that.find('[data-window-title]').attr('data-window-title');

                            var overflow = that.find('[data-window-overflow]').attr('data-window-overflow');

                            var fnClose = popup.options.close;

                            popup.setOptions({
                                close: function () {

                                    // Có hàm close
                                    if (fnClose != undefined) {
                                        // Gán vào this và gọi để ở method được gọi có thể truy vấn đối tượng this
                                        this.fnClose = fnClose;
                                        this.fnClose();
                                    }

                                    that.parent().first().remove();
                                },
                                width: width,
                                height: height
                            });

                            // Ưu tiên windowTitle
                            title = windowTitle != undefined ? windowTitle : title;

                            popup.title(title || '');
                            popup.center();

                            autoOpen = autoOpen == undefined ? true : autoOpen;

                            // Overflow
                            if (overflow != undefined) {
                                that.css("overflow", overflow);
                            }

                            if (autoOpen) {
                                popup.open();
                            }

                            // Mặc định là auto focus
                            if (autoFocus == undefined ||
                                autoFocus == true) {
                                that.setFocus();
                            }
                        }
                    },
                    error: function (e) {
                        showDefaultErrorMessage();
                    }
                });
            }
        },
        kendoWindowRefreshPost: function (url, title, data, options, autoFocus, autoOpen) {
            var id = $(this).attr("id");
            if (id != undefined) {
                // Get dữ liệu
                $.ajax({
                    type: "POST",
                    url: url,
                    data: data,
                    async: false,
                    traditional: true,
                    success: function (e) {
                        // Không có lỗi
                        if (!ajaxErrorHandler(e)) {
                            var that = $("#" + id);
                            var popup = that.data('kendoWindow');

                            // Có options
                            if (options != undefined &&
                                options != null) {
                                popup.setOptions(options);
                            }

                            popup.content(e);

                            // Lấy các thông tin thiết lập trong content
                            var width = that.find('[window-width]').length > 0 ?
                                that.find('[window-width]').attr('window-width') :
                                that.find('[data-window-width]').attr('data-window-width');

                            var height = that.find('[window-height]').length > 0 ?
                                that.find('[window-height]').attr('window-height') :
                                that.find('[data-window-height]').attr('data-window-height');

                            var windowTitle = that.find('[window-title]').length > 0 ?
                                that.find('[window-title]').attr('window-title') :
                                that.find('[data-window-title]').attr('data-window-title');

                            var overflow = that.find('[data-window-overflow]').attr('data-window-overflow');

                            var fnClose = popup.options.close;

                            popup.setOptions({
                                close: function () {

                                    // Có hàm close
                                    if (fnClose != undefined) {
                                        // Gán vào this và gọi để ở method được gọi có thể truy vấn đối tượng this
                                        this.fnClose = fnClose;
                                        this.fnClose();
                                    }

                                    that.parent().first().remove();
                                },
                                width: width,
                                height: height
                            });

                            // Ưu tiên windowTitle
                            title = windowTitle != undefined ? windowTitle : title;

                            popup.title(title || '');
                            popup.center();
                            autoOpen = autoOpen == undefined ? true : autoOpen;

                            // Overflow
                            if (overflow != undefined) {
                                that.css("overflow", overflow);
                            }

                            if (autoOpen) {
                                popup.open();
                            }

                            // Mặc định là auto focus
                            if (autoFocus == undefined ||
                                autoFocus == true) {
                                that.setFocus();
                            }
                        }
                    },
                    error: function (e) {
                        showDefaultErrorMessage();
                    }
                });
            }
        },
        kendoWindowAutoTopHeight: function () {

            var popupID = $(this).kendoWindowGetId();

            if (popupID != undefined) {

                var popup = $("#" + popupID).data("kendoWindow");

                popup.setOptions({
                    title: popup.title(),
                    maxHeight: $(window).height() - 70
                });

                // Canh vị trí
                popup.center();
            }
        },
        gridInCellClientEditCommit: function () {
            /// <summary>
            /// Commit tất cả các incell edit trên lưới
            /// </summary>
            if ($(this).data("incell-edit")) {
                var grid = $(this).grid();
                $("[data-cell-edit]", this).each(function () {
                    var that = $(this);

                    if (that.valueOfControl() !== undefined) {
                        var fieldName = that.data("cell-edit");
                        var field = that.getField();
                       
                        // Lấy dữ liệu hiện tại và dữ liệu trước
                        var value = that.valueOfControl();
                        var oldValue = that.data("old-value") ? field.parse(that.data("old-value")) : null;

                        if (oldValue == undefined || oldValue == null) {
                            oldValue = that.attr("data-old-value");
                        }


                        // Lấy data item
                        var dataItem = grid.dataSource.getByUid(that.closest('tr').data("uid"));

                        // Có dữ liệu thay đổi
                        if (fieldName != undefined && (value != oldValue)) {

                            var firstSym = fieldName.indexOf("[");

                            // fieldName là mảng
                            if (firstSym != -1) {
                                var fieldNameReal = fieldName.substr(0, firstSym);
                                var key = fieldName.substr(firstSym + 1, fieldName.length - fieldName.indexOf("]"));

                                dataItem[fieldNameReal][key] = value;
                            }
                            else {
                                dataItem[fieldName] = value;
                            }

                            dataItem.dirty = true;
                        }
                    }
                });

                // Làm mới dữ liệu
                grid.refresh();
            }
        },
        browserPaste: function () {
            $(this).kendoWindowClose();
            var val = $(this).val();
            var targetID = $(this).attr("target-id");

            if (val.length > 0 &&
                targetID.length > 0) {
                var kendoMultiselect = $("#" + targetID).getKendoMultiSelect();

                var data = kendoMultiselect.dataSource;
                var selected = [];

                $.each(val.split('\n'), function () {
                    if (this.length > 0) {
                        data.add({ name: this.substring(this.lastIndexOf('/') + 1), url: this });
                        selected.push(this);
                    }
                });

                kendoMultiselect.value(selected);
            }
        },
        /*control*/
        disableButton: function () {
            $(this).addClass("k-state-disabled");
            $(this).attr("onclick_old", $(this).attr("onclick"));
            $(this).removeAttr("onclick");
        },
        enableButton: function () {
            $(this).removeClass("k-state-disabled");
            $(this).attr("onclick", $(this).attr("onclick_old"));
        },
        setFocus: function () {
            /// <summary>
            /// Set focus cho control trong container. 
            /// Ưu tiên các input có attribute [first-focus=1], trong trường hợp không tìm thấy sẽ lấy input đầu tiên
            /// </summary>
            var container = this;

            var fn = function () {
                // Tìm phần tử có tag
                var obj = $(container).find("[first-focus=1]");

                // Không có dữ liệu hoặc bị disable, readonly
                if (obj.length == 0 ||
                    obj[0].disabled == true ||
                    obj.attr("readonly").length > 0) {

                    obj = $(container).find(":input:not(:button,[type=submit],[type=reset],[disabled],[readonly]):visible:first");
                }

                obj.focus().val(obj.val());
            };

            setTimeout(fn, 500);
        },
        getField: function (value) {
            /// <summary>
            /// Lấy thông tin field: loại, parse function
            /// </summary>
            /// <param name="value"></param>
            /// <returns type=""></returns>
            var field = null;
            var element = $(this);
            
            // Lấy value
            if (!value) {
                value = element.val();
            }

            if (element.data("role") != undefined) {
                var role = element.data("role");

                if (role == "multiselect" ||
                    role == "dropdownlist" ||
                    role == "combobox" ||
                    role == "numerictextbox") {
                    // Boolean
                    if (value == "true" || value == "false") {
                        field = {
                            type: "boolean", parse: gParsers["boolean"]
                        };
                    }
                    else {
                        field = {
                            type: "number", parse: gParsers["number"]
                        };
                    }
                } else if (role == "diemso" ) {
                    field = {
                        type: "number", parse: gParsers["number"]
                    };
                }
                else if (role == "datepicker" ||
                    role == "datetimepicker" ||
                    role == "timepicker") {

                    var format = "yyyy/MM/dd hh:mm tt";
                    var value = null;

                    if (role == "datepicker") {
                        format = element.getKendoDatePicker().options.parseFormats;
                        value = element.getKendoDatePicker().value();
                    }
                    else if (role == "datetimepicker") {
                        format = element.getKendoDateTimePicker().options.parseFormats;
                        value = element.getKendoDateTimePicker().value();
                    } else {
                        format = element.getKendoTimePicker().options.parseFormats;
                        value = element.getKendoTimePicker().value();
                    }

                    if (format.length > 1) {
                        format = format[format.length - 1];
                    }

                    field = {
                        type: "date", parse: gParsers["date"], format: format, value: value
                    };
                }
            }
            else if (element.attr("type") == "radio") {
                field = {
                    type: "boolean", parse: gParsers["boolean"]
                };
            }
            else if (element.attr("type") == "text") {
                field = {
                    type: "string", parse: gParsers["string"]
                };
            }

            return field;
        },
        valueOfControl: function () {
            /// <summary>
            /// Lấy giá trị của control tự động dựa theo loại control
            /// </summary>
            /// <returns type=""></returns>
            var value = $(this).val();

            var field = $(this).getField(value);

            if (field) {
                if (field.type == 'date') {
                    return field.parse(field.value, field.format);
                }

                return field.parse(value);
            }

            return value;
        },
        setView: function (url, fnSuccess, fnFailed) {
            /// <summary>
            /// Gán nội dung lấy từ server bằng ajax vào một container
            /// </summary>
            /// <param name="url"></param>
            /// <param name="fnSuccess"></param>
            /// <param name="fnFailed"></param>
            var container = $(this);

            $.ajax({
                type: "GET",
                url: url,
                async: false,
                traditional: true,
                beforeSend: function () {
                    kendo.ui.progress(container, true);
                },
                complete: function () {
                    kendo.ui.progress(container, false);
                },
                success: function (e) {
                    // Không có lỗi
                    if (!ajaxErrorHandler(e)) {
                        $(container).html(e);

                        var title = $(container).find('[window-title]').attr('window-title');

                        if (title != undefined) {
                            $('title').html(title);
                        }

                        if (fnSuccess != undefined) {
                            // Là function name
                            if (typeof fnSuccess == "function") {
                                fnSuccess(e);
                            }
                            else if (typeof getFunction(fnSuccess) == "function") {
                                getFunction(fnSuccess, ["e"]).apply(e);
                            }
                        }
                    }
                },
                error: function (e) {
                    showDefaultErrorMessage();
                    kendo.ui.progress(container, false);

                    if (fnFailed != undefined) {
                        // Là function name
                        if (typeof fnFailed == "function") {
                            fnFailed(e);
                        }
                        else if (typeof getFunction(fnFailed) == "function") {
                            getFunction(fnFailed, ["e"]).apply(e);
                        }
                    }
                }
            });
        },
    }),
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
        info: function (t, a, l) {
            toastr.clear();
            try {
                if (l == true) {
                    e.message(t, "info", a)
                } else {
                    window.parent.info_parrent(t);
                }     
                
            } catch (exxx) {
                e.message(t, "info", a)
            }
            
        },
        error: function (t, a, l) {
            toastr.clear();
            try {
                if (l == true) {
                    e.message(t, "error", a)
                } else {
                    window.parent.error_parrent(t);
                }                
            } catch (exxx) {
                e.message(t, "error", a)
            }            
        },
        warning: function (t, a, l) {           
            toastr.clear();
            try {
                if (l == true) {
                    e.message(t, "warning", a)
                } else {
                    window.parent.warning_parrent(t);
                }
            } catch (exx) {
                e.message(t, "warning", a)
            }            
        },
        success: function (t, a) {           
            toastr.clear();
            try {
                if (l == true) {
                    e.message(t, "success", a)
                } else {
                    window.parent.success_parrent(t);
                }
            } catch (exx) {
                e.message(t, "success", a)
            }
            
        },
        confirm: function (t, fnAccept) {
            //e.message(t, "success", a)
            toastr.success("<div style='text-align:right;'><button type='button' class='btn grey-cascade btn-sm' style='margin-right:15px;'>Không</button><button type='button' id='confirmButtonYes' class='btn blue btn-sm'>Đồng ý</button></div><div class='clearfix'></div>", t +'<br/><br/>',
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
        confirmDeleteOne: function (id, title, okCallBack) {
            var datarel = $('.' + id + '-' + id).attr('datarel');
            if (datarel !== "null") {
                datarel = " '" + datarel + "'";
            }
            else {
                datarel = "";
            }
            datarel = title + datarel + "?";

            toastr.warning("<div style='text-align:right;'><br /><button type='button' class='btn grey-cascade btn-sm' style='margin-right:15px;'>Không</button><button type='button' id='confirmDeleteOne' class='btn blue btn-sm'>Đồng ý</button><div><div class='clearfix'></div>", datarel,
                {
                    preventDuplicates: true,
                    closeButton: false,
                    allowHtml: true,
                    timeOut: 0,
                    showDuration: 300,
                    hideDuration: 1000,
                    showEasing: "swing",
                    extendedTimeOut: 0,
                    tapToDismiss: true,
                    positionClass: 'toast-center toast-top-center',
                    onShown: function (toast) {
                        $("#confirmDeleteOne").click(function () {
                            $('.modal-backdrop').remove();
                            toastr.remove();
                            okCallBack(id);
                        });
                    }
                });
        },
        ajaxPostData: function (url, data, fnSuccess, fnFailure, options) {
            var ajaxOptions = $.extend({}, {
                type: "POST",
                url: url,
                data: data,
                async: false,
                traditional: true,
                success: function (e) {   
                    // Không có lỗi
                    if (!ajaxErrorHandler(e) &&
                        fnSuccess != undefined &&
                        fnSuccess != null) {
                        fnSuccess(e);
                    }
                    else {
                        if (fnFailure != undefined &&
                            fnFailure != null) {
                            fnFailure(e);
                        }
                    }
                },
                error: function (e) {
                    showDefaultErrorMessage();
                }
            }, options);

            // Gửi dữ liệu
            $.ajax(ajaxOptions);
        },
        ajaxPostJson: function (url, jsonData, fnSuccess, fnFailure, options) {

        $.ajaxPostData(url, jsonData, fnSuccess, fnFailure, $.extend({}, options, { contentType: 'application/json; charset=utf-8' }));
        },
        ajaxPostModels: function (url, models, fnSuccess, fnFailure) {
            // Xây dựng dữ liệu
            var data = JSON.stringify({ models: models });

            $.ajaxPostJson(url, data, fnSuccess);
        },
    })
}(jQuery);

var ajaxSuccess = function () {
    alert("this is ajaxSuccess")
};

function ajaxErrorHandler(e) {
    
    toastr.clear();
    $('.fa-btn-sumbit').remove();
    try {
        $("body,html").animate({ scrollTop: 0 }, 400);
    } catch (e) {

    }
    var message = getErrors(e);
    if (message.length > 0) {
        var isWarning = false;

        try {
            if (e.Errors.Warning.errors.length > 0) {
                isWarning = true;
            }
        } catch (e) { }

        if (isWarning) {
            $.warning(message);
        } else {
            var isSuccess = false;

            try {
                if (data.Errors.Success.errors.length > 0) {
                    isSuccess = true;
                }
            } catch (e) { }
            if (isSuccess) {
                //$.success_parrent(message);
                window.parent.success_parrent(message);
            } else {
                $.error(message);
            }
        }

        var code = undefined;
        try {
            if (e.errors.message != undefined || e.Errors != undefined) {
                code = e.errors.message.code != undefined ? e.errors.message.code : e.Errors.message.code;
            }
        } catch (ex) {

        }

        // Hết session hoặc không có quyền
        if (code != undefined &&
            code == 403) {
            setTimeout(function () {
                // Làm mới trang
                location.reload();
            }, gOptions.NotyDisplayTimeout);
        }

        return true;
    }

    return false;
}

//function popupEditorSuccess(e, t, a, n) {
//    toastr.clear();
//    var r = getErrors(e);
//    if (r.length > 0) $.warning(r);
//    else {
//        $.success(n.Message), null != n.ReturnUrl && (location.href = n.ReturnUrl);
//        var i = n.GridID;
//        $("#" + n.PopupID).closest(".k-window, .k-overlay").remove(), $("#" + i).gridRefresh(), $(".k-overlay").remove();        
//    }
//}

function popupEditorSuccess(data, status, xhr, options) {
    
    $('.fa-btn-sumbit').remove();
    //debugger;
    //try {
    //    $("body,html").animate({ scrollTop: 0 }, 400);
    //} catch (e) {

    //}
    toastr.clear();
    var message = getErrors(data);
    var error = true;
   
    // Có dữ liệu trả về nhưng không phải là dữ liệu lỗi
    if (data.length > 0 && message.length == 0) {
        message = 'Có lỗi xảy ra';
        error = true;
    }
    // Không có lỗi
    else if (message.length == 0) {
        if (options.Message != null && options.Message != "") {
            message = options.Message;
        } else {
            message = 'Cập nhật thành công';
        }
        error = false;
    }

    if (message.length > 0) {
        if (error) {
            var isWarning = false;
            try {
                if (data.Errors.Warning.errors.length > 0) {
                    isWarning = true;
                }
            } catch (e) { }

            if (isWarning) {
                $.warning(message);
            } else {
                var isSuccess = false;

                try {
                    if (data.Errors.Success.errors.length > 0) {
                        isSuccess = true;
                    }
                } catch (e) { }
                if (isSuccess) {
                    //$.success_parrent(message);
                    window.parent.success_parrent(message);
                } else {
                    $.error(message);
                }
            }
        }
        else {
            //debugger;
            $.success(message);
            if (options != undefined) {

                if (options.ReturnUrl != null && options.ReturnUrl != "") {
                    location.href = options.ReturnUrl;
                }

                options.Data = data;         
                if (options.OnSuccess != null) {
                    getFunction(options.OnSuccess, ["options"]).apply(null, [options]);
                }
                var objPopup = options.FormID != null ? $('#' + options.FormID).closest('[data-role=window]') : null;
                if (objPopup != null) {

                    if (options.IsAdd) {
                        var continueAdd = $("#" + options.FormID).getAttr(gConts.ATTR_CONTINUE_ADD);

                        

                        // Tiếp tục thêm mới
                        if (eval(continueAdd)) {
                            try {
                                var popup = $(objPopup).data('kendoWindow');
                                $(objPopup).kendoWindowRefresh(options.Url, popup.title(), null, false);
                            } catch (e) {
                                setTimeout(function () {
                                    location.reload();
                                }, 200);
                            }

                        }
                        else {
                            $(objPopup).kendoWindowClose();
                            //setTimeout(function () {
                            //    // Làm mới trang
                            //    location.reload();
                            //}, gOptions.NotyDisplayTimeout);
                            return;
                        }
                    }
                    else {
                        // Nếu AutoClose null thì lấy từ options
                        options.AutoClose = options.AutoClose == null ? gOptions.PopupEditAutoClose : options.AutoClose;
                        //// Tự động đóng
                        //if (options.AutoClose) {
                        //    $(objPopup).kendoWindowClose();
                        //    return;
                        //}

                        $(objPopup).kendoWindowClose();
                        return;
                    }

                    $(objPopup).setFocus();
                }
            }
        }
    }
}

function popupEditorFailure(e, t, a) { }

function popupEditorBegin(e, t) { }

function onDataBinding(e) {
    try {
        var t = null;
        (t = this instanceof kendo.ui.Grid ? this : $(this).grid()) && (t.thead.find("[select='selecteAll']").each(function () {
            $(this).attr("checked", !1)
        }), gRecordIndex = (t.dataSource.page() - 1) * t.dataSource.pageSize(), gRecordIndex = isNaN(gRecordIndex) ? 0 : gRecordIndex)
    } catch (e) { }
}

function dataSourceRecordIndexIncrease() {
    return gRecordIndex = void 0 == gRecordIndex ? 1 : gRecordIndex + 1
}

function showPopup(e, t, a, n) {
    $.ajax({
        url: t,
        dataType: "html",
        beforeSend: function () { },
        success: function (t) {
            $('<div id="' + e + '"></div>').appendTo(document.body), $("#" + e).html(t), $("#" + e).kendoWindow({
                visible: !0,
                resizable: !1,
                modal: !0,
                deactivate: function () {
                    this.destroy()
                }
            }).data("kendoWindow").center()
        }
    })
}

function showPopupWindow(e, t, a, n) {
    var r = {
        modal: !!a || a,
        iframe: !1,
        draggable: !0,
        resizable: !0,
        scrollable: !0,
        actions: ["Close"],
        visible: !1,
        close: function () {
            $("#" + e).empty()
        }
    };
    n = $.extend(n, r), 0 == $("#" + e).length ? $("<div id='" + e + "' />").appendTo(document.body).kendoWindow(n).data("kendoWindow").center() : $("#" + e), $("#" + e).kendoWindowRefresh(t, "")
}

function showPopupWindowPost(popupID, url, data, isModal, options) {
    /// <summary>
    /// Hiển thị popup tự động khởi tạo kendo window
    /// </summary>
    /// <param name="popupID"></param>
    /// <param name="url"></param>
    /// <param name="isModal"></param>
    /// <param name="options"></param>
    var popup;
    var defaultOptions = {
        "modal": isModal ? true : isModal,
        "iframe": false,
        "draggable": true,
        //"title": "",
        //"width": 0,
        "resizable": true,
        "scrollable": true,
        "actions": ["Close"],
        "visible": false,
        "close": function () { $("#" + popupID).empty(); }
    };
    options = $.extend(options, defaultOptions);

    // Chưa tồn tại popup
    if ($("#" + popupID).length == 0) {
        popup = $("<div id='" + popupID + "' />").appendTo(document.body).kendoWindow(options).data('kendoWindow');

        // Canh vị trí
        popup.center();
    }
    else {
        popup = $('#' + popupID);
    }
    $("#" + popupID).kendoWindowRefreshPost(url, "", data);
}

function createTabLink(tenTab, link, id) {
    try {
        window.parent.createNewTab(tenTab, link, 'link_' + id, true);
    }
    catch {
        window.open(link, '_blank');
    }
}

function getErrors(e) {
    $('.fa-btn-sumbit').remove();
    var t = "";
    if (e.Errors || e.errors) {
        var a = void 0 == e.Errors ? e.errors : e.Errors;
        a && ($.each(a, function (e, a) {
            "errors" in a && $.each(a.errors, function () {
                t += this + "<br/>"
            })
        }), t = t.substr(0, t.lastIndexOf("<br/>")))
    }
    return t
}

function showDefaultErrorMessage() {
    $.error("Có lỗi xảy ra");
}

function toURL(e) {
    return e.value = e.value.toLowerCase(), e.value = e.value.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a"), e.value = e.value.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e"), e.value = e.value.replace(/ì|í|ị|ỉ|ĩ/g, "i"), e.value = e.value.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o"), e.value = e.value.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u"), e.value = e.value.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y"), e.value = e.value.replace(/đ/g, "d"), e.value = e.value.replace(/!|@|%|\^|\*|\(|\)|\+|\||\\|\=|\<|\>|\?|,|\.|\:|\;|\'|\"|\&|\#|\[|\]|~|$|\$|_/g, "-"), e.value = e.value.replace(/-+-/g, "-"), e.value = e.value.replace(/\/+\//g, "/"), e.value = e.value.replace(/ + /g, "-"), e.value = e.value.replace(/\/+-/g, "/"), e.value = e.value.replace(/-+\//g, "/"), e.value = e.value.replace(/ /g, "-"), e.value = e.value.replace(/^\-+|\-+$/g, ""), e
}

function error_handler(e) {
    if (alert(""), e.errors) {
        var t = "Errors:\n\n";
        $.each(e.errors, function (e, a) {
            "errors" in a && $.each(a.errors, function () {
                t += this + "\n\n"
            })
        }), alert(t)
    }
}

function LoadingMask(value) {
    if (value) {
        var html = "<div class=\"k-loading-mask\" style=\"width: 100%; height: 100%; top: 0px; left: 0px;background-color: #00000040;\"><span class=\"k-loading-text\">Loading...</span><div class=\"k-loading-image\"></div><div class=\"k-loading-color\"></div></div>";
        $(html).appendTo(document.body);
    } else {
        $(".k-loading-mask").remove();
    }
}
function loadingMaskEx(value) {
    if (value) {
        var html = '<div class="k-loading-mask" style="width: 100%;height: 100%;top: 0px;left: 0px;z-index: 100000;display: block;background-color: #607d8b2b;"><i class="fa fa-spinner fa-pulse fa-3x fa-fw" style="position:absolute;top:10px;left:50%;color:#397FAE;margin-top:20%;margin-left:-45px;font-size:90px;"></i></div>';
        $(html).appendTo(document.body);
    } else {
        $(".k-loading-mask").remove();
    }
}
function error_parrent(t) {
    $.error(t);
}

$(document).on("click", "*[data-toggle-selector]", function (e) {
    var selector = $(this).data("toggle-selector");

    if (selector != undefined) {
        var element = $(selector);

        element.toggle();

        $(this).find("span").removeClass("p-i-sort-asc");
        $(this).find("span").removeClass("p-i-sort-desc");

        $(this).find("span").addClass(element.css('display') == 'none' ? "p-i-sort-asc" : "p-i-sort-desc");
    }
});

/**********************************************************
 *  Live event handler
 **********************************************************/
// Bắt sự kiện không cho nhập giá trị ngoài giá trị combobox
$(document).on("change", "*[data-role=combobox]", function (evt) {
    var that = $(this);
    var combo = $(this).getKendoComboBox();

    // Kiểm tra giá trị
    if (combo.value() && combo.selectedIndex == -1) {
        combo._filterSource({
            value: "",
            field: combo.options.dataTextField,
            operator: "contains"
        });
        combo.value(null);
        //this.select(0);
    }
});

// Cascade
$(document).on("change", "*[data-cascade]", function (evt) {
    var that = $(this);
    $(that.data("cascade").split(',')).each(function () {
        that = $("#" + this);
        var role = that.data("role");

        if (role == "dropdownlist" ||
            role == "combobox") {

            if (role == "dropdownlist") {
                that.getKendoDropDownList().enable(true);
            }

            if (role == "combobox") {
                that.getKendoComboBox().enable(true);
            }

            that.comboBoxRefreshDataSource();
        }
        else {
            that.val(null);
        }
    });
});

// Đóng mở các element ẩn
$(document).on("click", "*[data-toggle-selector]", function (e) {
    var selector = $(this).data("toggle-selector");

    if (selector != undefined) {
        var element = $(selector);

        element.toggle();

        $(this).find("span").removeClass("p-i-sort-asc");
        $(this).find("span").removeClass("p-i-sort-desc");

        $(this).find("span").addClass(element.css('display') == 'none' ? "p-i-sort-asc" : "p-i-sort-desc");
    }
});

// Set min
$(document).on("change", "*[data-lessthan],*[data-greaterthan]", function () {
    var startValue = $(this).valueOfControl(),
        endID = $(this).data("lessthan"),
        step = parseFloat($(this).data("lessthan-step")),
        isLessThan = true;

    if (!endID) {
        endID = $(this).data("greaterthan");
        step = parseFloat($(this).data("greaterthan-step"));
        isLessThan = false;
    }

    if (endID && startValue) {
        var role = $(this).data("role");
        var endObj = null;
        var isDate = true;

        if (role == "timepicker") {
            endObj = $("#" + endID).data("kendoTimePicker");
        }
        else if (role == "datepicker") {
            endObj = $("#" + endID).data("kendoDatePicker");
        }
        else if (role == "datetimepicker") {
            endObj = $("#" + endID).data("kendoDateTimePicker");
        }
        else if (role == "numerictextbox") {
            endObj = $("#" + endID).data("kendoNumericTextBox");
            isDate = false;
        }

        if (endObj) {
            // Ngày
            if (isDate) {
                startValue = new Date(startValue.valueOf() + (isLessThan ? step : -step) * 1000);
            }
            else {
                startValue += isLessThan ? step : -step;
            }

            if (isLessThan)
                endObj.min(startValue);
            else
                endObj.max(startValue);
        }
    }
});

// Sự kiện keypress chuyển đến input tiếp theo
$(document).on("keypress", "*[data-grid-enter-next-focus]", function (e) {
    var code = (e.keyCode ? e.keyCode : e.which);
    if (code == 13) {
        var nextFocusSelector = $(this).data("grid-enter-next-focus");
        var moveToNextRow = $(this).data("grid-enter-to-next-row");

        var fnGetNextRow = function (currentRow) {
            var nextRow = currentRow.next();

            if (nextRow.hasClass("k-group-footer")) {
                return fnGetNextRow(nextRow);
            }

            if (nextRow.hasClass("k-grouping-row")) {
                return fnGetNextRow(nextRow);
            }

            return nextRow;
        }

        if (nextFocusSelector) {
            e.preventDefault();
            var row = $(this).closest("tr");
            var next = null;

            // Là dòng cuối cùng
            if ((row.is(':last-child') /* là dòng cuối */ || fnGetNextRow(row).length == 0 /* không tìm thấy dòng tiếp theo */) &&
                $("[data-grid-enter-next-focus]:last", row).is($(this))) {
                next = $(this).closest("table").find("[data-grid-enter-next-focus]:first")
            }
            else {
                next = $(nextFocusSelector, row);

                // Không tìm thấy trong dòng hiện tại hoặc có cờ chuyển sang dòng tiếp theo
                if (moveToNextRow == true || !next) {
                    // Tìm đến dòng tiếp theo
                    next = $(nextFocusSelector, fnGetNextRow(row));
                }
            }

            if (next) {
                var role = $(next).data("role");

                if (role == "numerictextbox") {
                    $(next).getKendoNumericTextBox().focus();
                }
                else {
                    $(next).focus();
                }
            }
        }
    }
});

/**********************************************************
 *  KendoUI Extend widget
 **********************************************************/

/**
 * Upload
 * @param {any} e
 */
function CMSUploadFile(e, url) {
    $("#" + e + " .input-file-attach").click();
    $("#" + e + " .input-file-attach").off('change');
    $("#" + e + " .input-file-attach").change(function () {
        //debugger;
        var id = "#" + e;
        var formData = new FormData();
        var arrFile = $(id + ' .input-file-attach').prop("files");
        for (var i = 0; i < arrFile.length; i++) {
            formData.append('fileAT', $(id + ' .input-file-attach').prop("files")[i]);
        }
        $.ajax({
            url: url,
            type: "POST",
            data: formData,
            contentType: false,
            cache: false,
            processData: false,
            success: function (file) {

                if (file.success) {
                    var ketqua = file.data;
                    var html = "";
                    //console.log($("#" + e + " .pmt-attachment-file li"));
                    for (var i = 0; i < ketqua.length; i++) {
                        var flag = true;
                        try {
                            $.each($("#" + e + " .pmt-attachment-file li"), function (index, item) {
                                //console.log('ddd', $($(this).children("input")['0']).attr('value'));
                                if ($($(this).children("input")['0']).attr('value').toString() == ketqua[i].ID.toString()) {
                                    flag = false;
                                    return false;
                                }
                            });
                        } catch (ex) { }
                        if (flag == true) {
                            html += '<li>' + ketqua[i].TenFile + ' <span onClick="CMSRemoveFile(this)" class="remove">&times;</span><input type="hidden" name="' + e + '" value="' + ketqua[i].ID + '" /><input type="hidden" name="' + e + '[]" value="' + ketqua[i].ID + '" /></li>';
                        }
                    }
                    $(id + ' .pmt-attachment-file').append(html);
                } else $.error(file.message);
            },
            async: true,
            beforeSend: function () {
                kendo.ui.progress($(id), true);
            },
            complete: function () {
                kendo.ui.progress($(id), false);
            }
        });
    });
}

function CMSRemoveFile(e) {
    $(e).parent().remove();
}

// Rút gọn ghi chú (Xem thêm)
function getShortSubstring(value, length) {
    if (value != null && value != "") {
        try {
            if (value.length > length) {
                var html = '<div class="shortcontent">' + kendo.toString(value.substring(0, length)) + "...</div>";
                html += '<div class="allcontent" style="display: none;">' + kendo.toString(value) + '</div>';
                html += '<span><a href="javascript://nop/" class="morelink" style="color: #F44336;display: block;text-align:right;padding-top: 5px;">Xem thêm</a></span>';
                return html;
            }
            else {
                return kendo.toString(value);
            }
        } catch (exxx) {
            return kendo.toString(value);
        }
    } else {
        return "";
    }
}

//Events Grid Kendo
function detailExpand(ev) {
    var expandedRow = $(ev.sender.wrapper).data('expandedRow');
    if (expandedRow && expandedRow[0] !== ev.masterRow[0]) {
        var $grid = $(ev.sender.wrapper).data('kendoGrid');
        $grid.collapseRow(expandedRow);
    }
    $(ev.sender.wrapper).data('expandedRow', ev.masterRow);
}
function detailCollapse(ev) {
    var expandedRow = $(ev.sender.wrapper).data('expandedRow');
    expandedRow.next().remove();
}
function onChange(e) {
    var row = this.select();
    if (row != null) {
        if (row.next(".k-detail-row").is(":visible")) {
            e.sender.collapseRow(row);
        } else {
            e.sender.expandRow(row);
        }
    }
}

function handleErrors(e) {
    if (e.errors) {
        var message = "";
        $.each(e.errors, function (key, value) {
            if ('errors' in value) {
                $.each(value.errors, function () {
                    message += this + "</br>";
                });
            }
        });
        $.warning(message);
    }
}

/**********************************************************
 *  Event Element
 **********************************************************/
//Check Number
function IsNumberKey(evt) {
    var e = evt || window.event;
    var charCode = e.which || e.keyCode;
    if (charCode > 32 && (charCode <= 47 || charCode > 57) || charCode == 13) {
        return false;
    }
    if (e.shiftKey) {
        return false;
    }

    return true;
}


// bổ sung get token phục vụ cho các action có validate
function getToken() {
    return $('input[name="__RequestVerificationToken"]').val();
}

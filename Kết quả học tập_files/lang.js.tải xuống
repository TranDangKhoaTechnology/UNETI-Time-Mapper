var controls = $('[langid]');

var uid = $("#UID").val();
var flag = $("#CurrentLanguage").val();
flag = (flag == undefined || flag == null) ? "" : flag;
if ($.trim(flag)) {
    flag = "_" + flag;
}
//
var jsonRiengUrl = "../Content/lang/" + uid + flag + ".json";
var jsonChungUrl = "../Content/lang/all" + flag + ".json";

var langChungData = null;
var langRiengData = null;
$.ajax({
    url: jsonChungUrl,
    dataType: "json",
    async: false,
    success: function (jsdata) {
        langChungData = jsdata;
    },
    error: function () {

    },
});

$.ajax({
    url: jsonRiengUrl,
    dataType: "json",
    async: false,
    success: function (jsdata) {
        langRiengData = jsdata;
    },
    error: function () {

    },
});

if (langChungData) {
    $("[lang]").each(function (index) {
        var val = langChungData[$(this).attr("lang")];
        $(this).html(val);
        $(this).attr("placeholder", val);
    });
}

if (langRiengData) {
    $("[lang]").each(function (index) {
        var val = langRiengData[$(this).attr("lang")];
        $(this).html(val);
        $(this).attr("placeholder", val);
    });
}

if (!(controls.length === 0)) {
    for (const i of controls) {
        var id = $(i).attr("langid");//controls[i].getAttribute('langid');
        var tmp = id.split("_");
        var keyname = null;
        if (tmp.length > 1) {
            var arr = [];
            for (var j = 0; j < tmp.length; ++j) {
                arr[j] = String(tmp[j]);
                if (String(tmp[j]).includes("-")) {
                    break;
                }
            }
            keyname = (arr.join('_'));
        }
        else {
            keyname = String(tmp);
        }

        var attr = keyname.split("-").pop();

        /* var keyname = viewname + '-' + name + ((attr == undefined || attr == null) ? '' : ('-' + attr));*/
        var keyValChung = null;
        if (langChungData) {
            keyValChung = langChungData[keyname];
        }       
        keyValChung = (keyValChung == undefined || keyValChung == null) ? '' : keyValChung;
        var keyValRieng = null;
        if (langRiengData) {
            keyValRieng = langRiengData[keyname];
        }
        
        keyValRieng = (keyValRieng == undefined || keyValRieng == null) ? '' : keyValRieng;
        var ctrl = $("[langid='" + id + "']")[0];
        if ($.trim(keyValChung)) {
            switch (attr) {
                case "combobox":
                    if (ctrl) {
                        $(ctrl).find("option:first-child").html(keyValChung);
                    }              
                    break;
                //case "kendocombobox":
                //    var ctrlID = ctrl.getAttribute("id");
                //    var comboBox = $('#' + ctrlID).data("kendoComboBox");
                //    if (comboBox) {
                //        $(comboBox.input).attr('placeholder', keyValChung);
                //    }
                //    break;
                default:
                    if (ctrl) {
                        ctrl[attr] = keyValChung;
                    }
            }
        }

        if ($.trim(keyValRieng)) {
            if (attr == "combobox") {
                if (ctrl) {
                    $(ctrl).find("option:first-child").html(keyValRieng);
                }
               
            }
            else {
                if (ctrl) {
                    ctrl[attr] = keyValRieng;
                }
                
            }
        }
    }
}

function ChangeKendoCbo() {
    $('[langid*="kendocombobox"]').each(function (index) {
        var id = $(this).attr("langid");
        var tmp = id.split("_");
        var keyname = null;
        if (tmp.length > 1) {
            var arr = [];
            for (var j = 0; j < tmp.length; ++j) {
                arr[j] = String(tmp[j]);
                if (String(tmp[j]).includes("-")) {
                    break;
                }
            }
            keyname = (arr.join('_'));
        }
        else {
            keyname = String(tmp);
        }

        var attr = keyname.split("-").pop();

        /* var keyname = viewname + '-' + name + ((attr == undefined || attr == null) ? '' : ('-' + attr));*/
        var keyValChung = null;
        if (langChungData) {
            keyValChung = langChungData[keyname];
        }
        keyValChung = (keyValChung == undefined || keyValChung == null) ? '' : keyValChung;
        var keyValRieng = null;
        if (langRiengData) {
            keyValRieng = langRiengData[keyname];
        }

        keyValRieng = (keyValRieng == undefined || keyValRieng == null) ? '' : keyValRieng;
        var ctrl = $("[langid='" + id + "']")[0];
        var ctrlID = ctrl.getAttribute("id");
        if ($.trim(keyValChung)) {
            var comboBox = $('#' + ctrlID).data("kendoComboBox");
            if (comboBox) {
                $(comboBox.input).attr('placeholder', keyValChung);
            }
            
        }

        if ($.trim(keyValRieng)) {
            var comboBox = $('#' + ctrlID).data("kendoComboBox");
            if (comboBox) {
                $(comboBox.input).attr('placeholder', keyValRieng);
            }
        }
    });
}

function loadURL(url) {
    window.location.href = url;
}
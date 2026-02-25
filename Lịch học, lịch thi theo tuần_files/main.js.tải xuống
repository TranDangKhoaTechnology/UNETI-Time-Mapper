$(".menu-btn > a").on("click", function(){
    $(".menu-top").toggleClass("active");
    return false;
});

jQuery(document).ready(function($) {
    if($('.border-scroll').length>0){
            $('.border-scroll').niceScroll({
                cursorborder: 'none',
                cursorcolor: '#999',
                autohidemode:'leave'
            });
        }
});

$(document).ready(function() {
    $('.tooltip').tooltipster(
        {
        delay: 100,
        maxWidth: 450,
        speed: 300,
        interactive: true,
        contentAsHTML: true,  
        animation: 'grow',
        }
    );
});

$(document).ready(function() {
    var mnArr = [];
    var mainMn = $('#mainMn');
    mainMn.find('li').each(function () {
        var obj = $(this);
        mnArr.push(obj);
    });
    $('.mainMnBtn').bind('click', function () {
            //show modal
            mainMn.css({display:'block', opacity:0, marginTop:'10%'});
            mainMn.animate({ opacity:1, marginTop:'0%'}, 200);

            mainMn.find('li').css({ marginTop: '20px', fontSize: '2em', opacity:0});
            for(var i=0; i < mnArr.length; i++){
                mnArr[i].delay(i*50 + 100).animate({ marginTop: '0px', fontSize: '1.6em', opacity:1},160, 'linear');
            }
    });

    mainMn.find('.closeMn').bind('click', function () {
         mainMn.animate({ opacity:0}, 200, function () {
            mainMn.hide();
        });
    });
});

$(document).ready(function () {

    $('.collapse').on('show.bs.collapse', function () {
        $('.collapse.in').collapse('hide');
    });

    $('#OpenAccordionAll').on('click', function () {
        $('.table-custom .collapse').collapse("toggle");
    });

});

// chatPopup

$(document).ready(function() {  

        var id = '#dialog';
    
        var maskHeight = $(document).height();
        var maskWidth = $(window).width();
    
        $('#mask').css({'width':maskWidth,'height':maskHeight});
        $('#mask').fadeIn(500); 
        $('#mask').fadeTo("slow",0.9);  
    
        var winH = $(window).height();
        var winW = $(window).width();
    
        $(id).css('top',  winH/3-$(id).height()/3);
        $(id).css('left', winW/2-$(id).width()/2);
    
        $(id).fadeIn(2000);     
    
    $('.window .closePopup').click(function (e) {
        e.preventDefault();
        
        $('#mask').hide();
        $('.window').hide();
        $('#boxes').remove();
    });
});

$(document).ready(function(){
    $( "a.scrollLink" ).click(function( event ) {
        event.preventDefault();
        $("html, body").animate({ scrollTop: $($(this).attr("href")).offset().top }, 500);
    });
});

$(document).ready(function() {
    jQuery(document).ready(function() {
        jQuery("#accordion-menu").jqueryAccordionMenu();
        jQuery(".colors a").click(function() {
            if ($(this).attr("class") != "default") {
                $("#accordion-menu").removeClass();
                $("#accordion-menu").addClass("accordion-menu").addClass($(this).attr("class"));
            } else {
                $("#accordion-menu").removeClass();
                $("#accordion-menu").addClass("accordion-menu");
            }
        });

    });
});



$(document).ready(function () {
    jQuery(document).ready(function () {
        jQuery("#accordion-menu-mb").jqueryAccordionMenu();
        jQuery(".colors a").click(function () {
            if ($(this).attr("class") != "default") {
                $("#accordion-menu-mb").removeClass();
                $("#accordion-menu-mb").addClass("accordion-menu").addClass($(this).attr("class"));
            } else {
                $("#accordion-menu-mb").removeClass();
                $("#accordion-menu-mb").addClass("accordion-menu");
            }
        });
    });
});

(function($, window, document, undefined) {
    var pluginName = "jqueryAccordionMenu";
    var defaults = {
        speed: 300,
        showDelay: 0,
        hideDelay: 0,
        singleOpen: true,
        clickEffect: true
    };

    function Plugin(element, options) {
        this.element = element;
        this.settings = $.extend({}, defaults, options);
        this._defaults = defaults;
        this._name = pluginName;
        this.init()
    };
    $.extend(Plugin.prototype, {
        init: function() {
            this.openSubmenu();
            this.submenuIndicators();
            if (defaults.clickEffect) {
                //this.addClickEffect()
            }
        },
        openSubmenu: function() {
            $(this.element).children("ul").find("li").bind("click touchstart", function(e) {
                e.stopPropagation();
                e.preventDefault();
                if ($(this).children(".submenu").length > 0) {
                    if ($(this).children(".submenu").css("display") == "none") {
                        $(this).children(".submenu").delay(defaults.showDelay).slideDown(defaults.speed);
                        $(this).children(".submenu").siblings("a").addClass("submenu-indicator-minus");
                        if (defaults.singleOpen) {
                            $(this).siblings().children(".submenu").slideUp(defaults.speed);
                            $(this).siblings().children(".submenu").siblings("a").removeClass("submenu-indicator-minus")
                        }
                        return false
                    } else {
                        $(this).children(".submenu").delay(defaults.hideDelay).slideUp(defaults.speed)
                    }
                    if ($(this).children(".submenu").siblings("a").hasClass("submenu-indicator-minus")) {
                        $(this).children(".submenu").siblings("a").removeClass("submenu-indicator-minus")
                    }
                }
                window.location.href = $(this).children("a").attr("href")
            })
        },
        submenuIndicators: function () {
            if ($(this.element).find('.submenu').length > 0) {
                $(this.element).find('.submenu').siblings('a').append('<span class="submenu-indicator"><i class="fa fa-angle-down" aria-hidden="true"></i></span>');
            }
        },
        //addClickEffect: function() {
        //    var ink, d, x, y;
        //    $(this.element).find("a").bind("click touchstart", function(e) {
        //        $(".ink").remove();
        //        if ($(this).children(".ink").length === 0) {
        //            $(this).prepend("<span class='ink'></span>")
        //        }
        //        ink = $(this).find(".ink");
        //        ink.removeClass("animate-ink");
        //        if (!ink.height() && !ink.width()) {
        //            d = Math.max($(this).outerWidth(), $(this).outerHeight());
        //            ink.css({
        //                height: d,
        //                width: d
        //            })
        //        }
        //        x = e.pageX - $(this).offset().left - ink.width() / 2;
        //        y = e.pageY - $(this).offset().top - ink.height() / 2;
        //        ink.css({
        //            top: y + 'px',
        //            left: x + 'px'
        //        }).addClass("animate-ink")
        //    })
        //}
    });
    $.fn[pluginName] = function(options) {
        this.each(function() {
            if (!$.data(this, "plugin_" + pluginName)) {
                $.data(this, "plugin_" + pluginName, new Plugin(this, options))
            }
        });
        return this
    }
})(jQuery, window, document);

$(document).ready(function () {
    $("#full-table").click(function () {
        $("#full-resize-table").toggleClass("container").toggleClass("container-fluid");
        $('#full-table span').toggleClass("glyphicon-resize-full").toggleClass("glyphicon-resize-small");
    });
    $('#config-tool-cog').on('click', function () { $('#config-tool').toggleClass('closed') });
    $('.sidebar-menu-inner').on('click', function () { $('#config-tool').addClass('closed') });

    if (typeof (Storage) !== "undefined") {
        var cl_color = localStorage.getItem("theme-sv-color");
        if (cl_color !== "undefined" && cl_color != "" && cl_color != null) {
            $('body').attr("id", cl_color);
        }
    }
});


$('body').on('click', '#config-tool-options .theme-color li', function () {
    if (typeof (Storage) !== "undefined") {
        var cl_color = $(this).data('style');
        if (cl_color !== "undefined" && cl_color != "" && cl_color != null) {
            localStorage.setItem("theme-sv-color", cl_color);
            $('body').attr("id", cl_color);
        }
    }
});

function loadingMaskEx(value) {
    if (value) {
        var html = '<div class="k-loading-mask" style="width: 100%;height: 100%;top: 0px;left: 0px;z-index: 100000;display: block;background-color: #607d8b2b;"><i class="fa fa-spinner fa-pulse fa-3x fa-fw" style="position:absolute;top:10px;left:50%;color:#397FAE;margin-top:20%;margin-left:-45px;font-size:90px;"></i></div>';
        $(html).appendTo(document.body);
    } else {
        $(".k-loading-mask").remove();
    }
}
function formatNumber(value) {
    return ("'" + value + "'").replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1.').replace(/'/g, "")
}

$(document).ready(function(){ 
	var touch 	= $('#resp-menu');
	var menu 	= $('.menu_news');
 
	$(touch).on('click', function(e) {
		e.preventDefault();
		menu.slideToggle();
	});
	
	$(window).resize(function(){
		var w = $(window).width();
		if(w > 767 && menu.is(':hidden')) {
			menu.removeAttr('style');
		}
	});
	
});
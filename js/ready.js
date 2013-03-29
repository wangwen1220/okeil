//////////////////////////////////////////////////////////////////////////////////////
// jQuery for lieko.com
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// @name ready.js
// @author Steven - http://www.lieko.com/wen
// @version 0.2
// @date April 20, 2010
// @copyright (c) 2010 Lieko Co. Ltd (lieko.com)
//////////////////////////////////////////////////////////////////////////////////////
$(function(){
  /*======判断窗口大小隐藏MSN======*/
  //if(!($.browser.msie && $.browser.version<7)){
    //当窗口宽度小于1186px时隐藏MSN
    setInterval(function(){$(window).width() < 1186? $("#extra1").hide() : $("#extra1").show();}, 1);
  //}

  /*======#sidebar导航效果======*/
  var $menu = $("#sidebar > dl.menu");
  var $menu_dd = $menu.children("dd").hide();
  var $sub_menu_dt = $menu_dd.prev("dt");/*有二级菜单的dt元素*/

  $sub_menu_dt.addClass("closed").hover(function(){
    $(this).addClass('hover');
  },function(){
    $(this).removeClass('hover');
  }).click(function(){
    $(this).toggleClass("closed").nextUntil("dt").toggle();
    return false;
  });

  //当前子菜单效果
  var h1_title = $("#pros > h1").attr("title") || "notitle";//如果没有title则赋值"notitle"，避免IE中的显示问题
  //var h1_text = $("#pros.case > h1").text();
  $menu.children().has("a[title='"+h1_title+"'], a[title^='"+h1_title+"']").addClass("current").filter("dd").prevAll("dt:first").removeClass("closed").nextUntil("dt").show();
  /*$("#sidebar > dl.menu a[href*='model=']").each(function(){
    if(h1_text.indexOf(this.title) == 0) $(this).parent().addClass("current").not("dt").prevAll("dt:first").removeClass("closed").nextUntil("dt").show();
  });*/

  //单击标题关闭当前菜单，默认关闭Products Technology菜单
  $("#sidebar > h2").css("cursor", "pointer").click(function(){
    //$sub_menu_dt.toggleClass("closed").nextUntil("dt").toggle();
    $(this).next("dl, ul").slideToggle("fast");
    return false;
  }).filter(".closed").next(":not(:has('.current'))").slideToggle("fast");

  /*======若产品只有一页，则隐藏分页======*/
  if($("#pager > em").text() == "1 / 1") $("#pager").addClass("hide");

  /*======若Products页面没有相关产品，则隐藏底部分页======*/
  if($("#pros > div.pager > span.list").text() == "") $("#pros > div.pager").hide();

  /*======隐藏第一行产品的虚线======*/
  $("body#type_list #pros > div.list:first").addClass("first");

  /*======渐隐产品编号======*/
  /*$("#pros > div").hover(function(){
    $(this).children(".number").stop(true, true).animate({"margin-top": "-20px"}, 300);
  },function(){
    $(this).children(".number").stop(true, true).animate({"margin-top": "0"}, 600);
  });*/

  /*======渐隐首页产品标签======*/
  /*$("#home #pros > div > span").hover(function(){首页改版已经取消
    $(this).stop(true, true).animate({"margin-top": "-5px", "margin-left": "-6px"}, 300);
  },function(){
    $(this).stop(true, true).animate({"margin-top": "0", "margin-left": "0"}, 600);
  });*/

  /*======搜索框行为======*/
  $('#sform input:text').focus(function(){
    //with (event.srcElement) 不兼容FF
    if (this.value == "Keywords, Number or Model, For example: IP-4-A" || this.value == "Palabras claves, número o modelo, por ejemplo: IP-4-A") this.value = ""
  }).blur(function(){
    if(this.value == "") this.value = this.defaultValue
  });

  /*======#color产品颜色切换效果======*/
  var $picview_a = $("#picview > a");
  var $color_a = $("#color > li > a");

  //$picview_a.click(function(){return false});

  $color_a.click(function(){
    var this_href = this.href;
    $picview_a.attr("href", this_href).children("img").stop(true, true).fadeOut(800, function(){
      var $this =$(this);
      if(this.src == this_href){
        $this.fadeIn(800);
      }else{
        $this.attr("src", this_href).load(function(){
          $this.fadeIn(800);
        });
      }
    });
    $(this).parent().addClass("current").siblings().removeClass("current");
    return false;
  });

  /*======移除统计图片alt属性======*/
  $("#footer img").removeAttr("alt");

  /*======外部链接======*/
  $("#links a.external_links").attr("target", "_blank");

  /*======IE6 only======*/
  if($.browser.msie && $.browser.version<7){//找出IE6浏览器
    try{
      document.execCommand("BackgroundImageCache", false, true);//修正IE6无法缓存背景图片的bug－不过好像不起作用，奶奶的IE6
    }catch(err){}

    //修正MSN问题
    var $msn=$('#extra1 a.msn');
    //$("#extra1 .msn").fixed({right:"5px",top:"10px"});
    $msn.css("position", "absolute");
    $(window).scroll(function(){
      $msn.animate({'top': $(document).scrollTop()+30+'px'}, 1);
    });
  }

  /*======safari only======*/
  if($.browser.safari){//判断是否是safari浏览器
    $("body").css("margin", "0");/*在css中已经设置了，但是不清楚为什么无效？*/
    $("ul").css("padding", "0");
  }

  /*======切换产品显示方式======*/
  //if($('body.pro_list').length){
    //$.getScript("/js/jqcookie.js", function(){*/
      //var $pro_box = $("#pros > div:not(.pager)");
      //var pros_view_as = $.cookie('lieko_pros_view_as');
      //if(pros_view_as){
        //$pro_box.attr('class', pros_view_as);
        //$("#pros span.view_as > span." + pros_view_as).addClass('current').siblings().removeClass('current');
      //}

      //$("#pros span.view_as > span:not(.current)").live('click', function(){
        //$pro_box.toggleClass("gallery list");
        //$(this).add($(this).siblings()).toggleClass("current");
        ////$.cookie('lieko_pros_view_as', $pro_box.attr('class'));
      //});
    //});
  //}

  /*======加载并执行jquery.lazyload======*/
  if($("body.jqbox.jqpager").length){
    $.getScript("/js/jqlazyload.js", function(){
      $("#pros > div.hide > a > img").lazyload({
        placeholder: "/img/loading.gif",
        effect: "fadeIn"
      });
    });
  }

  /*======加载并执行jquery.lightbox======*/
  if($("body.jqbox").length){
    $.getScript("/js/jqbox.js", function(){
      $('#pros > div > h3 > a[href$=".jpg"]').lightBox();
      $('#main .jqbox').each(function(){
          $(this).find('a[href$=".jpg"]').lightBox();
      });
    });
  }

  /*======Products页面产品图片展示效果======*/
  /*$probox = $('#pros.prolist > div.box');
  var x = 8, y = 16;
  $probox.children('a.img').toggle(function(){
    $(this).parent().addClass("img_view").siblings('.box, .pager').hide();
    return false;
  },function(){
    $(this).parent().removeClass("img_view").siblings('.box, .pager').show();
  });*/

  /*======加载并执行jquery.quickpaginate======*/
  if($("body.jqpager").length){
    $.getScript("/js/jqpager.js", function(){
      $("#pros > div").quickpaginate({
        perpage: 20,
        pager: $("#jqpager")
      });

      $("#pros > div.poster").quickpaginate({
        perpage: 1,
        pager: $("#jqpager")
      });
    });
  }
});

/*##################加载scroll2top##################*/
$.getScript("/js/scroll2top.min.js");

/*##################页面加载完执行##################*/
$(window).load(function(){
  /*======首页Slideshow效果======*/
  var $gallery = $("#gallery");
  $gallery.children('dt').css({opacity: 0}).siblings('dd').css({opacity: 0.7, bottom: "-100px"});
  $gallery.children('dt.show').css({opacity: 1}).next('dd').css({bottom: 0});
  setTimeout(function(){gallery($gallery)}, 1400);

  $gallery.hover(function(){
    $(this).css("opacity", 0.9);
  }, function(){
    $(this).css("opacity", 1);
  });

  /*======滚动banner效果======*/
  var $num = $("#banner > ul.num > li");
  var numLen = $num.length;
  var index = 1;
  var bannerTimer;
  $num.mouseover(function() {
    index = $num.index(this);
    showImg(index);
  }).eq(0).mouseover();

  //滑入停止动画，滑出开始动画
  $('#banner').hover(function() {
    clearInterval(bannerTimer);
  }, function() {
    bannerTimer = setInterval(function() {
      if(++index == numLen) index = 0;
      showImg(index);
    }, 5000);
  }).trigger("mouseleave");

  /*======新产品列表滚动效果======*/
  /*var $new_pros = $("#new_pros > ul");
  var scrollerTimer;
  $new_pros.hover(function(){
    clearInterval(scrollerTimer);
  }, function(){
    scrollerTimer = setInterval(function(){
      scroller($new_pros);
    }, 5000);
  }).trigger("mouseleave");*/

  /*======#picview产品放大效果======*/
  if($("body.view").length){
    var $picview_aimg = $("#picview > a:first");
    var $hides = $picview_aimg.next().add($picview_aimg.parent().next());
    var picview_img = new Image();
    picview_img.src = $picview_aimg.children("img")[0].src;
    var zoom_width = picview_img.width > 587 ? 587 : picview_img.width;

    $picview_aimg.toggle(function(){
      $(this).animate({width: zoom_width + "px", height: zoom_width + "px"}, 300);
      $hides.hide();
    },function(){
      $(this).animate({width: "307px", height: "307px"}, 300, function(){
        $hides.show();
      });
    });
  }

  /*======热卖产品滚动效果======*/
  var page = 1;
  var i = 5;//每版放5个图片
  var scrollerTimer;
  var $tips = $("#hot_products > div.tips > span");
  var $prolist = $("#hot_products > div.prolist > ul");
  var $btn =  $("#hot_products > div.btn > span");
  var len = $prolist.children("li").length;
  var page_count = Math.ceil(len / i);//只要不是整数，就往大的方向取最小的整数
  var page_width = $prolist.parent().width();//获取框架内容的宽度,不带单位

  $('#hot_products').hover(function(){
    clearInterval(scrollerTimer);
  }, function(){
    scrollerTimer = setInterval(function(){
      scroll_next();
    }, 5000);
  }).trigger("mouseleave");

  //next按钮
  $btn.filter(".next").click(scroll_next = function() {
    if(!$prolist.is(":animated")) {
      if(page == page_count) {//已经到最后一个版面了,如果再向后，必须跳转到第一个版面。
        $prolist.animate({ left: 0 }, 1000);//通过改变left值，跳转到第一个版面
        page = 1;
      }else {
        $prolist.animate({ left: '-=' + page_width }, 1000);//通过改变left值，达到每次换一个版面
        page++;
      }
      $tips.eq(page-1).addClass("current").siblings().removeClass("current");
    }
  });

  //prev按钮
  $btn.filter(".prev").click(function() {
    if(!$prolist.is(":animated")) {
      if(page == 1) {//已经到第一个版面了,如果再向前，必须跳转到最后一个版面。
        $prolist.animate({ left: '-=' + page_width*(page_count-1) }, 1000);//通过改变left值，跳转到最后一个版面
        page = page_count;
      }else {
        $prolist.animate({ left: '+=' + page_width }, 1000);//通过改变left值，达到每次换一个版面
        page--;
      }
      $tips.eq(page-1).addClass("current").siblings().removeClass("current");
    }
  });

    //验证VIP注册表单
  if($('#vip_form').length){
    $.getScript("/js/jq.validate.js", function(){
      //电话号码验证
      jQuery.validator.addMethod("tel", function(value, element) {
        value = value.replace(/\s+/g, "");
        var patrn = /^[+]{0,1}(\d){1,3}[ ]?([-]?((\d)|[ ]){1,12})+$/;//必须以数字开头，除数字外，可含有“-”
        return this.optional(element) || value.match(/^(([0\+]\d{2,3}-)?(0\d{2,3})-)?(\d{7,8})(-(\d{3,}))?$/) || patrn.test(value);
      }, "Please specify a valid tel number");

      //联系电话码验证
      jQuery.validator.addMethod("phone", function(value, element) {
        value = value.replace(/\s+/g, "");
        var mobile = /^((0|\+44)7(5|6|7|8|9){1}\d{2}\s?\d{6})$/
        var tel = /^(([0\+]\d{2,3}-)?(0\d{2,3})-)?(\d{7,8})(-(\d{3,}))?$/;
        var patrn = /^[+]{0,1}(\d){1,3}[ ]?([-]?((\d)|[ ]){1,12})+$/;//必须以数字开头，除数字外，可含有“-”
        return this.optional(element) || value.match(tel) || value.match(mobile) || patrn.test(value);
      }, "Please specify a valid phone number");
      $('#vip_form').validate({
        rules: {
          country: "required",
          company: "required",
          name: {
            required: true,
            minlength: 2
          },
          email: {
            required: true,
            email: true
          },
          //tel: "required",
          tel: {
            required: true,
            phone: true
          },
          fax: {
            required: true,
            tel: true
          }
        },

        messages: {
          country: "Please enter your country",
          company: "Please enter your company name",
          name: {
            required: "Please enter your name",
            minlength: "Your name must consist of at least 2 characters"
          },
          email: "Please enter a valid email address",
          tel: {
            required: "Please enter your tel number",
            minlength: "Please specify a valid phone number"
          },
          fax: {
            required: "Please enter your fax number",
            minlength: "Please specify a valid fax number"
          }
        }
      });
    });
  }

  /*======about.asp图片效果======*/
  /*(if($("body#about").length){
    $("#main div.workshop img").mouseover(function(){
      var position=$(this).position();
      if(position.left!='0'){
        $(this).animate({left: '0px'},1000);
      }
      else{
        $(this).animate({left: '-260px'},1000);
      }
    });
  }*/
});

/*##################gallery函数##################*/
function gallery($obj){
  //if no images have the show class, grab the first image
  var $dt_current = $obj.children('dt.show');

  //Get next image, if it reached the end of the slideshow, rotate it back to the first image
  var $dt_next = $dt_current.next('dd').next('dt').length ? $dt_current.next('dd').next('dt') : $obj.children('dt:first');

  //Set the fade in effect for the next image, show class has higher z-index
  $dt_next.css({opacity: 0}).addClass('show').fadeTo(1000, 1);

  //Hide the current image
  $dt_current.fadeTo(1000, 0).removeClass('show');

  //$dt_current.next('dd').removeClass('show').animate({bottom: "-100px"}, 200);//这个效果在NND IE6中惨不忍睹！
  $dt_current.next('dd').removeClass('show').css({bottom: "-100px"});

  $dt_next.next('dd').addClass('show').animate({bottom: 0}, 400);

  setTimeout(function(){gallery($obj)}, 4000);
}

/*##################通过控制top，来显示不同的幻灯片##################*/
function showImg(index) {
  var bannerHeight = $("#banner").height();
  $("#banner > ul.slider").stop(true, false).animate({ top: -bannerHeight * index }, 600);
  $("#banner > ul.num > li").removeClass("on").eq(index).addClass("on");
}

/*##################新产品列表滚动##################*/
/*function scroller($obj){
  var scrollHeight = $obj.children("li:first").height();//获取行高
  $obj.animate({marginTop: -scrollHeight + "px"}, 200, function(){
    $obj.css("margin-top", 0).children("li:first").appendTo($obj);//把元素移到最后
  });
}*/
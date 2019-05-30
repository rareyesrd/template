$(window).scroll(function() {    
    var scroll = $(window).scrollTop();

    if (scroll >= 10) {
        $(".clearHeader").addClass("darkHeader");
    } else {
       $(".clearHeader").removeClass("darkHeader");
    }
});




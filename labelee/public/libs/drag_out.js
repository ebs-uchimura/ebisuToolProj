// prevent pulldown reload
var $target = $(window);
var windowScrollTop;
var prevY = -1;
var direction = 0; // 0:neutral, -1:gotoTop, 1:gotoBottom
$target.on('touchstart', function(e){
    windowScrollTop = $target.scrollTop();
    prevY = e.originalEvent.changedTouches[0].pageY;
    direction = 0;
});
$target.on('touchmove', function(e){
    var currentY = e.originalEvent.changedTouches[0].pageY;
    if (direction == 0 && prevY != -1) {
        if (currentY > prevY) { direction = -1; }
        if (currentY < prevY) { direction =  1; }
    }
    if (windowScrollTop <= 0 && direction == -1) {
        e.preventDefault();
        return false;
    }
    prevY = currentY;
});
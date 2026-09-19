(function(){
    const $window  = $(window);
    //要素の取得
    var canvas = document.getElementById('labelee_canvas');
    var btn = document.getElementById('add_image');

    //要素内のクリックされた位置を取得するグローバル（のような）変数
    var x;
    var y;

    //標準Z軸
    var zind = 1001;

    //マウス監視
    var elements = document.getElementsByClassName("drag-and-drop");
    //マウスが要素内で押されたとき、又はタッチされたとき発火
    for(var i = 0; i < elements.length; i++) {
        elements[i].addEventListener("mousedown", mdown, false);
        elements[i].addEventListener("touchstart", mdown, false);
    }

    //マウスが押された際の関数
    function mdown(e) {
        //クラス名に .drag を追加
        this.classList.add("drag");
        //タッチデイベントとマウスのイベントの差異を吸収
        if(e.type === "mousedown") {
            var event = e;
        } else {
            var event = e.changedTouches[0];
        }
        //要素内の相対座標を取得
        x = event.pageX - this.offsetLeft;
        y = event.pageY - this.offsetTop;
        if(y < screen.height/2){
            //ムーブイベントにコールバック
            document.body.addEventListener("mousemove", mmove, false);
            document.body.addEventListener("touchmove", mmove, false);
        } else {
            x = event.pageX;
            y = event.pageY;
        }
    }

    //マウスカーソルが動いたときに発火
    function mmove(e) {
        $window.on('touchmove.noScroll', e => {
            e.preventDefault();
        });
        //ドラッグしている要素を取得
        var drag = document.getElementsByClassName("drag")[0];
        //同様にマウスとタッチの差異を吸収
        if(e.type === "mousemove") {
            var event = e;
        } else {
            var event = e.changedTouches[0];
            //スクロール禁止
            no_scroll();
        }
        //フリックしたときに画面を動かさないようにデフォルト動作を抑制
        e.preventDefault();
        //マウスが動いた場所に要素を動かす
        drag.style.top = event.pageY - y + "px";
        drag.style.left = event.pageX - x + "px";
        //マウスボタンが離されたとき、またはカーソルが外れたとき発火
        drag.addEventListener("mouseup", mup, false);
        document.body.addEventListener("mouseleave", mup, false);
        drag.addEventListener("touchend", mup, false);
        document.body.addEventListener("touchleave", mup, false);
    }

    //マウスボタンが上がったら発火
    function mup(e) {
        //スクロール再開
        var drag = document.getElementsByClassName("drag")[0];
        $window.off('.noScroll');
        //ムーブベントハンドラの消去
        document.body.removeEventListener("mousemove", mmove, false);
        if(drag){
            var b = Number(drag.style.top.slice( 0, -2 ));
            if(b < screen.height/2.5){
                drag.addEventListener('click', function() {
                    drag.removeEventListener("mouseup", mup, false);
                    }, false);
                document.body.removeEventListener("touchmove", mmove, false);
                drag.addEventListener('click', function() {
                    drag.removeEventListener("touchend", mup, false);
                }, false);
                //クラス名 .drag も消す
                drag.classList.remove("drag");
                return_scroll();
            } else {
                drag.style.top = "0px";
                drag.style.left = "0px";
            }
        };     
    }
    function no_scroll(){
		//PC用
		var scroll_event = 'onwheel' in document ? 'wheel' : 'onmousewheel' in document ? 'mousewheel' : 'DOMMouseScroll';
		$(document).on(scroll_event,function(e){e.preventDefault();});
		//SP用
		$(document).on('touchmove.noScroll', function(e) {e.preventDefault();});
	}
    //スクロール復活用関数
    function return_scroll(){
        //PC用
        var scroll_event = 'onwheel' in document ? 'wheel' : 'onmousewheel' in document ? 'mousewheel' : 'DOMMouseScroll';
        $(document).off(scroll_event);
        //SP用
        $(document).off('.noScroll');
    }
 
  })();



$(document).ready(function(){

//--------------------------------------超链虚线 ↓↓↓--------------------------------------------------------------		
		
		$('a,input[type="button"],input[type="submit"]').bind('focus',function(){
					if(this.blur){	this.blur();};
		});


//---------------------------------导航 ↓↓↓---------------------------------------------------------------------

var delayTime = []; 
			jQuery('.menus li').each(function(index) {
				$(this).hover(function() {
					var _self = this;
					//$(_self).find('a.ap').stop().fadeTo(500, 1);
					$(_self).find('a.ap').addClass("ap2"); 
					delayTime[index] = setTimeout(function() {
						$(_self).find('ul:eq(0)').slideDown(300);
					},
					1) 
				},function() {					
					clearTimeout(delayTime[index]);
				//	var pcl=$(this).attr("class");
					//if(pcl!="case"){$(this).find('a.ap').stop().fadeTo(400, 0);};
					//$(this).find('a.ap').stop().fadeTo(400, 0);
					$(this).find('a.ap').removeClass("ap2");
					$('ul', this).slideUp(300);
					
					
				})
			});

//---------------------------------------------------------------------------------------------------



$(".inewsti li:last").mouseover(function(){	
		$(".inewsti").addClass("inewstiov");
		$(".inewsmain ul:first").stop().animate({'top':'-180px'},{queue:false,duration:300});
		$(".inewsmain ul:last").stop().animate({'top':'0px'},{queue:false,duration:300});
		
	
	})
	
	$(".inewsti li:first").mouseover(function(){	
		$(".inewsti").removeClass("inewstiov");
		$(".inewsmain ul:first").stop().animate({'top':'0px'},{queue:false,duration:300});
		$(".inewsmain ul:last").stop().animate({'top':'180px'},{queue:false,duration:300});
		
	
	})



//----------------------------------------------------
});







/*隔行换色  gid 代表ID  gli 代表 标签 如li tr a等 gcolor 偶数行颜色*/
/*window.onload=*/function showtable(gid,gli,gcolor){   

var tablename=document.getElementById(gid);

var li=tablename.getElementsByTagName(gli);

for (var i=0;i<=li.length;i++){

if (i%2==0){

li[i].style.backgroundColor="";

}else li[i].style.backgroundColor=gcolor;

}

}



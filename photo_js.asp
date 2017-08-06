//By  :CM.ivan
//Time:2009-07-12
//For :Shuangfu.
<%
dim show_Num,this_x,this_y,show_x,show_y
    show_Num=request.QueryString("show_Num")
 if show_Num="" or isnumeric(show_Num)=false then show_Num=8
    this_x=2
    this_y=show_Num/2
    show_x=1
    show_y=show_Num
%>

//<><><><> 记录产品的相关信息 <><><><>
var f_title=new Array();
var f_size =new Array();
var f_time =new Array();
var f_pic  =new Array();
var f_note =new Array();
var f_info_stats =0;

//<><><><> 义初始定义位置\状态,左 <><><><>
var f_x  =new Array();     //初始化位置数组 x轴
var f_y  =new Array();     //初始化位置数组 Y轴

var f_x_mix=new Array();
var f_y_mix=new Array();

var f_zt =new Array();    //记录每个对象的状态
var f_boj=new Array();    //定义数组对象
var p_add =10;	          //缓冲变量
var t_add =20;	          //时间变量

// <><><><> 最小化大小 <><><><>
var f_mix_w =100;
var f_mix_h =80;
// <><><><> 正常大小 <><><><>
var f_w =128/2+11;
var f_h =152/2+12;
// <><><><> 最大化位置 <><><><>
var f_max_x =162;
var f_max_y =0;
// <><><><> 最大化大小 <><><><>
var f_max_w =440;
var f_max_h =547;

var f_onshow=0;      //记录当前以大图方式展示的图片id值
var f_onshowing=0;   //记录是否正在放大缩小

var x,y,num



    num=0;
for (num=1;num<=<%=show_Num%>;num++){
	f_boj[num]=$("mk_pic_item_"+num+"");
	
//临时记录产品的相关信息
f_title[num] ="产品标题"+num;
f_size [num] ="尺寸"+num;
f_time [num] ="2010-03-10";
f_pic  [num] ="images/"+num+".gif";
f_note [num] ="相关描述"+num+"(当前为页面详细页)";
	
	}
	

//初始化，图片之间的间隔
var root_x;
var root_y;
    root_x=160;
    root_y=184;

    num=0;
for (x=0;x<<%=this_x%>;x++){
for (y=0;y<<%=this_y/2%>;y++){
    num++;
	f_x[num]=root_x*x;
    f_y[num]=root_y*y;
	f_zt[num] =0;
//-----------------------------------
    f_boj[num].style.left=reSize(f_x[num]);
    f_boj[num].style.top =reSize(f_y[num]);

}}	

//<><><><> 义初始定义位置\状态,右 <><><><>
    num=<%=this_y%>;
for (x=0;x<<%=this_x%>;x++){
for (y=0;y<<%=this_y/2%>;y++){
    num++;
	f_x[num]=root_x*(x+2);
    f_y[num]=root_y*y;
	f_zt[num] =0;
//-----------------------------------
    f_boj[num].style.left=reSize(f_x[num]);
    f_boj[num].style.top =reSize(f_y[num]);
}}


// <><><><> 最小化位置,左 <><><><>
    num=0;
for (y=0;y<<%=this_y%>;y++){
for (x=0;x<<%=show_x%>;x++){
    num++;
	f_x_mix[num]=root_x*x;
    f_y_mix[num]=parseInt(root_y/2)*y;
}}	

// <><><><> 最小化位置,右 <><><><>
    num=<%=this_y%>;
for (y=0;y<<%=this_y%>;y++){
for (x=0;x<<%=show_x%>;x++){
    num++;
	f_x_mix[num]=root_x*x+parseInt(root_x/2);
    f_y_mix[num]=parseInt(root_y/2)*y;
}}	


//根据浏览器返回值
function reSize(num){
if(navigator.appName.indexOf("Explorer") > -1){
//IE
  return parseInt(num);
  }else{
//Firefox对"px"要求严格
  return parseInt(num) + "px";
  }
}

//返回对象
function $(obj){
 return document.getElementById(obj);
}

//展示图片
function isshow(id){
	//屏蔽flash
  $("mk_flash_box").style.display="none";
  $("mk_pic_info").style.display="none";
if (getstats(0)){
  for (num=1;num<=<%=show_Num%>;num++)f_zt[num]=1;
  showpic(id);
  showbanner(id);
}
else if(getstats(2)&&f_onshowing==0){
  showbanner(id);
  }
}

//关闭图片
function isclose(id){
if(getstats(2))for (num=1;num<=<%=show_Num%>;num++)f_zt[num]=1;
	//屏蔽flash
  $("mk_flash_box").style.display="none";
  $("mk_pic_info").style.display="none";
  $("mk_pic_show").style.display="none";
  bannerclose(f_onshow);
  showlist(id);
}





//返回状态的
function showstats(sNum){
	var stats;
	    stats=true;
  for (num=1;num<=<%=show_Num%>;num++){
      if(f_zt[num]==sNum){
	    stats=false;
		  }
  }
  return stats;
}

//返回状态的
function getstats(sNum){
	var stats;
	    stats=true;
  for (num=1;num<=<%=show_Num%>;num++){
      if(f_zt[num]!=sNum){
	    stats=false;
		  }
  }
  return stats;
}








function showbanner(id){
//------------------------------------------------------------------
    $("mk_pic_show").style.display="none";
var mk_pic=$("mk_pic_item_"+id+"");
var bannerBox1=$("mk_pic_item_border_1");
    bannerBox1.style.left  =reSize(mk_pic.style.left);
    bannerBox1.style.top   =reSize(mk_pic.style.top);
    bannerBox1.style.width =reSize(mk_pic.offsetWidth);
    bannerBox1.style.height=reSize(mk_pic.offsetHeight);
    bannerBox1.style.display="block";

//------------------------------------------------------------------
var bannerBox2=$("mk_pic_item_border_2");
    bannerBox2.style.left  =reSize(f_max_x);
    bannerBox2.style.top   =reSize(f_max_y);
    bannerBox2.style.width =reSize(f_max_w);
    bannerBox2.style.height=reSize(f_max_h);

	bannerclose(f_onshow)  //关闭当前展示id
	f_onshow=id;           //记录新的展示id值
    bannershow(f_onshow);
	}




//基本状态到幻灯片状态
function showpic(id){
  for (num=1;num<=<%=show_Num%>;num++){
	  if (f_zt[num]!=0){
//------------------------------------
      var temp_x=parseInt(f_boj[num].style.left);
      var temp_y=parseInt(f_boj[num].style.top);
	  if(Math.abs(temp_x-f_x_mix[num])<p_add){
		f_boj[num].style.left = reSize(f_x_mix[num]);
        f_boj[num].style.zoom = 0.5;
		}else{
		f_boj[num].style.left=reSize(temp_x-Math.ceil((temp_x-f_x_mix[num])/p_add));
        f_boj[num].style.zoom = 0.5;
			}
			
	  if(Math.abs(temp_y-f_y_mix[num])<p_add){
		f_boj[num].style.top  = reSize(f_y_mix[num]);
        f_boj[num].style.zoom = 0.5;
		}else{
		f_boj[num].style.top  = reSize(temp_y-Math.ceil((temp_y-f_y_mix[num])/p_add));	
        f_boj[num].style.zoom = 0.5;
			}


	  if((Math.abs(parseInt(f_boj[num].style.left)-f_x_mix[num])==0)&&Math.abs(parseInt(f_boj[num].style.top)-f_y_mix[num])==0){
		f_zt[num]=2;
		}
	}
//------------------------------------
  }
  
   if (!showstats(1)){
	   setTimeout('showpic()',t_add)
   }
	
} 



//还原状态状态到列表展示
function showlist(id){
for (num=1;num<=<%=show_Num%>;num++){
	  if (f_zt[num]!=0){
//------------------------------------
      var temp_x=parseInt(f_boj[num].style.left);
      var temp_y=parseInt(f_boj[num].style.top);
	  
	  if(Math.abs(temp_x-f_x[num])<p_add){
		f_boj[num].style.left=reSize(f_x[num]);
        f_boj[num].style.zoom = 1;
		}else{
		f_boj[num].style.left=reSize(temp_x-Math.ceil((temp_x-f_x[num])/p_add));
        f_boj[num].style.zoom = 1;	
			}
	  if(Math.abs(temp_y-f_y[num])<p_add){
		f_boj[num].style.top=reSize(f_y[num]);
        f_boj[num].style.zoom = 1;
		}else{
		f_boj[num].style.top=reSize(temp_y-Math.ceil((temp_y-f_y[num])/p_add));	
        f_boj[num].style.zoom = 1;
			}

	  if((Math.abs(parseInt(f_boj[num].style.left)-f_x[num])==0)&&Math.abs(parseInt(f_boj[num].style.top)-f_y[num])==0){
		f_zt[num]=0;
		}
	}
//------------------------------------
  }
  
   if (!showstats(1)){
	   setTimeout('showlist()',t_add)
   }else{
	   $("mk_pic_item_border_1").style.display="none";
	   $("mk_pic_item_border_2").style.display="none";
	   }
}






//边框打开
function bannershow(id){
   var thisObj=$("mk_pic_item_border_1");
   var temp_x=parseInt(thisObj.style.left);
   var temp_y=parseInt(thisObj.style.top);
   var temp_w=parseInt(thisObj.style.width);
   var temp_h=parseInt(thisObj.style.height);

if(Math.abs(temp_x-f_max_x)<p_add){
		thisObj.style.left=reSize(f_max_x);
	   }else{
        thisObj.style.left=reSize(temp_x-Math.ceil((temp_x-f_max_x)/p_add));	
		   }
	   
if(Math.abs(temp_y-f_max_y)<p_add){
		thisObj.style.top=reSize(f_max_y);
	   }else{
        thisObj.style.top=reSize(temp_y-Math.ceil((temp_y-f_max_y)/p_add));	
		   }


if(Math.abs(temp_w-f_max_w)<p_add){
		thisObj.style.width=reSize(f_max_w);
	   }else{
        thisObj.style.width=reSize(temp_w-Math.ceil((temp_w-f_max_w)/p_add));	
		   }
	   
if(Math.abs(temp_h-f_max_h)<p_add){
		thisObj.style.height=reSize(f_max_h);
	   }else{
        thisObj.style.height=reSize(temp_h-Math.ceil((temp_h-f_max_h)/p_add));	
		   }



if((Math.abs(temp_x-f_max_x)!=0)||Math.abs(temp_y-f_max_y)!=0||Math.abs(temp_w-f_max_w)!=0||Math.abs(temp_h-f_max_h)!=0){
   f_onshowing=1;
   setTimeout("bannershow("+id+")",t_add);
}else{
	f_onshowing=0;
	var boxHTML;

  boxHTML="<div class=showbox><table>";
  boxHTML=boxHTML+"<tr><td class='close'><span onclick='javascript:tochang("+id+",0);'>翻页&gt;&gt;</span>&nbsp;<span onclick='javascript:isclose("+id+");'>[close]</span></div></td></tr>";
  boxHTML=boxHTML+"<tr><td><img src='"+$("mk_pic_item_"+id).getElementsByTagName("img")[0].src+"' /></td></tr>";
  boxHTML=boxHTML+"<tr><td class='show_note' valign='top'>简单介绍："+$("mk_pic_item_"+id+"").getElementsByTagName("a")[0].innerText+"</td></tr>";
  boxHTML=boxHTML+"</table></div>";

	    $("mk_pic_show").innerHTML=boxHTML;
	    $("mk_pic_item_"+id+"").style.display="none";
        $("mk_pic_show").style.display="block";
	
var bannerBox1=$("mk_pic_item_border_1");
    bannerBox1.style.left   =reSize(f_x_mix[id]);
    bannerBox1.style.top    =reSize(f_y_mix[id]);
    bannerBox1.style.width  =reSize(f_w);
    bannerBox1.style.height =reSize(f_h);
    bannerBox1.style.display="block";
	//thisObj.style.display ="none";
	}
		
   
}




//边框关闭
function bannerclose(id){
if (f_onshow!=0){

   var thisObj=$("mk_pic_item_border_2");
       thisObj.style.display="block";
   var temp_x=parseInt(thisObj.style.left);
   var temp_y=parseInt(thisObj.style.top);
   var temp_w=parseInt(thisObj.style.width);
   var temp_h=parseInt(thisObj.style.height);
//--------------
   var toObj=$("mk_pic_item_"+id+"");
   var temp_to_x=parseInt(toObj.style.left);
   var temp_to_y=parseInt(toObj.style.top);
   var temp_to_w=f_w;
   var temp_to_h=f_h;

if(Math.abs(temp_x-temp_to_x)<p_add){
		thisObj.style.left=reSize(temp_to_x);
	   }else{
        thisObj.style.left=reSize(temp_x-Math.ceil((temp_x-temp_to_x)/p_add));	
		   }
	   
if(Math.abs(temp_y-temp_to_y)<p_add){
		thisObj.style.top=reSize(temp_to_y);
	   }else{
        thisObj.style.top=reSize(temp_y-Math.ceil((temp_y-temp_to_y)/p_add));	
		   }


if(Math.abs(temp_w-temp_to_w)<p_add){
		thisObj.style.width=reSize(temp_to_w);
	   }else{
        thisObj.style.width=reSize(temp_w-Math.ceil((temp_w-temp_to_w)/p_add));	
		   }
	   
if(Math.abs(temp_h-temp_to_h)<p_add){
		thisObj.style.height=reSize(temp_to_h);
	   }else{
        thisObj.style.height=reSize(temp_h-Math.ceil((temp_h-temp_to_h)/p_add));	
		   }



if((Math.abs(temp_x-temp_to_x)!=0)||Math.abs(temp_y-temp_to_y)!=0||Math.abs(temp_w-temp_to_w)!=0||Math.abs(temp_h-temp_to_h)!=0){
  f_onshowing=1;
  setTimeout("bannerclose("+id+")",t_add);
}else{
	$("mk_pic_item_"+id+"").style.display="block";
	thisObj.style.display="none";
    f_onshowing=0;
	}


}
}






//图片转换效果
function tochang(id,ty){
	    f_info_stats=ty;
		if (ty==0){
	    $("mk_pic_show").style.display="none";
		$("mk_flash_box").style.display="block";
		}else{
		$("mk_pic_show").style.display="none";
		$("mk_pic_info").style.display="none";
		$("mk_flash_box").style.display="block";
			}
		$("flash").gotoframe(2);
		$("flash").play();
	}





function changshow(id){
//---------------------------
if (f_info_stats==0){
var mk_pic_info=$("mk_pic_info");
	mk_pic_info.style.display="block";
	
var boxHTML;

boxHTML="<div class='showbox'>";
boxHTML=boxHTML+"<table width=\"96%\" border=\"0\" cellspacing=\"5\" bgcolor=\"#DBDBDB\">";
boxHTML=boxHTML+"    <tr>";
boxHTML=boxHTML+"      <td width=\"49%\"><table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"1\" bgcolor=\"#666666\">";
boxHTML=boxHTML+"        <tr>";
boxHTML=boxHTML+"          <td height=\"20\" colspan=\"2\" valign=\"top\" bgcolor=\"#C0C0C0\"><span class='closebt' onclick='javascript:isclose("+id+");'>[close]</span></td>";
boxHTML=boxHTML+"        </tr>";
boxHTML=boxHTML+"        <tr>";
boxHTML=boxHTML+"          <td width=\"61%\" valign=\"top\" bgcolor=\"#C0C0C0\"><table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"1\" bgcolor=\"#999999\">";
boxHTML=boxHTML+"            <tr>";
boxHTML=boxHTML+"              <td width=\"25%\" bgcolor=\"#C0C0C0\">名称：</td>";
boxHTML=boxHTML+"              <td width=\"75%\" bgcolor=\"#C0C0C0\">"+f_title[id]+"</td>";
boxHTML=boxHTML+"              </tr>";
boxHTML=boxHTML+"            <tr>";
boxHTML=boxHTML+"              <td bgcolor=\"#C0C0C0\">规格：</td>";
boxHTML=boxHTML+"              <td bgcolor=\"#C0C0C0\">"+f_size[id]+"</td>";
boxHTML=boxHTML+"              </tr>";
boxHTML=boxHTML+"            <tr>";
boxHTML=boxHTML+"              <td bgcolor=\"#C0C0C0\">时间：</td>";
boxHTML=boxHTML+"              <td bgcolor=\"#C0C0C0\">"+f_time[id]+"</td>";
boxHTML=boxHTML+"              </tr>";
boxHTML=boxHTML+"            <tr>";
boxHTML=boxHTML+"              <td bgcolor=\"#C0C0C0\">描述：</td>";
boxHTML=boxHTML+"              <td bgcolor=\"#C0C0C0\">"+f_note[id]+"</td>";
boxHTML=boxHTML+"            </tr>";
boxHTML=boxHTML+"          </table></td>";
boxHTML=boxHTML+"          <td width=\"39%\" valign=\"top\" bgcolor=\"#C0C0C0\"><img src=\""+f_pic[id]+"\" width=\"128\" height=\"128\" /></td>";
boxHTML=boxHTML+"        </tr>";
boxHTML=boxHTML+"        <tr>";
boxHTML=boxHTML+"          <td height=\"20\" colspan=\"2\" valign=\"top\" bgcolor=\"#C0C0C0\"><span class='closebt' onclick='javascript:tochang("+id+",1);'>翻页&gt;&gt;</span></td>";
boxHTML=boxHTML+"        </tr>";
boxHTML=boxHTML+"      </table></td>";
boxHTML=boxHTML+"      </tr>";
boxHTML=boxHTML+"  </table></div>";
	

	mk_pic_info.innerHTML=boxHTML;
	$("mk_flash_box").style.display="none";
//---------------------------
}else{
	    $("mk_pic_show").style.display="block";
		$("mk_flash_box").style.display="none";
	}
	
	
	}


// JavaScript Document
// JavaScript Document
//�������ʾһ���㣬�ò���ڿ�Ϊdiv2������ 
function showdiv(divname){ 
var div3 = document.getElementById(divname); //��Ҫ�����Ĳ� 
div3.style.display="block"; //div3��ʼ״̬�ǲ��ɼ��ģ����ÿ�Ϊ�ɼ� 
//window.event�����¼�״̬�����¼�������Ԫ�أ�����״̬�����λ�ú���갴ť״. 
//clientX�����ָ��λ������ڴ��ڿͻ������ x ���꣬���пͻ����򲻰�����������Ŀؼ��͹������� 
div3.style.left=event.clientX+10; //���Ŀǰ��X���ϵ�λ�ã���10��Ϊ�����ұ��ƶ�10��px���㿴������ 
div3.style.top=event.clientY+5; 
div3.style.position="absolute"; //����ָ��������ԣ�����div3���޷�������궯 
//var div2 =document.getElementById('div2'); 
//div3.innerText=div2.innerHTML; 
} 
//�رղ�div3����ʾ 
function closediv(divname){ 
var div3 = document.getElementById(divname); 
div3.style.display="none"; 
}

function winshow(pagename,w,h){
  window.open(pagename,null,"width="+w+",height="+h);
}
function checkbox(obj,num){
  var id;
  for (i=1;i<=num;i++){
	id=obj+i;
	if(document.getElementById(id).checked==""){
	  document.getElementById(id).checked="checked";
	}
	else{
	  document.getElementById(id).checked="";
	}
  }
}



var xhr;
function getXHR() {
 try {
  xhr=new ActiveXObject("Msxml2.XMLHTTP");
 } catch (e) {
 try {
  xhr=new ActiveXObject("Microsoft.XMLHTTP");
 } catch (e) {
 xhr=false;
}
}
if(!xhr&&typeof XMLHttpRequest!='undefined') {
xhr=new XMLHttpRequest();
}
return xhr;
}

function openXHR(method,url) {
getXHR();
xhr.open(method,url,true);
xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
xhr.onreadystatechange=function() {
if(xhr.readyState!=4)return;
  document.write(xhr.responseText);//responseBody
}
xhr.send(null);
}

function loadXML(method,url) {
getXHR();
xhr.open(method,url,true);
xhr.setRequestHeader("Content-Type","text/xml");
xhr.setRequestHeader("Content-Type","GBK");
xhr.onreadystatechange=function() {
if(xhr.readyState!=4) return;
   document.write(xhr.responseText);
  //$("abc").innerHTML = xhr.responseText;
}
xhr.send(null);
}

function $(idValue)
{
  return document.getElementById(idValue);
}

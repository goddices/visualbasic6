<!--
function CheckContext() {
  if (form1.message.value == "") {
    alert("Please input message");
    form1.message.focus();
    return (false);
  }
}

function IsSure(URL) {
  var RetVal = confirm("Are you sure to delete?");
  if (RetVal) { window.location.href=URL;}
}

function Reply() {
	window.location.href="#REPLY";
	form1.user.focus();
}

function Top() {
	window.location.href="#TOP";
}

function Download(Page) {
	var xmlHttp = Ajax_GetXMLHttpRequest();
	xmlHttp.open("GET", Page, true);
	xmlHttp.send("");
}

//Ajax Start
function Ajax_GetXMLHttpRequest() {
	if (window.ActiveXObject) {
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	else if (window.XMLHttpRequest) {
		return new XMLHttpRequest();
	}
}
function GetPage(Page, oComment) {
	var xmlHttp = Ajax_GetXMLHttpRequest();
	xmlHttp.open("GET", Page, true); 
	xmlHttp.onreadystatechange = function() {
		if (xmlHttp.readyState < 4) { 
		    oComment.innerHTML="<img src='../loading.gif'>"; 
		} else if (xmlHttp.readyState == 4) { 
			var response = xmlHttp.responseText; 
			oComment.innerHTML=response; 
		} 
	}  
	xmlHttp.send("");
}

function PostPage(AddPage, oComment) {
	var xmlHttp = Ajax_GetXMLHttpRequest();
    var user = form1.user.value; //escape(form1.user.value);
	var msg = form1.msg.value; //escape(form1.msg.value);
	if (msg == "") {alert("Please input message");form1.msg.focus();return (false);}

	xmlHttp.open("POST", AddPage, true);
	xmlHttp.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
    var SendData = 'user='+encodeURIComponent(user)+'&msg='+encodeURIComponent(msg);
	xmlHttp.onreadystatechange=function() { 
		if (xmlHttp.readyState==2) {
			oComment.innerHTML="<img src='../loading.gif'>";
		} else if (xmlHttp.readyState==4)  {   
			oComment.innerHTML=xmlHttp.responseText;
		}
	}
    
	xmlHttp.send(SendData);
	document.getElementById('msg').value="";
	return false;
}

function DeletePage(Page, oComment) {
    var xmlHttp = Ajax_GetXMLHttpRequest();
	xmlHttp.open("GET", Page, true);
	xmlHttp.send("");
	xmlHttp.onreadystatechange = function() {
		if (xmlHttp.readyState < 4) { 
		    oComment.innerHTML="<img src='../loading.gif'>";
		} else if (xmlHttp.readyState == 4) { 
			var response = xmlHttp.responseText; 
			oComment.innerHTML=response; 
		} 
	}
}

// Toogle support fucntions
function Toggle(secid)
{
	var sectionId = document.getElementById(secid);
	if (sectionId == null) return;
	if (sectionId.style.display == '') {
		sectionId.style.display = 'none';
		var ImgSrc = document.getElementById("i" + secid);
		ImgSrc.src = "../plus.gif";

	} else {
		sectionId.style.display = '';
		var ImgSrc = document.getElementById("i" + secid);
		ImgSrc.src = "../minus.gif";
		form1.user.focus();
	}
}

function NewPost(secid)
{
	window.location.href="#REPLY";
	var sectionId = document.getElementById(secid);
	if (sectionId == null) return;
	sectionId.style.display = '';
	var ImgSrc = document.getElementById("i" + secid);
	ImgSrc.src = "../minus.gif";
	form1.user.focus();
}

function Title(id, str)
{
	var isPF = (typeof(IsPrinterFriendly) != "undefined");
	document.write('<a href="javascript:Toggle(\'' + id + '\')"><img width="9" height="9" border="0" id="i' + id + '" src="'+(isPF?'../minus':'../plus')+'.gif"/></a> ');
	document.write('<a href="javascript:Toggle(\'' + id + '\')" style="text-decoration:none;color=#000000">'+str+'</a>');
}

function CheckHide(id)
{
	var isPF = (typeof(IsPrinterFriendly) != "undefined");
	if(!isPF){
		var oDiv = document.getElementById(id);
		if(oDiv != null) oDiv.style.display = "none";
	}
}

function ParseView(str) {
	//var strOut = str.replace(/(http:\/\/([\w.]+\/?)\S*)/gi, "<a href='$1'>$1</a>");
	var strOut = str.replace(/(<*>)/gi, "$1");
	document.write(strOut);
}

function CheckBBSContext() {
  if (form1.title.value == "") {
    alert("Please input title!");
    form1.title.focus();
    return (false);
  }
  
  if (form1.random.value == ""){
    alert("Please input validate code!");
    form1.random.focus();
    return (false);
  }
  
  if (form1.message.value == "") {
    alert("Please input message!");
    form1.message.focus();
    return (false);
  }
}

function AddToFav(value,title){
	window.external.AddFavorite(value,title);
}

function SetHomePage(){
	this.homepage.style.behavior='url(#default#homepage)';this.homepage.sethomepage('http://www.xuyibo.org/');
}

function ResizeImage(objImage,maxWidth) {
try{
  if(maxWidth>0){
   var objImg = $(objImage);
   if(objImg.width()>maxWidth){
    objImg.width(maxWidth).css("cursor","pointer").click(function(){
     try{showModelessDialog(objImage.src);}catch(e){window.open(objImage.src);}
    });
   }
  }
}catch(e){};
}
//-->
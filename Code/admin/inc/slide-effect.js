/**********************************************/
/*  Purpose:        JavaScript for slide      */
/*  Author:         Foolin                    */
/*  E-mail:         Foolin@126.com            */
/*  Create on:      2009-9-19 15:50:17        */
/**********************************************/


//图片边框效果 : 在<img>中加入class='img'即可。
function imgEffect(){
		var nImg = document.getElementsByTagName("img");
		for(var i = 0; i < nImg.length && nImg[i].className == 'img'; i++){
				nImg[i].onmouseover = function(){ this.style.background = '#9dd96b'}    
				nImg[i].onmouseout = function(){ this.style.background = '#FFF';} 
		}
}

//li滑动效果 : 在<ul>中加入class='li-slide'即可。
function liEffect(){
	var nUl = document.getElementsByTagName("ul");
	for(var i = 0; i < nUl.length; i++){
			if(nUl[i].className != 'li-slide') 	continue;
			var nLi = nUl[i].childNodes;
			for(var j = 0; j < nLi.length; j++){
					nLi[j].onmouseover = function(){ this.style.background = '#b5e28d'}    
					nLi[j].onmouseout = function(){ this.style.background = '#FFF';}
			}
	}
}

//tr滑动效果 : 在<table>中加入class='form'即可。
function trEffect(){
	var nTable = document.getElementsByTagName("table");
	for(var i = 0; i < nTable.length; i++){
			if( nTable[i].className != 'form') 	continue;
			var nTr =  nTable[i].childNodes[0].childNodes; //过滤掉<tbody>
			for(var j = 0; j < nTr.length; j++){
					nTr[j].onmouseover = function(){ this.style.background = '#51C7FF';}    
					nTr[j].onmouseout = function(){ this.style.background = '#F0F8FF';}
			}
	}
}

//初始化效果
function initEffect(){
		imgEffect();
		liEffect();
		trEffect();
}

window.onload = initEffect; //网页载入运行
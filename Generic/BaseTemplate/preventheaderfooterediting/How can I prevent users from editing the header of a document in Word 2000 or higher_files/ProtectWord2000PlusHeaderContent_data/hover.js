image1=new Image()
image1.src="Images/OtherPics/TermsofUseSlice2Bottom.gif"
image2=new Image()
image2.src="Images/OtherPics/TermsofUseSlice4Bottom.gif"

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

var timeoutID;
function StatusText(TextToDisplay){
	clearTimeout(timeoutID)
	window.status = TextToDisplay
	timeoutID = setTimeout("ClearStatus()",4000)
}
function ClearStatus(){
	window.status = ""
}

function LandscapeWindow(URLRef){
var WindowInfo=window.open(URLRef,"","left=90,top=10,width=618,height=471");
}
function PortraitWindow(URLRef){
var WindowInfo=window.open(URLRef,"","left=160,top=0,width=430,height=573");
}

var timerID = null;
// avoid error of passing event object in older browsers
if (!document.layers && !document.all) event = "none";

function showToolTip(sTipText, sBGColor, sFontFace, sSize, e) {
    if (document.layers) { // Netscape 4.7x
        var ToolTip = document.layers.tooltip;
        var ToolTipDoc = ToolTip.document;
        ToolTipDoc.open();
        ToolTipDoc.write('<body><div id="tt" style="position:absolute;border:solid;border-width:1;border-color:#000000;layer-background-color:' + sBGColor + ';">');
        ToolTipDoc.write('<font face="' + sFontFace + '" size="' + sSize + '">' + sTipText);
        ToolTipDoc.write('</font></div></body>');
        ToolTipDoc.close();
        ToolTip.clip.width = ToolTipDoc.layers.tt.clip.width;
        ToolTip.clip.height = ToolTipDoc.layers.tt.clip.height;
        document.captureEvents(Event.MOUSEMOVE);
        document.onmousemove = delayToolTip;
        delayToolTip(e);
     }
     if (document.all) { // IE
        var ToolTip = document.all.tooltip;
        var sInnerHTML = '<div id="tt" style="position:absolute;border:solid;border-width:1;padding:2;border-color:#000000;background-color:' + sBGColor + ';">';
        sInnerHTML += '<font face="' + sFontFace + '" size="' + sSize + '">' + sTipText + '</font></div>';
        ToolTip.innerHTML = sInnerHTML;
        ToolTip.noWrap = true;
        document.onmousemove = delayToolTip;
        delayToolTip(e);
     }
} // showToolTip()

function delayToolTip(e) {
    if (timerID) clearTimeout(timerID);
    var nX = 0;
    var nY = 0;
    if (document.layers) { // Netscape 4.7x
        nX = e.pageX;
        nY = e.pageY;
    }
    if (document.all) { // IE
        nX = window.event.clientX + document.body.scrollLeft;
        nY = window.event.clientY + document.body.scrollTop;
    }
    timerID = setTimeout("displayToolTip(" + nX + ", " + nY + ")", 300);
} // delayToolTip()

function displayToolTip(nX, nY) {
    if (document.layers) { // Netscape 4.7x
        document.releaseEvents(Event.MOUSEMOVE);
        document.onmousemove = null;
        var ToolTip = document.layers.tooltip;
        var ToolTipClip = ToolTip.clip;
        var nRightMargin = window.innerWidth + window.pageXOffset;
        if ((nX + ToolTipClip.width) > (nRightMargin - 10)) { // if the tooltip runs off right
            nX = nRightMargin - ToolTipClip.width - 20;
        }
        if ((nY + ToolTipClip.height) > (window.innerHeight + window.pageYOffset - 38)) { // if tooltip runs off bottom
            nY -= ToolTipClip.height + 22;
        }
        ToolTip.left = nX;
        ToolTip.top = nY + 20;
        ToolTip.visibility = 'visible';
     }
     if (document.all) { // IE
        document.onmousemove = null;
        var ToolTip = document.all.tooltip;
        var ToolTipStyle = ToolTip.style;
        var InnerToolTip = ToolTip.all.tt;
        var nRightMargin = document.body.clientWidth + document.body.scrollLeft;
        if ((nX + InnerToolTip.clientWidth) > nRightMargin) { // if the tooltip runs off right
            nX = nRightMargin - InnerToolTip.clientWidth - 6;
        }
        if ((nY + InnerToolTip.clientHeight) > (document.body.clientHeight + document.body.scrollTop - 20)) { // if tooltip runs off bottom
            nY -= InnerToolTip.clientHeight + 30;
        }
        ToolTipStyle.pixelLeft = nX;
        ToolTipStyle.pixelTop = nY + 20;
        ToolTipStyle.visibility = 'visible';
     }
} // displayToolTip()

function hideToolTip() {
    if (timerID) clearTimeout(timerID);
    if (document.layers) { // Netscape 4.7x
        document.releaseEvents(Event.MOUSEMOVE);
        document.onmousemove = null;
        document.layers.tooltip.visibility = 'hidden';
    }
    if (document.all) { // IE
     document.onmousemove = null;
     var ToolTipStyle = document.all.tooltip.style;
     ToolTipStyle.visibility = 'hidden';
     ToolTipStyle.pixelLeft = 0; // move layer to top left corner, otherwise when you
     ToolTipStyle.pixelTop = 0;  // set innerHTML again, it may cause a horizontal scrollbar
    }
} // hideToolTip()

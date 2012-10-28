//
// Common code
//
var ieVer = 0;
var curLang = null;
var showAll = true;
var cook = null;
var baseUrl = document.scripts[document.scripts.length - 1].src.replace(/[^\/]+.js/, "");

if (navigator.appName == "Microsoft Internet Explorer") {
	var ver = navigator.appVersion;
	var v = new Number(ver.substring(0,ver.indexOf(".", 0)));
	if (v >= 4) {
		ieVer = 4;

		// Look for a version number buried somewhere in the version string.
		var toks = ver.split(/[^0-9.]/);
		if (toks) {
			for (var i = 0; i < toks.length; i++) {
				var tok = toks[i];
				if (tok.indexOf(".", 0) > 0) {
					if (tok >= 5)
						ieVer = 5;
				}
			}
		}
	}
}

if (ieVer >= 4) {
	window.onload = bodyOnLoad;
}

function bodyOnClick()
{
	if (ieVer >= 4) {
		var elem = window.event.srcElement;
		for (; elem; elem = elem.parentElement) {
			if (elem.id == "reftip")
				return;
		}

		hideTip();
		closeMenu();
		hideSeeAlso();
	}
}

function bodyOnLoad()
{
	if (ieVer >= 4) {
		initLangs();
		initReftips();
		initSeeAlso();
		document.body.onclick = bodyOnClick;
	}

}

//
// Language filtering
//
function initLangs()
{
	var hdr = document.all.hdr;
	if (!hdr)
		return;

	var langs = new Array;
	var spans = document.all.tags("SPAN");
	if (spans) {
		var iElem = spans.length;
		for (iElem = 0; iElem < spans.length; iElem++) {
			var elem = spans[iElem];
			if (elem.className == "lang") {

				// Update the array of unique language names.
				var a = elem.innerText.split(",");
				for (var iTok = 0; iTok < a.length; iTok++) {
					var m = a[iTok].match(/([A-Za-z].*[A-Za-z+#])/);
					if (m) {
						var iLang = 0;
						while (iLang < langs.length && langs[iLang] < m[1])
							iLang++;
						if (iLang == langs.length || langs[iLang] != m[1]) {
							var before = langs.slice(0,iLang);
							var after = langs.slice(iLang);
							langs = before.concat(m[1]).concat(after);
						}
					}
				}
			}
		}
	}

	if (langs.length > 0) {
		var pres = document.all.tags("PRE");
		if (pres) {
			for (var iPre = 0; iPre < pres.length; iPre++)
				initPreElem(pres[iPre]);
		}

		var obj = document.all.obj_cook;
		if (obj && obj.object) {
			cook = obj;
			if (obj.getValue("lang.all") != "1") {
				var lang = obj.getValue("lang");
				var c = langs.length;
				for (var i = 0; i != c; ++i) {
					if (langs[i] == lang) {
						curLang = langs[i];
						showAll = false;
					}
				}
			}
		}

		var iLim = document.body.all.length;
		var head = null;
		for (var i = 0; i < iLim; i++) {
			var elem = document.body.all[i];
			if (elem.tagName.match(/^(P|PRE|[DOU]L)$/))
				break;
			if (elem.tagName.match(/^H[1-6]$/)) {
				head = elem;
				head.insertAdjacentHTML('BeforeEnd', '<SPAN CLASS=ilang></SPAN>');
			}
		}

		var td = hdr.insertCell(0);
		if (td) {
			// Localizable strings.
			var L_Filter_Tip = "Language Filter";		// tooltip for language button
			var L_Language = "Language";				// heading for menu of programming languages
			var L_Show_All = "Show All";				// label for 'show all languages' menu item

			// Add the language button to the button bar.
			td.className = "button1";
			td.style.width = "19px";
			td.onclick = langMenu;
			td.innerHTML = '<IMG SRC="' + baseUrl + 'Filter.gif' + '" ALT="' +
				L_Filter_Tip + '" BORDER=0>';

			// Add the menu.
			var div = '<DIV ID="lang_menu" CLASS=langMenu><B>' + L_Language + '</B><UL>';
			for (var i = 0; i < langs.length; i++)
				div += '<LI><A HREF="" ONCLICK="chooseLang(this)">' + langs[i] + '</A><BR>';
			div += '<LI><A HREF="" ONCLICK="chooseAll()">' + L_Show_All + '</A></UL></DIV>';
			document.body.insertAdjacentHTML('BeforeEnd', div);
		}

		if (!showAll)
			filterLang();
	}
}

function initPreElem(pre)
{
	var htm0 = pre.outerHTML;

	var reLang = /<span\b[^>]*class="?lang"?[^>]*>/i;
	var iFirst = -1;
	var iSecond = -1;

	iFirst = htm0.search(reLang);
	if (iFirst >= 0) {
		iPos = iFirst + 17;
		iMatch = htm0.substr(iPos).search(reLang);
		if (iMatch >= 0)
			iSecond = iPos + iMatch;
	}

	if (iSecond < 0) {
		var htm1 = trimPreElem(htm0);
		if (htm1 != htm0) {
			pre.insertAdjacentHTML('AfterEnd', htm1);
			pre.outerHTML = "";
		}
	}
	else {
		var rePairs = /<(\w+)\b[^>]*><\/\1>/gi;

		var substr1 = htm0.substring(0,iSecond);
		var tags1 = substr1.replace(/>[^<>]+(<|$)/g, ">$1");
		var open1 = tags1.replace(rePairs, "");
		open1 = open1.replace(rePairs, "");

		var substr2 = htm0.substring(iSecond);
		var tags2 = substr2.replace(/>[^<>]+</g, "><");
		var open2 = tags2.replace(rePairs, "");
		open2 = open2.replace(rePairs, "");

		pre.insertAdjacentHTML('AfterEnd', open1 + substr2);
		pre.insertAdjacentHTML('AfterEnd', trimPreElem(substr1 + open2));
		pre.outerHTML = "";
	}	
}

function trimPreElem(htm)
{
	return htm.replace(/[ \r\n]*((<\/[BI]>)*)[ \r\n]*<\/PRE>/g, "$1</PRE>").replace(
		/\w*<\/SPAN>\w*((<[BI]>)*)\w*\r\n/g, "\r\n</SPAN>$1"
		);
}

function getBlock(elem)
{
	while (elem && elem.tagName.match(/^([BIUA]|SPAN|CODE|TD)$/))
		elem = elem.parentElement;
	return elem;
}

function langMenu()
{
	bodyOnClick();

	window.event.returnValue = false;
	window.event.cancelBubble = true;

	var div = document.all.lang_menu;
	var lnk = window.event.srcElement;
	if (div && lnk) {
		var x = lnk.offsetLeft + lnk.offsetWidth - div.offsetWidth;
		div.style.pixelLeft = (x < 0) ? 0 : x;
		div.style.pixelTop = lnk.offsetTop + lnk.offsetHeight;
		div.style.visibility = "visible";
	}
}

function chooseLang(item)
{
	window.event.returnValue = false;
	window.event.cancelBubble = true;

	if (item) {
		closeMenu();
		curLang = item.innerText;
		showAll = false;
	}

	if (cook) {
		cook.putValue('lang', curLang);
		cook.putValue('lang.all', '');
	}

	filterLang();
}

function chooseAll()
{
	window.event.returnValue = false;
	window.event.cancelBubble = true;

	closeMenu();

	showAll = true;
	if (cook)
		cook.putValue('lang.all', '1');

	unfilterLang();
}

function closeMenu()
{
	var div = document.all.lang_menu;
	if (div && div.style.visibility != "hidden") {
		var lnk = document.activeElement;
		if (lnk && lnk.tagName == "A")
			lnk.blur();

		div.style.visibility = "hidden";
	}
}

function getNext(elem)
{
	for (var i = elem.sourceIndex + 1; i < document.all.length; i++) {
		var next = document.all[i];
		if (!elem.contains(next))
			return next;
	}
	return null;
}

function filterMatch(text, name)
{
	var a = text.split(",");
	for (var iTok = 0; iTok < a.length; iTok++) {
		var m = a[iTok].match(/([A-Za-z].*[A-Za-z+#])/);
		if (m && m[1] == name)
			return true;
	}
	return false;
}

function topicHeading(head)
{
	var iLim = document.body.children.length;
	var idxLim = head.sourceIndex;

	for (var i = 0; i < iLim; i++) {
		var elem = document.body.children[i];
		if (elem.sourceIndex < idxLim) {
			if (elem.tagName.match(/^(P|PRE|[DOU]L)$/))
				return false;
		}
		else
			break;
	}
	return true;
}

function filterLang()
{
	var spans = document.all.tags("SPAN");
	for (var i = 0; i < spans.length; i++) {
		var elem = spans[i];
		if (elem.className == "lang") {
			var newVal = filterMatch(elem.innerText, curLang) ? "block" : "none";
			var block = getBlock(elem);
			block.style.display = newVal;
			elem.style.display = "none";

			if (block.tagName == "DT") {
				var next = getNext(block);
				if (next && next.tagName == "DD")
					next.style.display = newVal;
			}
			else if (block.tagName == "DIV") {
				block.className = "filtered2";
			}
			else if (block.tagName.match(/^H[1-6]$/)) {
				if (topicHeading(block)) {
					if (newVal != "none") {
						var tag = null;
						if (block.children && block.children.length) {
							tag = block.children[block.children.length - 1];
							if (tag.className == "ilang") {
								tag.innerHTML = (newVal == "block") ?
									'&nbsp; [' + curLang + ']' : "";
							}
						}
					}
				}
				else {
					var next = getNext(block);
					while (next && !next.tagName.match(/^(H[1-6]|DIV)$/)) {
						next.style.display = newVal;
						next = getNext(next);
					}
				}
			}
		}
		else if (elem.className == "ilang") {
			elem.innerHTML = '&nbsp; [' + curLang + ']';
		}
	}

	if (ieVer == 4) {
		document.body.style.display = "none";
		document.body.style.display = "block";
	}
}

function unfilterLang(name)
{
	var spans = document.all.tags("SPAN");
	for (var i = 0; i < spans.length; i++) {
		var elem = spans[i];
		if (elem.className == "lang") {
			var block = getBlock(elem);
			block.style.display = "block";
			elem.style.display = "inline";

			if (block.tagName == "DT") {
				var next = getNext(block);
				if (next && next.tagName == "DD")
					next.style.display = "block";
			}
			else if (block.tagName == "DIV") {
				block.className = "filtered";
			}
			else if (block.tagName.match(/^H[1-6]$/)) {
				if (topicHeading(block)) {
					var tag = null;
					if (block.children && block.children.length) {
						tag = block.children[block.children.length - 1];
						if (tag && tag.className == "ilang")
							tag.innerHTML = "";
					}
				}
				else {
					var next = getNext(block);
					while (next && !next.tagName.match(/^(H[1-6]|DIV)$/)) {
						next.style.display = "block";
						next = getNext(next);
					}
				}
			}
		}
		else if (elem.className == "ilang") {
			elem.innerHTML = "";
		}
	}
}


//
// Reftips (parameter popups)
//
function initReftips()
{
	var DLs = document.all.tags("DL");
	var PREs = document.all.tags("PRE");
	if (DLs && PREs) {
		var iDL = 0;
		var iPRE = 0;
		var iSyntax = -1;
		for (var iPRE = 0; iPRE < PREs.length; iPRE++) {
			if (PREs[iPRE].className == "syntax") {
				while (iDL < DLs.length && DLs[iDL].sourceIndex < PREs[iPRE].sourceIndex)
					iDL++;			
				if (iDL < DLs.length) {
					initSyntax(PREs[iPRE], DLs[iDL]);
					iSyntax = iPRE;
				}
				else
					break;
			}
		}

		if (iSyntax >= 0) {
			var last = PREs[iSyntax];
			last.insertAdjacentHTML(
				'AfterEnd',
				'<DIV ID=reftip CLASS=reftip STYLE="position:absolute;visibility:hidden;overflow:visible;"></DIV>'
				);
		}
	}
}

function initSyntax(pre, dl)
{
	var strSyn = pre.outerHTML;
	var ichStart = strSyn.indexOf('>', 0) + 1;
	var terms = dl.children.tags("DT");
	if (terms) {
		for (var iTerm = 0; iTerm < terms.length; iTerm++) {
			var words = terms[iTerm].innerText.replace(/\[.+\]/g, " ").replace(/,/g, " ").split(" ");
			var htm = terms[iTerm].innerHTML;
			for (var iWord = 0; iWord < words.length; iWord++) {
				var word = words[iWord];

				if (word.length > 0 && htm.indexOf(word, 0) < 0)
					word = word.replace(/:.+/, "");

				if (word.length > 0) {
					var ichMatch = findTerm(strSyn, ichStart, word);
					while (ichMatch > 0) {
						var strTag = '<A HREF="" ONCLICK="showTip(this)" CLASS="synParam">' + word + '</A>';

						strSyn =
							strSyn.slice(0, ichMatch) +
							strTag +
							strSyn.slice(ichMatch + word.length);

						ichMatch = findTerm(strSyn, ichMatch + strTag.length, word);
					}
				}
			}
		}
	}

	// Replace the syntax block with our modified version.
	pre.outerHTML = strSyn;
}

function findTerm(strSyn, ichPos, strTerm)
{
	var ichMatch = strSyn.indexOf(strTerm, ichPos);
	while (ichMatch >= 0) {
		var prev = (ichMatch == 0) ? '\0' : strSyn.charAt(ichMatch - 1);
		var next = strSyn.charAt(ichMatch + strTerm.length);
		if (!isalnum(prev) && !isalnum(next) && !isInTag(strSyn, ichMatch)) {
			var ichComment = strSyn.indexOf("/*", ichPos);
			while (ichComment >= 0) {
				if (ichComment > ichMatch) {
					ichComment = -1;
					break; 
				}
				var ichEnd = strSyn.indexOf("*/", ichComment);
				if (ichEnd < 0 || ichEnd > ichMatch)
					break;
				ichComment = strSyn.indexOf("/*", ichEnd);
			}
			if (ichComment < 0) {
				ichComment = strSyn.indexOf("//", ichPos);
				while (ichComment >= 0) {
					if (ichComment > ichMatch) {
						ichComment = -1;
						break; 
					}
					var ichEnd = strSyn.indexOf("\n", ichComment);
					if (ichEnd < 0 || ichEnd > ichMatch)
						break;
					ichComment = strSyn.indexOf("//", ichEnd);
				}
			}
			if (ichComment < 0)
				break;
		}
		ichMatch = strSyn.indexOf(strTerm, ichMatch + strTerm.length);
	}
	return ichMatch;
}

function isInTag(strHtml, ichPos)
{
	return strHtml.lastIndexOf('<', ichPos) >
		strHtml.lastIndexOf('>', ichPos);
}

function isalnum(ch)
{
	return ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9') || (ch == '_'));
}

function showTip(link)
{
	bodyOnClick();

	var tip = document.all.reftip;
	if (!tip || !link)
		return;

	window.event.returnValue = false;
	window.event.cancelBubble = true;

	// Hide the tip if necessary and initialize its size.
	tip.style.visibility = "hidden";
	tip.style.pixelWidth = 260;
	tip.style.pixelHeight = 24;

	// Find the link target.
	var term = null;
	var def = null;
	var DLs = document.all.tags("DL");
	for (var iDL = 0; iDL < DLs.length; iDL++) {
		if (DLs[iDL].sourceIndex > link.sourceIndex) {
			var dl = DLs[iDL];
			var iMax = dl.children.length - 1;
			for (var iElem = 0; iElem < iMax; iElem++) {
				var dt = dl.children[iElem];
				if (dt.tagName == "DT" && dt.style.display != "none") {
					if (findTerm(dt.innerText, 0, link.innerText) >= 0) {
						var dd = dl.children[iElem + 1];
						if (dd.tagName == "DD") {
							term = dt;
							def = dd;
						}
						break;
					}
				}
			}
			break;
		}
	}

	if (def) {
		window.linkElement = link;
		window.linkTarget = term;
		tip.innerHTML = '<DL><DT>' + term.innerHTML + '</DT><DD>' + def.innerHTML + '</DD></DL>';
		window.setTimeout("moveTip()", 0);
	}
}

function jumpParam()
{
	hideTip();

	window.linkTarget.scrollIntoView();
	document.body.scrollLeft = 0;

	flash(3);
}

function flash(c)
{
	window.linkTarget.style.background = (c & 1) ? "#CCCCCC" : "";
	if (c)
		window.setTimeout("flash(" + (c-1) + ")", 200);
}

function moveTip()
{
	var tip = document.all.reftip;
	var link = window.linkElement;
	if (!tip || !link)
		return; //error

	var w = tip.offsetWidth;
	var h = tip.offsetHeight;

	if (w > tip.style.pixelWidth) {
		tip.style.pixelWidth = w;
		window.setTimeout("moveTip()", 0);
		return;
	}

	var maxw = document.body.clientWidth;
	var maxh = document.body.clientHeight;

	if (h > maxh) {
		if (w < maxw) {
			w = w * 3 / 2;
			tip.style.pixelWidth = (w < maxw) ? w : maxw;
			window.setTimeout("moveTip()", 0);
			return;
		}
	}

	var x,y;

	var linkLeft = link.offsetLeft - document.body.scrollLeft;
	var linkRight = linkLeft + link.offsetWidth;

	var linkTop = link.offsetTop - document.body.scrollTop;
	var linkBottom = linkTop + link.offsetHeight;

	var cxMin = link.offsetWidth - 24;
	if (cxMin < 16)
		cxMin = 16;

	if (linkLeft + cxMin + w <= maxw) {
		x = maxw - w;
		if (x > linkRight + 8)
			x = linkRight + 8;
		y = maxh - h;
		if (y > linkTop)
			y = linkTop;
	}
	else if (linkBottom + h <= maxh) {
		x = maxw - w;
		if (x < 0)
			x = 0;
		y = linkBottom;
	}
	else if (w <= linkRight - cxMin) {
		x = linkLeft - w - 8;
		if (x < 0)
			x = 0;
		y = maxh - h;
		if (y > linkTop)
			y = linkTop;
	}
	else if (h <= linkTop) {
		x = maxw - w;
		if (x < 0)
			x = 0;
		y = linkTop - h;
	}
	else if (w >= maxw) {
		x = 0;
		y = linkBottom;
	}
	else {
		w = w * 3 / 2;
		tip.style.pixelWidth = (w < maxw) ? w : maxw;
		window.setTimeout("moveTip()", 0);
		return;
	}

	link.style.background = "#FFFF80";
	tip.style.pixelLeft = x + document.body.scrollLeft;
	tip.style.pixelTop = y + document.body.scrollTop;
	tip.style.visibility = "visible";
}

function hideTip()
{
	if (window.linkElement) {
		window.linkElement.style.background = "";
		window.linkElement = null;
	}

	var tip = document.all.reftip;
	if (tip) {
		tip.style.visibility = "hidden";
		tip.innerHTML = "";
	}
}

function beginsWith(s1, s2)
{
	// Does s1 begin with s2?
	return s1.substring(0, s2.length) == s2;
}

//
// See Also popups
//
function initSeeAlso()
{
	// Localizable strings.
	var L_See_Also = "See Also";
	var L_Requirements = "Requirements";
	var L_QuickInfo = "QuickInfo";

	var hdr = document.all.hdr;
	if (!hdr)
		return;

	var divS = new String;
	var divR = new String;

	var heads = document.all.tags("H4");
	if (heads) {
		for (var i = 0; i < heads.length; i++) {
			var head = heads[i];
			var txt = head.innerText;
			if (beginsWith(txt, L_See_Also)) {
				divS += head.outerHTML;
				var next = getNext(head);
				while (next && !next.tagName.match(/^(H[1-4]|DIV)$/)) {
					divS += next.outerHTML;
					next = getNext(next);
				}
			}
			else if (beginsWith(txt, L_Requirements) || beginsWith(txt, L_QuickInfo)) {
				divR += head.outerHTML;
				var next = getNext(head);
				while (next && !next.tagName.match(/^(H[1-4]|DIV)$/)) {
					divR += next.outerHTML;
					next = getNext(next);
				}
			}
		}
	}

	var pos = getNext(hdr.parentElement);
	if (pos) {
		if (divR != "") {
			divR = '<DIV ID=rpop CLASS=sapop>' + divR + '</DIV>';
			var td = hdr.insertCell(0);
			if (td) {
				td.className = "button1";
				td.style.width = "19px";
				td.onclick = showRequirements;
				td.innerHTML = '<IMG SRC="' + baseUrl + 'Requirements.gif' + '" ALT="' + L_Requirements + '" BORDER=0>';
				if (ieVer == 4)
					document.body.insertAdjacentHTML('AfterBegin', divR);
				else
					document.body.insertAdjacentHTML('BeforeEnd', divR);
			}
		}
		if (divS != "") {
			divS = '<DIV ID=sapop CLASS=sapop>' + divS + '</DIV>';
			var td = hdr.insertCell(0);
			if (td) {
				td.className = "button1";
				td.style.width = "19px";
				td.onclick = showSeeAlso;
				td.innerHTML = '<IMG SRC="' + baseUrl + 'SeeAlso.gif' + '" ALT="' + L_See_Also + '" BORDER=0>';
				if (ieVer == 4)
					document.body.insertAdjacentHTML('AfterBegin', divS);
				else
					document.body.insertAdjacentHTML('BeforeEnd', divS);
			}
		}
	}
}

function showSeeAlso()
{
	bodyOnClick();

	window.event.returnValue = false;
	window.event.cancelBubble = true;

	var div = document.all.sapop;
	var lnk = window.event.srcElement;

	if (div && lnk) {
		div.style.pixelTop = lnk.offsetTop + lnk.offsetHeight;
		div.style.visibility = "visible";
	}
}

function showRequirements()
{
	bodyOnClick();

	window.event.returnValue = false;
	window.event.cancelBubble = true;

	var div = document.all.rpop;
	var lnk = window.event.srcElement;

	if (div && lnk) {
		div.style.pixelTop = lnk.offsetTop + lnk.offsetHeight;
		div.style.visibility = "visible";
	}
}

function hideSeeAlso()
{
	var div = document.all.sapop;
	if (div)
		div.style.visibility = "hidden";

	var div = document.all.rpop;
	if (div)
		div.style.visibility = "hidden";
}


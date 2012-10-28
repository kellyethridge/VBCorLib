writeCSS(scriptPath());

function scriptPath()
{
	var col = document.scripts;
	return col[col.length - 1].src;
}

function writeCSS(spath)
{
	// Get a base CSS name based on the browser.
	var css = "backsdkn.css";
	if (navigator.appName == "Microsoft Internet Explorer") {
		var sVer = navigator.appVersion;
		sVer = sVer.substring(0, sVer.indexOf("."));
		if (sVer >= 4)
			css = "backsdk4.css";
		else
			css = "backsdk3.css";
	}

	// The CSS is in the same directory as the script.
	css = spath.replace(/linkcss.js/, css);
	document.writeln('<LINK REL="stylesheet" HREF="' + css + '">');
}


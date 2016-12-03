var xls_to_json = require("xls-to-json");
var fs = require("fs");
var exec = require('child_process').exec;

var n = 30;

var convert = function(index, data){
	var html_content = '<!DOCTYPE html><html><head><style type="text/css">@font-face {font-family: "Eras Light ITC";font-style: normal;font-weight: normal;src: local("Eras Light ITC"), url("ERASLGHT.woff") format("woff");}.text {position: absolute;font-size: 40px;text-align: center;font-family: "Eras Light ITC";}#name {top: 552px;left: 435px;width: 564px;}#college {top: 620px;left: 191px;width: 873px;}#position {top: 689px;left: 404px;width: 280px;}#event {top: 757px;left: 249px;width: 720px;}</style></head><body><img src="template.jpg"><div id="name" class="text">' + data[index].name + '</div><div id="college" class="text">' + data[index].college + '</div><div id="position" class="text">' + data[index].position + '</div><div id="event" class="text">' + data[index].event + '</div></body></html>';
	fs.writeFile("html_files/" + index%n + ".html", html_content, function(err){
		if(err){
			throw err;
		}
		console.log("HTMLfied: " + data[index].name);
		exec("wkhtmltopdf --zoom 2.22 --orientation Landscape -L 15 " + "html_files/" + index%n + ".html pdf_files/" + data[index].name.replace(/[^a-zA-Z0-9]/g, '') + ".pdf", function(err, stderr, stdout){
			if(err){
				throw err;
			}
			console.log("PDFfied: " + data[index].name);
			if(index+n < data.length) convert(index+n, data);
		});
	});
};

xls_to_json({
	input: "excel_sheet/content.xls",
	output: "content.json"
}, function(err, json){
	if(err) {
		console.log(err);
	} else {
		console.log	(json);
		for(var i=0; i<n&&i<json.length; ++i){
			convert(i, json);
		}
	}
});
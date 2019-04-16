/*
author: flack
date: 2019-03-08
desc: 前端强大的纯JS导出组件Flink.js
version: v1.1.5
*/

;(function(win, doc, global, undefined){
	//开启严格模式
	"use strict";
	var defaults = {
		type: 'excel',  //默认excel,(excel word csv pdf image)
		dataSource: 'json',  //默认json,(json table html)
		sheetName: 'sheet1',  //Worksheet名称
		sheetHeader: [],  //要导出文件的表头,eg:['姓名','年龄','住址']
		showFirstLine: true,  //默认true,是否显示首行数据
		data: {},  //要导出的json格式数据
		element: '',  //标签元素id,当数据类型DataSource=table时,此参数为必填项
		hideRow: false,  //默认false,隐藏行不导出
		hideColumn: false,  //默认false,隐藏列不导出
		isSequence: false,  //默认false,导出文件中是否带序号
		sequenceName: '序号',  //添加的序号名称
		fileName: 'download.xls'
	}

	//构造函数定义一个类传参数
	var fk = function(options){
		if (typeof fk.instance === 'object'){
			return fk.instance;
		}

		this.options = options = _self(this).setOptions(options);
		// this.setOptions.call(this, options);
		this.init();
		fk.instance = this;
		return this;
	}

	/*
	*
	* 私有方法
	*/
	var _self = function (_this){
		/* 
		* =============================================================
		* 私有方法 start
		* =============================================================
		*/
		var $this = _this;  //全局对象
		this.init = function(){
			// 界面渲染or绑定事件
		},

		// 设置基础选项
		this.setOptions = function(opts){
			for(var o in opts){
				defaults[o] = opts[o];
			}

			return defaults;
		},

		this.base64 = function (s){
			// 支持汉字进行解码
			return window.btoa(unescape(encodeURIComponent(s)));
		},

		// Edge
		// Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134
		// Chrome
		// Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Safari/537.36
		// Firefox
		// Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:65.0) Gecko/20100101 Firefox/65.0
		this.getExplorer = function(){
			var explorer = window.navigator.userAgent;
			if(explorer.lastIndexOf('MSIE') >= 0){
				return 'ie';
			}else if(explorer.lastIndexOf('Firefox') >= 0){
				return 'firefox';
			}else if(explorer.lastIndexOf('Chrome') >= 0){
				return 'chrome';
			}else if(explorer.lastIndexOf('Opera') >= 0){
				return 'opera';
			}else if(explorer.lastIndexOf('Safari') >= 0){
				return 'safari';
			}
		},

		// 根据数据源分类处理数据
		this.handlerData = function(opts){
			var content = "";
			switch(opts.dataSource.toUpperCase()){
				case "JSON":
					content = this.handlerJson(opts);
					break;
				case "TABLE":
					content = this.handlerTable(opts);
					break;
				case "HTML":
					content = this.handlerHtml(opts);
					break;
				default:
					content = this.handlerJson(opts);
					break;
			}

			return content;
		},

		// 数据格式是json
		this.handlerJson = function(opts){
			var content = "";
			// 表头
			var header = "";
			if(opts.sheetHeader && opts.sheetHeader.length > 0){
				header = "<tr>";
				for(var i in opts.sheetHeader){
					header += "<td>" + opts.sheetHeader[i] + "</td>";
				}
				header += "</tr>";
			}

			// 内容
			// 是否显示首行数据
			var firstLineIndex = 0;
			if(opts.showFirstLine){
				firstLineIndex = 0;
			}else{
				firstLineIndex = 1;
			}

			var body = "";
			for (var i = firstLineIndex; i < opts.data.length; i++){
				// 输出每行
				// console.log(opts.data[i]);
				body += "<tr>";
				for(var item in opts.data[i]){
					// 输出每行的列
					// 增加\@ \t为了不让表格显示科学计数法或者其他格式
					body += `<td style="mso-number-format:'\@';">${ opts.data[i][item] + '\t'}</td>`;
				}
				body += "</tr>";
			}

			content = "<body><table>" + header + body + "</table></body>";

			return content;
		},

		// 数据格式为table
		this.handlerTable = function(opts){
			var addRow = function(tableObj, content){
				var tr = document.createElement('tr');
				tr.innerHTML = content;
				tableObj.appendChild(tr);
			};

			var deleteRow = function(tr){
				var root = tr.parentNode()
				root.removeChild(tr);
			}

			var content = "";
			var objectTable = document.getElementById(opts.element);
			// 创建输出表格
			var exportTable = document.createElement('table');

			var rowIndex = 1;  //行索引值
			var columnIndex = 0;  //列数起始值
			if(opts.isSequence){
				columnIndex = -1;
			}

			if(opts.hideRow && opts.hideColumn){
				for(var i = 0; i < objectTable.rows.length; i++){
					if(objectTable.rows[i].getAttribute("data-export") 
						&& objectTable.rows[i].getAttribute("data-export").toUpperCase() === "FALSE"){
						// 不导出行
						continue;
					}else{
						//将data-export=true的行添加到导出表格对象中
						//将没有定义data-export属性的行添加到导出表格对象中
						var tr = exportTable.insertRow(exportTable.rows.length);

						for(var j = columnIndex; j < objectTable.rows[i].cells.length; j++){
							if(i == 0 && j == -1){
								// 第一行第一列
								//此列为增加列的名称
								var th = tr.insertCell();
								th.innerHTML = opts.sequenceName;
							}else{
								if(j == -1){
									// 此列为增加列的值，非第一行第一列
									var td0 = tr.insertCell();
									td0.innerHTML = rowIndex;
									rowIndex++;
								}else{
									if(objectTable.rows[i].cells[j].getAttribute("data-export") 
										&& objectTable.rows[i].cells[j].getAttribute("data-export").toUpperCase() === "FALSE"){
										// 不导出列
										continue;
									}else{
										var td = tr.insertCell();
										td.innerHTML = objectTable.rows[i].cells[j].innerHTML;
									}
								}
							}
						}
					}
				}
			}else if(opts.hideRow && !opts.hideColumn){
				for(var i = 0; i < objectTable.rows.length; i++){
					if(objectTable.rows[i].getAttribute("data-export") 
						&& objectTable.rows[i].getAttribute("data-export").toUpperCase() === "FALSE"){
						// 不导出行
						continue;
					}else{
						//将data-export=true的行添加到导出表格对象中
						//将没有定义data-export属性的行添加到导出表格对象中
						// addRow(exportTable, objectTable.rows[i].innerHTML);
						var tr = exportTable.insertRow(exportTable.rows.length);

						for(var j = columnIndex; j < objectTable.rows[i].cells.length; j++){
							if(i == 0 && j == -1){
								// 第一行第一列
								//此列为增加列的名称
								var th = tr.insertCell();
								th.innerHTML = opts.sequenceName;
							}else{
								if(j == -1){
									// 此列为增加列的值，非第一行第一列
									var td0 = tr.insertCell();
									td0.innerHTML = rowIndex;
									rowIndex++;
								}else{
									var td = tr.insertCell();
									td.innerHTML = objectTable.rows[i].cells[j].innerHTML;
								}
							}
						}
					}
				}
			}else if(!opts.hideRow && opts.hideColumn){
				for(var i = 0; i < objectTable.rows.length; i++){
					//将data-export=true的行添加到导出表格对象中
					//将没有定义data-export属性的行添加到导出表格对象中
					var tr = exportTable.insertRow(exportTable.rows.length);

					for(var j = columnIndex; j < objectTable.rows[i].cells.length; j++){
						if(i == 0 && j == -1){
							// 第一行第一列
							//此列为增加列的名称
							var th = tr.insertCell();
							th.innerHTML = opts.sequenceName;
						}else{
							if(j == -1){
								// 此列为增加列的值，非第一行第一列
								var td0 = tr.insertCell();
								td0.innerHTML = rowIndex;
								rowIndex++;
							}else{
								if(objectTable.rows[i].cells[j].getAttribute("data-export") 
									&& objectTable.rows[i].cells[j].getAttribute("data-export").toUpperCase() === "FALSE"){
									// 不导出列
									continue;
								}else{
									var td = tr.insertCell();
									td.innerHTML = objectTable.rows[i].cells[j].innerHTML;
								}
							}
						}
					}
				}
			}else{
				exportTable = objectTable;
			}


			// // 移除行
			// if(opts.hideRow){
			// 	for(var i = 0; i < objectTable.rows.length; i++){
			// 		if(objectTable.rows[i].getAttribute("data-export")){
			// 			// 不导出行
			// 			if(objectTable.rows[i].getAttribute("data-export").toUpperCase() === "FALSE"){
			// 				continue;
			// 			}else{
			// 				//将data-export=true的行添加到导出表格对象中
			// 				// exportTable.insertRow(exportTable.rows.length);
			// 				this.addRow(exportTable, objectTable.rows[i].innerHTML);
			// 			}
			// 		}else{
			// 			//将没有定义data-export属性的行添加到导出表格对象中
			// 			// exportTable.insertRow(exportTable.rows.length);
			// 			this.addRow(exportTable, objectTable.rows[i].innerHTML);
			// 		}
			// 	}
			// }else{
			// 	exportTable = objectTable;
			// }

			// // 移除列
			// if(opts.hideColumn){
			// 	console.log(exportTable);
			// 	for(var i = 0; i < exportTable.rows.length; i++){
			// 		for(var j = 0; j < exportTable.rows[i].cells.length; j++){
			// 			if(exportTable.rows[i].cells[j].getAttribute("data-export")){
			// 				// console.log(exportTable.rows[i].cells[j].innerHTML);
			// 				if(exportTable.rows[i].cells[j].getAttribute("data-export").toUpperCase() === "FALSE"){
			// 					// 移除
			// 					exportTable.rows[i].cells[j].parentNode.removeChild(exportTable.rows[i].cells[j]);
			// 				}
			// 			}
			// 		}
			// 	}
			// }

			content = "<body>" + exportTable.outerHTML + "</body>";
			return content;
		},

		// 数据格式为html
		this.handlerHtml = function(opts){
			//<body><table>${content}</table></body>
			var content = "";
			var body = document.getElementById(opts.element);
			// 这地方的内容要移除掉<link/> <script> <style>
			// 这里还需要解决保持页面样式的问题
			var scriptTag = body.getElementsByTagName("script");

			// console.log(scriptTag);
			// for(var item in scriptTag){
			// 	console.log(scriptTag[item]);
			// 	if(scriptTag[item].getAttribute("data-export") != "false"){
			// 		body.removeChild(scriptTag[item]);
			// 		console.log(scriptTag[item].getAttribute("data-export"));
			// 	}
			// }

			console.log(body.outerHTML);
			
			content = body.outerHTML;

			return content;
		},

		this.excel = function(opts){
			var content = this.handlerData(opts);

			var uri = "";
			var template = "";

			if(this.getExplorer() == "ie"){
				var excelApp = new ActiveXObject("Excel.Application");
				var workBook = excelApp.Workbooks.Add();
				var workSheet = workBook.Worksheets(1);
				workSheet.innerHTML = content;
				excelApp.Visible = true;
				var fileName = excelApp.Application.GetSaveAsFilename(opts.fileName, "Excel Spreadsheets (*.xls),*.xls");
				workBook.SaveAs(fileName);
				workBook.Close(savechanges = false);
				excelApp.Quit();
				excelApp = null;
			}else{
				uri = 'data:application/vnd.ms-excel;base64,';
				//下载的表格模板数据
				template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" 
				xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
				<head><meta charset="UTF-8"><!--[if gte mso 9]><xml>
				<x:ExcelWorkbook>
				<x:ExcelWorksheets>
				<x:ExcelWorksheet>
				<x:Name>${opts.sheetName}</x:Name>
				<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
				</x:ExcelWorksheet>
				</x:ExcelWorksheets>
				</x:ExcelWorkbook>
				</xml><![endif]--></head>${content}</html>`;
				//<body><table>${content}</table></body>
			}
			return {
				uri: uri,
				template: template,
				ext: '.xls'
			};
		},

		this.excelBlob = function(opts){
			var content = this.handlerData(opts);

			var uri = "";

			if(this.getExplorer() == "ie"){
				var excelApp = new ActiveXObject("Excel.Application");
				var workBook = excelApp.Workbooks.Add();
				var workSheet = workBook.Worksheets(1);
				workSheet.innerHTML = content;
				excelApp.Visible = true;
				var fileName = excelApp.Application.GetSaveAsFilename(opts.fileName, "Excel Spreadsheets (*.xls),*.xls");
				workBook.SaveAs(fileName);
				workBook.Close(savechanges = false);
				excelApp.Quit();
				excelApp = null;
			}else{
				//下载的表格模板数据
				var template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" 
				xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
				<head><meta charset="UTF-8"><!--[if gte mso 9]><xml>
				<x:ExcelWorkbook>
				<x:ExcelWorksheets>
				<x:ExcelWorksheet>
				<x:Name>${opts.sheetName}</x:Name>
				<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
				</x:ExcelWorksheet>
				</x:ExcelWorksheets>
				</x:ExcelWorkbook>
				</xml><![endif]--></head>${content}</html>`;
				//<body><table>${content}</table></body>

				var blob = new Blob([template], {type: "application/vnd.ms-excel"});
				uri = URL.createObjectURL(blob);
			}
			return {
				uri: uri,
				ext: '.xls'
			};
		},

		this.word = function(){
			return {
				uri: '',
				template: '',
				ext: '.doc'
			};
		},

		this.csv = function(){
			return {
				uri: '',
				template: '',
				ext: '.csv'
			};
		},

		this.pdf = function(){
			return {
				uri: '',
				template: '',
				ext: '.pdf'
			};
		},

		this.image = function(){
			return {
				uri: '',
				template: '',
				ext: '.png'
			};
		},

		/* 
		* =============================================================
		* 私有方法 end
		* =============================================================
		*/
	}

	//原型链上提供方法
	fk.prototype = {
		version: "1.1.5",
		constructor: fk,		

		// 暴露的接口
		export: function(opts){
			try{
				var $this = this;
				$this.setOptions.call($this, opts);

				var downloadFileName = $this.options.fileName;
				var ext = $this.options.fileName.extension();

				var re = {};
				switch($this.options.type.toUpperCase()){
					case 'EXCEL':
						// re = $this.excel($this.options);
						re = $this.excelBlob($this.options);
						break;
					case 'WORD':
						re = $this.word($this.options);
						break;
					case 'CSV':
						re = $this.csv($this.options);
						break;
					case 'PDF':
						re = $this.pdf($this.options);
						break;
					case 'IMAGE':
						re = $this.image($this.options);
						break;
					default:
						re = $this.excel($this.options);
						break;
				}

				// 非IE，要下载
				if(re.uri != ""){
					if(ext){
						// 传后缀的文件名
					}else{
						// 未传后缀的文件名
						ext = re.ext;
						downloadFileName += ext;
					}

					//下载模板 这个方法浏览器控制台会出现一堆的英文
					// window.location.href = re.uri + $this.base64(re.template);

					var link = document.createElement("a");
					link.style = "display: none;"
					link.href = re.uri; // + $this.base64(re.template);
					// 下载的文件名
					link.download = downloadFileName;
					document.body.appendChild(link);
					link.click();
					document.body.removeChild(link);
				}
			}catch(e){
				console.log(e);
			}
		}
	}

	//兼容CommonJs规范 
	if(typeof module !== 'undefined' && module.exports){
		module.exports = fk;
	}

	//兼容AMD/CMD规范
	if(typeof define === 'function') define(function(){
		return fk;
	});

	//注册全局变量，兼容直接使用script标签引入插件
	//win.fk = fk;
	global.fk = fk;

})(window, document, this);



// 获取文件后缀名
String.prototype.extension = function(){
	var ext = null;
	var name = this.toLowerCase();
	var i = name.lastIndexOf(".");
	if(i > -1){
		ext = name.substring(i);
	}

	return ext;
};


// 判断Array中是否包含某个值
// Array.prototype.containIn = function(obj){
// 	for(var i = 0; i < this.length; i++){
// 		if(this[i] === obj){
// 			return true;
// 		}
// 	}

// 	return false;
// };


// 对象克隆
// Object.prototype.Clone = function(){
// 	var objClone;
// 	if(this.constructor == Object){
// 		objClone = new this.constructor();
// 	}else{
// 		objClone = new this.constructor(this.valueOf());
// 	}

// 	for(var key in this){
// 		if(objClone[key] != this[key]){
// 			if(typeof(this[key]) == 'object'){
// 				objClone[key] = this[key].Clone();
// 			}else{
// 				objClone[key] = this[key];
// 			}
// 		}
// 	}

// 	objClone.toString = this.toString;
// 	objClone.valueOf = this.valueOf;
// 	return objClone;
// }


// 这个方法有问题，死循环了
// var cloneObj = function(obj){
// 	var newObj = {};
// 	if(obj instanceof Array){
// 		newObj = [];
// 	}

// 	for(var key in obj){
// 		var val = obj[key];
// 		// newObj[key] = typeof val === 'object'? cloneObj(val): val;
// 		// newObj[key] = typeof val === 'object'? arguments.callee(val): val;
// 		if(typeof(val) == 'object'){
// 			console.log(1);
// 		}else{
// 			console.log(2);
// 		}
// 	}

// 	return newObj;
// };

require('mm_expand');
const exceljs = require('exceljs');

/**
 * 执行Excel
 */
class Excel {
	/**
	 * 构造函数
	 * @param {Object} config 配置参数
	 */
	constructor(config) {
		var time = new Date();
		this.config = {
			file: './number.xlsx',
			convert: true,
			// extensions: "xlsx",
			sheet: 1, // 'Sheet1'
			excel: {
				// 创建者
				creator: 'MM',
				// 最后修改者
				editor: 'MM',
				// 创建时间
				create_time: time,
				// 修改时间
				edit_time: time,
				// 最后修改时间
				last_time: time
			},
			format: [],
			params: [],
			option: null
		};
		$.push(this.config, config, true);
		// 原始值
		this.original = [];

		// 转换后的值
		this.list = [];

		// excel工作簿
		this.book = null;
	}
}


/**
 * 新建参数模型
 * @param {Object} key
 * @param {Array} arr
 * @return {Object} 返回格式
 */
Excel.prototype.model = function(name, title, type) {
	return {
		name,
		title,
		type
	};
};

/**
 * 新建格式
 * @param {Array} list 列表
 * @param {String} id ID字段
 * @param {String} name 名称字段
 * @param {String} title 标题字段
 * @return {Object} 返回格式模型
 */
Excel.prototype.format = function(list, id, name, title) {
	return {
		list,
		id,
		name,
		title
	};
};

/**
 * 创建表
 * @param {Object} config 配置参数
 */
Excel.prototype.add_book = function(config) {
	if (!config) {
		config = this.config.excel;
	}
	var book = new exceljs.Workbook();
	book.creator = config.creator;
	book.lastModifiedBy = config.editor;
	book.created = config.create_time;
	book.modified = config.edit_time;
	book.lastPrinted = config.last_time;
	this.book = book;
};

/**
 * 读取过滤取值
 * @param {Object} item 对象
 * @return {Object} 返回想要的对象值
 */
Excel.prototype.load_filter = function(item) {
	var fmt = this.config.format;
	fmt.map((m) => {
		var lt = m.list;
		if (m.id) {
			for (var i = 0; i < lt.length; i++) {
				var o = lt[i];
				if (o[m.name] == item[m.key]) {
					item[m.id] = o[m.id];
					delete item[m.key];
				}
			}
		} else {
			for (var i = 0; i < lt.length; i++) {
				var val = lt[i];
				if (val == item[m.key]) {
					item[m.key] = i;
				}
			}
		}
	});
	return item;
};

/**
 * 读取对应列
 */
Excel.prototype.load_col = function(sheet) {
	var col = this.columns();
	if (!sheet) {
		sheet = this.sheet;
	}
	var row = sheet.getRow(1);
	var kv = {};
	var idx = 0;
	var fmt = this.config.format;
	row._cells.forEach((o) => {
		idx++;
		for (var i = 0; i < col.length; i++) {
			var oj = col[i];
			if (o.value == oj.header) {
				kv[idx] = {
					key: oj.key,
					dataType: oj.type || 'varchar'
				};
				continue;
			};
		}
		for (var i = 0; i < fmt.length; i++) {
			var oj = fmt[i];
			if (o.value == oj.title) {
				kv[idx] = {
					key: oj.key,
					dataType: oj.type || 'int'
				};
				continue;
			};
		}
	});
	return kv;
};

/**
 * 保存对应列
 */
Excel.prototype.save_col = function() {
	var col = this.columns();
	var idx = 0;
	var fmt = this.config.format;
	fmt.map((m) => {
		for (var i = 0; i < col.length; i++) {
			var o = col[i];
			if (o.key == m.key) {
				o.header = m.title;
				o.dataType = 'varchar';
			}
		}
	});
	return col;
};


/**
 * 保存过滤取值
 * @param {Object} item 对象
 * @return {Object} 返回想要的对象值
 */
Excel.prototype.save_filter = function(item) {
	var jobj = Object.assign({}, item);
	var fmt = this.config.format;
	fmt.map((m) => {
		if (m.table) {
			for (var k in jobj) {
				if (k == m.key) {
					var lt = m.list;
					for (var i = 0; i < lt.length; i++) {
						var o = lt[i];
						if (o[k] == jobj[k]) {
							jobj[m.key] = o[m.name];
						}
					}
				}
			}
		} else {
			for (var k in jobj) {
				if (k == m.key) {
					if (jobj[k] !== '' && jobj[k] !== undefined && jobj[k] !== null) {
						var index = jobj[k];
						if (index < m.list.length) {
							jobj[k] = m.list[index];
						} else {
							jobj[k] = '';
						}
					} else {
						jobj[k] = '';
					}
				}
			}
		}
	});
	return jobj;
};

/**
 * 读取excel
 * @param {Object} func 加载函数
 * @return {Object} 返回excel对象
 * @return {String|Number} nameOrId 名称或表id
 */
Excel.prototype.load = function(func, nameOrId) {
	var _this = this;
	if (!this.book) {
		this.add_book();
	}
	var cg = this.config;
	// console.log(this.book);
	return new Promise(function(resolve, reject) {
		if (_this.book) {
			var f = cg.file;
			var arr = f.split('.');
			var ext = arr[arr.length - 1];
			if (ext === 'xls') {
				ext = 'xlsx';
			}
			try {
				_this.book[ext].readFile(f.fullname(__dirname)).then(function() {
					var sheet = _this.book.getWorksheet(nameOrId !== undefined ? nameOrId : cg.sheet);
					if (sheet) {
						var list = [];
						var cols = _this.load_col(sheet);
						if (func) {
							sheet.eachRow((o) => {
								var obj = {};
								for (var k in cols) {
									var val = o.getCell(Number(k)).value;
									if (cols[k].dataType.indexOf('int') !== -1 || cols[k].dataType === 'double' || cols[k].dataType ===
										'float') {
										val = Number(val || '0');
									}
									obj[cols[k].key] = val;
								}
								list.push(func(obj));
							});
						} else {
							sheet.eachRow((o) => {
								var obj = {};
								for (var k in cols) {
									var val = o.getCell(Number(k)).value;
									if (cols[k].dataType.indexOf('int') !== -1 || cols[k].dataType === 'double' || cols[k].dataType ===
										'float') {
										val = Number(val || '0');
									}
									obj[cols[k].key] = val;
								}
								list.push(_this.load_filter(obj));
							});
						}
						var lt = list.length > 0 ? list.splice(1) : [];
						resolve(lt);
					} else {
						console.log('worksheet does not exist');
						resolve([]);
					}
				}).catch(function(err) {
					reject(err);
					resolve([]);
				});
			} catch (err) {
				reject(err);
			}
		} else {
			console.log('workbook does not exist');
			resolve([]);
		}
	});
}

/**
 * 设置列
 */
Excel.prototype.columns = function() {
	var columns = [];
	var cg = this.config;
	var lt = cg.params;
	for (var i = 0; i < lt.length; i++) {
		var o = lt[i];
		var key = o.name;
		var fmt = cg.format[key] || {};
		var m = Object.assign({
			key,
			header: o.title,
			type: o.type,
			dataType: o.dataType,
			style: {}
		}, fmt);
		if (o.dataType == 'datetime' || o.dataType == 'timestamp') {
			m.style = {
				numFmt: 'yyyy-dd-mm hh:mm:ss'
			};
		} else if (o.dataType == 'date') {
			m.style = {
				numFmt: 'yyyy-dd-mm'
			};
		} else if (o.dataType == 'time') {
			m.style = {
				numFmt: 'hh:mm:ss'
			};
		}
		m.style.alignment = {
			vertical: 'middle',
			horizontal: 'center'
		};
		columns.push(m);
	}
	return columns;
};

/**
 * 保存
 * @param {Array} jarr 列表
 * @param {Function} func 过滤回调函数
 * @param {String} file 文件路径
 * @return {String} 保存成功返回路径
 */
Excel.prototype.save = async function(jarr, func, file) {
	var f = file || this.config.file;
	var arr = f.split('.');
	f = f.fullname(__dirname);
	var ext = arr[arr.length - 1];
	if (ext == 'xls') {
		ext = 'xlsx';
	}
	try {
		if (jarr) {
			if (!this.book) {
				this.add_book();
			}
			if (this.book) {
				this.sheet = this.book.addWorksheet();
				this.sheet.columns = this.save_col();
				if (func) {
					for (var i = 0; i < jarr.length; i++) {
						var row = this.sheet.addRow(func(jarr[i]));
					}
				} else {
					for (var i = 0; i < jarr.length; i++) {
						var row = this.sheet.addRow(this.save_filter(jarr[i]));
					}
				}

				this.sheet.getRow(1).font = {
					bold: true
				};
				await this.book[ext].writeFile(f, this.config.option);
			} else {
				console.log('workbook does not exist');
			}
		} else if (this.sheet && this.book) {
			await this.book[ext].writeFile(f, this.config.option);
		} else {
			console.error('worksheet does not exist');
		}
	} catch (e) {
		console.error(e);
		f = '';
	}
	return f;
};

/**
 * 键值转换
 * @param {String} prop 属性
 * @param {String} key 键
 * @param {Object} value 值
 * @param {String} name = ['name'] 字段名, 用来取结果
 */
Excel.prototype.convert = function(prop, key, value, name = 'name') {
	var val;
	var list = this.format[prop];
	if (list) {
		for (var i = 0; i < list.length; i++) {
			var o = list[i];
			if (o[key] == value) {
				val = o[name]
				break;
			}
		}
	}
	return val;
};

/**
 * 清理缓存
 */
Excel.prototype.clear = function(){
	this.list = [];
	this.original = [];
	this.book = null;
};

module.exports = Excel;

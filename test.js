const Excel = require('./index.js');

/* 调用示例 */
async function test() {
	var body;
	var config = {
		file: './city.xls',
		convert: true,
		params: [{
				"name": "city_id",
				"title": "城市ID",
				"description": "",
				"type": "number",
				"dataType": "mediumint",
				"number": {
					"range": [
						1,
						8388607
					]
				}
			},
			{
				"name": "show",
				"title": "是否可见",
				"description": "0为仅表单可见，1为仅表单和搜索时可见 ，2为均可见",
				"type": "number",
				"dataType": "smallint",
				"number": {
					"max": 2
				}
			},
			{
				"name": "display",
				"title": "显示顺序",
				"description": "",
				"type": "number",
				"dataType": "smallint",
				"number": {
					"max": 1000
				}
			},
			{
				"name": "province_id",
				"title": "所属省份ID",
				"description": "",
				"type": "number",
				"dataType": "mediumint",
				"number": {
					"range": [
						1,
						8388607
					]
				}
			},
			{
				"name": "name",
				"title": "城市名称",
				"description": "",
				"type": "string",
				"dataType": "varchar",
				"string": {
					"notEmpty": 1
				}
			}
		],
		format: [{
				title: '所属省份',
				id: 'province_id',
				key: 'province',
				name: 'name',
				list: [{
						province_id: 1,
						name: '广东省'
					},
					{
						province_id: 2,
						name: '广西省'
					},
					{
						province_id: 3,
						name: '湖南省'
					}
				]
			},
			{
				title: '是否可见',
				key: 'show',
				list: ['否', '是']
			}
		]
	};
	var excel = new Excel(config);
	var jarr = await excel.load();
	// var jarr = await excel.load((o) => {
	// 	if(o.show == '是' || o.show == '1' || o.show == 'Y'){
	// 		o.show = 1;
	// 	}
	// 	else {
	// 		o.show = 0;
	// 	}
	// 	return o;
	// });

	console.log(jarr);
	var ex = new Excel(config);
	ex.config.file = 'city.xlsx';
	var file = await ex.save(jarr);
	console.log(file);
}

test();

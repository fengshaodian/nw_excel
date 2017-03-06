if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('1 - 副本.xlsx');
var first_sheet_name = workbook.SheetNames[0];
var max_rng = XLSX.utils.decode_range(workbook.Sheets[first_sheet_name]['!ref']);
var max_col = max_rng.e.c;
var max_row = max_rng.e.r;
var table_col_head = [];
var field_col_head = ['sq','user_id','sub_company','company','e_type','e_sub_type','source','status','current_process','operator','user_name','address','contact_name','contact_tel','contact_cellphone','contact_operator','contact_operator_tel','time_division','step_flag','payment','measure_type','reduction_type','original_vol','new_vol','total_vol','e_user_type','vol_level','vol_type','e_from_station','e_from_line','e_from_transform','e_from_sub_transform','apply_time','answer_time','design_time','design_answer_time','mid_time','mid_answer_time','done_time','done_answer_time','finish_time','new_interface','publicity_flag','overtime_flag','file_time','file_name','transfer_flag','design_company','construct_company','apply_check_time','district_flag','temp_flag'];
var roa = [];
var row_data={};


for (var i = 0 ; i <= max_row; i++){
    for (var j = 0 ; j <= max_col; j++)
    {
    	
    	if (i == 0){
    		var col_head = workbook.Sheets[first_sheet_name][XLSX.utils.encode_cell({c:j, r:i})].v;
	    	workbook.Sheets[first_sheet_name][XLSX.utils.encode_cell({c:i, r:0})].v = field_col_head[i];		    	    	
	    	table_col_head.push({field:field_col_head[j],title:col_head});
    	}
    	else{
    		row_data[field_col_head[j]] = workbook.Sheets[first_sheet_name][XLSX.utils.encode_cell({c:j, r:i})].v;  	    	    		
    	}	    	    	
    }
    if (i > 0 ){
    	roa.push(row_data);
    	row_data={};
    }
}


var fs = require("fs");
var file = "test.db";
var exists = fs.existsSync(file);
if(!fs.existsSync(file)){
	console.log("create db");
	fs.openSync(file,"w");
}
var sqlite3 = require("sqlite3").verbose();
var db = new sqlite3.Database(file);
var field_col_head = ['sq','user_id','sub_company','company','e_type','e_sub_type','source','status','current_process','operator','user_name','address','contact_name','contact_tel','contact_cellphone','contact_operator','contact_operator_tel','time_division','step_flag','payment','measure_type','reduction_type','original_vol','new_vol','total_vol','e_user_type','vol_level','vol_type','e_from_station','e_from_line','e_from_transform','e_from_sub_transform','apply_time','answer_time','design_time','design_answer_time','mid_time','mid_answer_time','done_time','done_answer_time','finish_time','new_interface','publicity_flag','overtime_flag','file_time','file_name','transfer_flag','design_company','construct_company','apply_check_time','district_flag','temp_flag'];
db.serialize(function(){
	 if(!exists) {
		 //不存在则创建数据库
	        db.run("CREATE TABLE SQ_TABLE (sq TEXT PRIMARY KEY, user_id TEXT, sub_company TEXT, company TEXT, e_type TEXT, e_sub_type TEXT, source TEXT, status TEXT, current_process TEXT, operator TEXT, user_name TEXT, address TEXT, contact_name TEXT, contact_tel TEXT, contact_cellphone TEXT, contact_operator TEXT, contact_operator_tel TEXT, time_division TEXT, step_flag TEXT, payment TEXT, measure_type TEXT, reduction_type TEXT, original_vol TEXT, new_vol TEXT, total_vol TEXT, e_user_type TEXT, vol_level TEXT, vol_type TEXT, e_from_station TEXT, e_from_line TEXT, e_from_transform TEXT, e_from_sub_transform TEXT, apply_time TEXT, answer_time TEXT, design_time TEXT, design_answer_time TEXT, mid_time TEXT, mid_answer_time TEXT, done_time TEXT, done_answer_time TEXT, finish_time TEXT, new_interface TEXT, publicity_flag TEXT, overtime_flag TEXT, file_time TEXT, file_name TEXT, transfer_flag TEXT, design_company TEXT, construct_company TEXT, apply_check_time TEXT, district_flag TEXT, temp_flag TEXT)");
	    }
	 	
	    var stmt = db.prepare("INSERT INTO SQ_TABLE VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
	    stmt.run(roa.sq, roa.user_id , roa.sub_company , roa.company , roa.e_type , roa.e_sub_type , roa.source , roa.status , roa.current_process , roa.operator , roa.user_name , roa.address , roa.contact_name , roa.contact_tel , roa.contact_cellphone , roa.contact_operator , roa.contact_operator_tel , roa.time_division , roa.step_flag , roa.payment , roa.measure_type , roa.reduction_type , roa.original_vol , roa.new_vol , roa.total_vol , roa.e_user_type , roa.vol_level , roa.vol_type , roa.e_from_station , roa.e_from_line , roa.e_from_transform , roa.e_from_sub_transform , roa.apply_time , roa.answer_time , roa.design_time , roa.design_answer_time , roa.mid_time , roa.mid_answer_time , roa.done_time , roa.done_answer_time , roa.finish_time , roa.new_interface , roa.publicity_flag , roa.overtime_flag , roa.file_time , roa.file_name , roa.transfer_flag , roa.design_company , roa.construct_company , roa.apply_check_time , roa.district_flag , roa.temp_flag );                                   
	    stmt.finalize();                                                   
	    db.close(); 

	   // Insert random data
	    var rnd;
	    for (var i = 0; i < 10; i++) {
	        rnd = Math.floor(Math.random() * 10000000);
	        stmt.run("Thing " + rnd);
	    }

	    stmt.finalize();
	    db.each("SELECT rowid AS id, thing FROM Stuff", function(err, row) {
	        console.log(row.id + ": " + row.thing);
	    });
	
});
//var max_rng = XLSX.utils.decode_range(workbook.Sheets[first_sheet_name]['!ref']);
//var max_col = max_rng.e.c;
//var table_col_head = [];
//var field_col_head = ['sq','user_id','sub_company','company','e_type','e_sub_type','source','status','current_process','operator','user_name','address','contact_name','contact_tel','contact_cellphone','contact_operator','contact_operator_tel','time_division','step_flag','payment','measure_type','reduction_type','original_vol','new_vol','total_vol','e_user_type','vol_level','vol_type','e_from_station','e_from_line','e_from_transform','e_from_sub_transform','apply_time','answer_time','design_time','design_answer_time','mid_time','mid_answer_time','done_time','done_answer_time','finish_time','new_interface','publicity_flag','overtime_flag','file_time','file_name','transfer_flag','design_company','construct_company','apply_check_time','district_flag','temp_flag'];
//
//for (var i = 0 ; i <= max_col; i++)
//{
//	var col_head = workbook.Sheets[first_sheet_name][XLSX.utils.encode_cell({c:i, r:0})].v;
//	workbook.Sheets[first_sheet_name][XLSX.utils.encode_cell({c:i, r:0})].v = field_col_head[i];	
//	table_col_head.push({field:field_col_head[i],title:col_head});
//}
//
//var roa = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name]);

//console.log(roa);

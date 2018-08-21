/**
 * 读出学生 代码信息
 */

// 通过Excel 读取学生代码信息 (读取第一个SHEET的内容)
// domFileId 读取excel信息的domID
// dataArr 存放学生信息的数组
// dataMap 存放学生姓名 与对应 下标的对照关系
function readStudent(domFileId ,dataArr,dataMap,isIE){
    try{
        var rABS = true; //是否将文件读取为二进制字符串
        var file = document.getElementById(domFileId).files[0];
        var reader = new FileReader();
        reader.onload = function(e) {
            try {
                var data = e.target.result;
                if (!rABS) data = new Uint8Array(data);
                var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
                /* DO SOMETHING WITH workbook HERE */
                var obj = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);// 拿到表格对象。默认表格第一行是字段，从第二行开始是数据
                if (obj != null && obj.length > 0) {
                    var name;
                    var code;
                    for (var i = 0; i < obj.length; i++) {
                        code = obj[i]["代码"];
                        name = obj[i]["名称"];
                        dataArr.push(createStudent(name, code));
                        dataMap[name] = dataArr.length - 1;
                    }
                    dataArr.sort(studentCompare);
                    if (dataArr.length > 0) {
                        for (var i = 0; i < dataArr.length; i++) {
                            dataMap[dataArr[i].name] = i;
                        }
                    }
                }
                studentStatus_ = 1;
                $("#stu_status").text("成功");
                $("#stu_count").text(dataArr.length);
                console.log(dataMap);
                console.log("===========");
                console.log(dataArr);
            } catch (e) {
                studentStatus_ = -1;
                $("#stu_status").text("读取失败");
            }
        };
        if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file);
    }catch (e){
        studentStatus_ = -1;
        $("#stu_status").text("读取失败");
    }
}

function studentCompare(x, y){
    var codeX = x.code;
    var codeY = y.code;
    if(codeX == null ||  codeX == ""){
        codeX = "0";
    }
    if(codeY == null ||  codeY == ""){
        codeY = "0";
    }
    if(codeX < codeY){
        return -1;
    }else if(codeX > codeY){
        return 1;
    }else {
        return 0;
    }
}

// 创建学生信息
// name 姓名
// code 编号
function createStudent(name,code){
    var stu = {};
    stu.name = $.trim(name);
    stu.code = code;
    return stu;
}

// 根据学生的姓名 获得学生的编号
// dataArr 存放学生信息的数组
// dataMap 存放学生姓名 与对应 下标的对照关系
// name 学生姓名
function getStudentCodeByName(dataArr, dataMap, name){
    if(dataArr == null || dataMap == null ){
        return "";
    }
    if(name == null || name == ""){
        return "";
    }
    var index = dataMap[name];
    if(index != null && dataArr[index] != null){
        return dataArr[index].code;
    }
    return "";
}
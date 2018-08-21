/**
 * 读出 收据表 信息
 */

// 通过Excel 读取收据信息 (读取第一个SHEET的内容)
// domFileId 读取excel信息的domID
// dataArr 存放收据信息的数组
// dataMap 存放收据 key 与对应 下标的对照关系
// time 符合读取条件的日期格式
function readIncome(domFileId ,dataArr,dataMap,time,isIE){
    try{
        var rABS = true; //是否将文件读取为二进制字符串
        var file = document.getElementById(domFileId).files[0];
        var reader = new FileReader();
        reader.onload = function(e) {
            try {
                var data = e.target.result;
                if (!rABS) data = new Uint8Array(data);
                var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
                var obj = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: "A"});//

                var classTotal = 0;
                var otherTotal = 0;
                var total = 0;
                // 读取符合条件的 收据表信息
                if (obj != null && obj.length > 0) {
                    var timeTest = time.replace(/[^0-9]/ig,"");
                    for (var i = 0; i < obj.length; i++) {
                        if (obj[i].B != null) {
                            var incomeTime = obj[i].B.replace(/[^0-9]/ig,"");
                            if(incomeTime.substring(0, timeTest.length) == timeTest){
                                dataArr.push(createIncome(obj[i].C, obj[i].D, obj[i].I, obj[i].J, obj[i].K, obj[i].L, obj[i].M, obj[i].N, obj[i].O, obj[i].P, obj[i].Q, obj[i].R, obj[i].S, obj[i].T));
                            }
                        }
                    }
                }
                if (dataArr.length > 0) {
                    for (var i = 0; i < dataArr.length; i++) {
                        classTotal = classTotal + dataArr[i].classTotal;
                        otherTotal = otherTotal + dataArr[i].otherTotal;
                    }
                    total = classTotal + otherTotal;
                }

                console.log(dataArr);

                incomeStatus_ = 1;
                $("#inc_status").text("成功");
                $("#inc_count").text(dataArr.length);
                $("#inc_amount1").text(classTotal.toFixed(4));
                $("#inc_amount2").text(otherTotal.toFixed(4));
                $("#inc_amount3").text(total.toFixed(4));
            } catch (e) {
                incomeStatus_ = -1;
                console.log(e);
                $("#inc_status").text("读取失败");
            }
        };
        if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file);
    }catch (e){
        incomeStatus_ = -1;
        console.log(e);
        $("#inc_status").text("读取失败");
    }
}

// 创建学生信息
// no 收据号 C
// name 学生姓名 D
// i11 一对一 新签 I
// i12 一对一 续费 J
// i21 班课 新签 K
// i22 班课 续费 L
// i31 托管 新签 M
// i32 托管 续费 N
// i41 艺术 新签 O
// i42 艺术 续费 P
// i51 其他 活动1 Q
// i54 其他 活动2 R
// i52 其他 车费 S
// i53 其他 资料费 T
function createIncome(no,name,i11,i12,i21,i22,i31,i32,i41,i42,i51,i54,i52,i53){
    var income = {};
    income.no = no;
    income.name = $.trim(name);
    income.i11 = (i11 == null || i11=="")?0:i11*1;
    income.i12 = (i12 == null || i12=="")?0:i12*1;
    income.i21 = (i21 == null || i21=="")?0:i21*1;
    income.i22 = (i22 == null || i22=="")?0:i22*1;
    income.i31 = (i31 == null || i31=="")?0:i31*1;
    income.i32 = (i32 == null || i32=="")?0:i32*1;
    income.i41 = (i41 == null || i41=="")?0:i41*1;
    income.i42 = (i42 == null || i42=="")?0:i42*1;
    income.i51 = (i51 == null || i51=="")?0:i51*1;
    income.i52 = (i52 == null || i52=="")?0:i52*1;
    income.i53 = (i53 == null || i53=="")?0:i53*1;
    income.i54 = (i54 == null || i54=="")?0:i54*1;

    income.classTotal = income.i11+income.i12+income.i21+income.i22+income.i31+income.i32+income.i41+income.i42;
    income.otherTotal = income.i51+income.i52+income.i53+income.i54;
    income.total = income.classTotal + income.otherTotal;

    return income;
}

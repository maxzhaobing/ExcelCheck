/**
 * 读出 课耗表 信息
 */

// 通过Excel 读取收据信息 (读取第一个SHEET的内容)
// domFileId 读取excel信息的domID
// dataArr 存放课耗信息的数组
// dataMap 存放课耗 key 与对应 下标的对照关系
function readClass(domFileId ,dataArr,dataMap){
    try{
        var rABS = true; //是否将文件读取为二进制字符串
        var file = document.getElementById(domFileId).files[0];
        var reader = new FileReader();
        reader.onload = function(e) {
            try {
                var data = e.target.result;
                if (!rABS) data = new Uint8Array(data);
                var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
                var obj = XLSX.utils.sheet_to_json(workbook.Sheets["数据源"]);//

                var total = 0;
                var nobanxing = 0;//没有班型
                var noPriceCount = 0;// 有实际课耗，但是没有价格的记录数
                // 读取符合条件的 收据表信息
                if (obj != null && obj.length > 0) {
                    for (var i = 0; i < obj.length; i++) {
                        var classInfo = createClass(
                            obj[i]["班型"],
                            obj[i]["学生姓名"],
                            obj[i]["单次收费"],
                            obj[i]["实收款"],
                            obj[i]["欠费补交"],
                            obj[i]["第一次续费"],
                            obj[i]["第二次续费"],
                            obj[i]["退费"],
                            obj[i]["实际课耗"]
                        );
                        dataArr.push(classInfo);
                    }
                }
                console.log(dataArr);
                console.log(JSON.stringify(dataArr));
                // 将没有单次收费 信息的数据 进行处理
                if (dataArr.length > 0) {
                    for (var i = 0; i < dataArr.length; i++) {
                        if(dataArr[i].type == null){
                            nobanxing++;
                        }

                        if(dataArr[i].price == null){
                            if(dataArr[i].spend > 0){
                                noPriceCount++;
                            }
                            dataArr[i].price = dataArr[i].total;
                        }
                    }
                }

                if (dataArr.length > 0) {
                    for (var i = 0; i < dataArr.length; i++) {
                        total = total + dataArr[i].total;
                    }
                }
                classStatus_ = 1;

                var info =  "";
                if(nobanxing > 0 || noPriceCount>0){
                    info += "注意：";
                }
                if(nobanxing > 0){
                    info += "存在"+nobanxing+"条记录没有班型信息；";
                }
                if(noPriceCount > 0){
                    info += "存在"+noPriceCount+"条记录有实际课耗但没有单次价格；";
                }
                $("#cla_status").text("成功");
                $("#cla_count").text(dataArr.length);
                $("#cla_amount").text(total.toFixed(4));
                $("#cla_remark").text(info);
            } catch (e) {
                classStatus_ = -1;
                $("#cla_status").text("读取失败");
                console.log(e);
            }
        };
        if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file);
    }catch (e){
        classStatus_ = -1;
        $("#cla_status").text("读取失败");
    }
}

// 创建收据信息
// type 班型
// name 学生姓名
// price 单次收费
// ssk 实收款
// qfbj 欠费补交
// xf1 第一次续费
// xf2 第二次续费
// tf 退费
// spend 实际课耗
function createClass(type,name,price,ssk,qfbj,xf1,xf2,tf,spend){
    var classInfo = {};
    classInfo.type = $.trim(type);;
    classInfo.name = $.trim(name);
    classInfo.price = (price == null || price=="")?null:price*1;
    classInfo.ssk = (ssk == null || ssk=="")?0:ssk*1;
    classInfo.qfbj = (qfbj == null || qfbj=="")?0:qfbj*1;
    classInfo.xf1 = (xf1 == null || xf1=="")?0:xf1*1;
    classInfo.xf2 = (xf2 == null || xf2=="")?0:xf2*1;
    classInfo.tf = (tf == null || tf=="")?0:tf*1;
    classInfo.spend = (spend == null || spend=="")?0:spend*1;

    classInfo.total = classInfo.ssk+classInfo.qfbj+classInfo.xf1+classInfo.xf2-classInfo.tf;

    return classInfo;
}

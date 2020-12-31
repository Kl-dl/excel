//引入文件
//给导出按钮增加一个onclick='downloadexcel("导出表的id","设置文件名")'；
//给需要导出的table定义一个:id="table";

//执行
function downloadexcel(obj,tablename){
    var ret = window.confirm("是否导出Excel?");
    if(ret){
        var oHtml = document.getElementById(obj).outerHTML;
        var excelHtml = `
                            <html>
                            <head>
                                <meta charset='utf-8' />
                            </head>
                            <body>
                                ${oHtml}
                            </body>
                            </html>
                        `;
        var excelBlob = new Blob([excelHtml], {type: 'application/vnd.ms-excel'})
        var oA = document.createElement('a');
        oA.href = URL.createObjectURL(excelBlob);
        oA.download = tablename+'.xls';
        oA.click();
    }else{
        return false;
    }
}
//
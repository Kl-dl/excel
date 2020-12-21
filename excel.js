//引入文件
//定义一个id为downloadexcel的按钮；
//给需要导出的table定义一个:id="table";

//执行
$(function () {
    $("#downloadexcel").click(function(){
        var ret = window.confirm("是否导出Excel?");
        if(ret){
            downloadexcel("#table","导出的表名");
        }else{
            return false;
        }
    })
});
//

function downloadexcel(obj,filename){
    var title = [[]];
    for(var i = 0; i < $(obj).find("tr:eq(0) th").length; i++){
        title[0][i] = $(obj).find("tr:eq(0) th:eq("+i+")").text();		
    }
    if(title[0].length == 0){
        for(var i = 0; i < $(obj).find("tr:eq(0) td").length; i++){
            title[0][i] = $(obj).find("tr:eq(0) td:eq("+i+")").text();		
        }
    }
    var data = [];
    var arr = [];
    for(var i = 1; i < $(obj).find("tr").length; i++){
        for(var n = 0; n < $(obj).find("tr:eq("+i+") td").length; n++)	{
            console.log($(obj).find("tr:eq("+i+") td:eq("+n+")").text());
            arr[n] = $(obj).find("tr:eq("+i+") td:eq("+n+")").text();
        }
        data[(i-1)] = arr;
        arr = [];
    }
    toExcel(filename, data, title); 
}


function toExcel(FileName, JSONData, ShowLabel) {  

    var arrData = typeof JSONData != 'object' ? JSON.parse(JSONData) : JSONData;  
      
    var excel = '<table>';      
      
    var row = "<tr align='left'>";
    for (var i = 0, l = ShowLabel.length; i < l; i++) {  
        for (var key in ShowLabel[i]) {
            row += "<td>" + ShowLabel[i][key] + '</td>';  
        }
    }
       
    excel += row + "</tr>";  

    for (var i = 0; i < arrData.length; i++) {  
        var rowData = "<tr align='left'>"; 

        for (var y = 0; y < ShowLabel.length; y++) {
            for(var k in ShowLabel[y]){
                if (ShowLabel[y].hasOwnProperty(k)) {
                     rowData += "<td style='vnd.ms-excel.numberformat:@'>" + (arrData[i][k]===null? "" : arrData[i][k]) + "</td>";　　　　 
                }
            }
        }

        excel += rowData + "</tr>";  
    }  

    excel += "</table>";  

    var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel'>";  
    excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';  
    excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel';  
    excelFile += '; charset=UTF-8">';  
    excelFile += "<head>";  
    excelFile += "<!--[if gte mso 9]>";  
    excelFile += "<xml>";  
    excelFile += "<x:ExcelWorkbook>";  
    excelFile += "<x:ExcelWorksheets>";  
    excelFile += "<x:ExcelWorksheet>";  
    excelFile += "<x:Name>";  
    excelFile += "{worksheet}";  
    excelFile += "</x:Name>";  
    excelFile += "<x:WorksheetOptions>";  
    excelFile += "<x:DisplayGridlines/>";  
    excelFile += "</x:WorksheetOptions>";  
    excelFile += "</x:ExcelWorksheet>";  
    excelFile += "</x:ExcelWorksheets>";  
    excelFile += "</x:ExcelWorkbook>";  
    excelFile += "</xml>";  
    excelFile += "<![endif]-->";  
    excelFile += "</head>";  
    excelFile += "<body>";  
    excelFile += excel;  
    excelFile += "</body>";  
    excelFile += "</html>";  

    var uri = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelFile);  
      
    var link = document.createElement("a");      
    link.href = uri;  
      
    link.style = "visibility:hidden";  
    link.download = FileName + ".xls";  
      
    document.body.appendChild(link);  
    link.click();  
    document.body.removeChild(link);  
}

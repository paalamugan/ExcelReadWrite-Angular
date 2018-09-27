import { Component, OnInit } from '@angular/core';
declare var XLSX:any;
type AOA = any[][];
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'Excel Sheet Data';
  data: AOA = [ [1, 2], [3, 4] ];
	wopts=XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';
  onFileChange(evt: any) {
    /* wire up file reader */
    var HTMLOUT = document.getElementById('htmlout');
   if(HTMLOUT.innerText.length!=0) {
    HTMLOUT.innerHTML="";
   }
		const target: DataTransfer = <DataTransfer>(evt.target);
		if (target.files.length !== 1){
      throw new Error('Cannot use multiple files');
    } 
		const reader: FileReader = new FileReader();
		reader.onload = (e: any) => {
      /* read workbook */
      this.htmltrue=true;
			const bstr: string = e.target.result;
		//	const wb= XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});
      var workbook = XLSX.read(bstr, {type:"binary"});
      workbook.SheetNames.forEach(sheetName => {
        var htmlstr=XLSX.write(workbook,{sheet:sheetName,type:'string',bookType:'html'});
        HTMLOUT.innerHTML +=htmlstr;
      });

     
			/* grab first sheet */
		//const wsname: string = wb.SheetNames[0];
		//	const ws= XLSX.WorkSheet = wb.Sheets[wsname];

			/* save data */
		//	this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
		};
		reader.readAsBinaryString(target.files[0]);
  }
  htmltrue:boolean=true;
  export(tableID, filename = '') {
    var downloadLink;
    var dataType = 'application/vnd.ms-excel';
    var tableSelect = document.getElementById(tableID);
    if(tableSelect.innerText.length==0){
      alert("please select one Excel File");
    }else{
      if(!this.htmltrue){
        alert("already data has downloaded");
      }else{
        var tableHTML = tableSelect.outerHTML.replace(/ /g, '%20');
    
        filename = filename?filename+'.xls':'excel_data.xls';
        
        downloadLink = document.createElement("a");
        
        document.body.appendChild(downloadLink);
        
        if(navigator.msSaveOrOpenBlob){
            var blob = new Blob(['\ufeff', tableHTML], {
                type: dataType
            });
            navigator.msSaveOrOpenBlob( blob, filename);
        }else{
            // Create a link to the file
            downloadLink.href = 'data:' + dataType + ', ' + tableHTML;
        
            // Setting the file name
            downloadLink.download = filename;
            
            //triggering the function
            downloadLink.click();
            this.htmltrue=false;
        }
       
      }
      }
     
    

    /* generate worksheet */
    //var HTMLOUT = document.getElementById('htmlout');
	//	const ws= XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(htmlout);

		/* generate workbook and add the worksheet */
		//const wb= XLSX.WorkBook = XLSX.utils.book_new();
	//	XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

		/* save to file */
		//XLSX.writeFile(wb, this.fileName);
  }
  applyFliter(){
    var input, filter, table, tr, td, i;
    input = document.getElementById("myInput");
    filter = input.value.toUpperCase();
    table = document.getElementById("htmlout");
    tr = table.getElementsByTagName("tr");
    for (i = 1; i < tr.length; i++) {
    td = tr[i].getElementsByTagName("td")[0];
    if (td) {
    if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
    tr[i].style.display = "";
    } else {
    tr[i].style.display = "none";
    }
    } 
    }
    }
    ngOnInit(){
      //var HTMLOUT = document.getElementById('htmlout');
      //   var url = "../assets/Onward_TS_JULY2018_Sremugaan.xls";
      
      // var req = new XMLHttpRequest();
      // req.open("GET", url, true);
      // req.responseType = "arraybuffer";
      // req.onload = function(e) {
      //   var data = new Uint8Array(req.response);
      //   var arr =new Array();
      //   for(var i=0;i!=data.length;i++){
      //     arr[i]=String.fromCharCode(data[i]);
      //   }
      //   var bstr = arr.join("");
      //   HTMLOUT.innerHTML="";
        
      //    var workbook = XLSX.read(bstr, {type:"binary"});
      //    workbook.SheetNames.forEach(sheetName => {
      //      var htmlstr=XLSX.write(workbook,{sheet:sheetName,type:'string',bookType:'html'});
      //      HTMLOUT.innerHTML +=htmlstr;
      //    });
        
      // }
      // req.send();
    }
}

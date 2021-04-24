window.onload = function()
{
	var sheets;
	var sheetList = new Array();
	var data;

	var filesProcessed = 0;
	var totalFiles=0;

	var printPageHtml = $('#print-page').html();

	var grade = '';
	var section = '';

	var invoiceData, feeInfoName;
	var selectedSheetName = "";

	var filesList = new Array();

	var defaultFeeDescriptionTable = $('.fee-description')[0].outerHTML;
	var invoiceNumber;

	$( "#datepicker" ).datepicker().datepicker('setDate', 'today')
	.datepicker( "option", "dateFormat", 'dd-mm-yy' );

	function rd(t,r,c)
	{
		var cellData = rdh(t,r,c);
		if (cellData != null) return cellData.innerHTML;
		else return null;
	}

	function rdh(t,r,c)
	{
		if (t.rows[r] == undefined) return null;
		return t.rows[r].cells[c];
	}

	function wrv(t,r,c,v)
	{
		t.rows[r].cells[c].innerHTML = v;
	}

	var ExcelToJSON = function() {

		this.parseExcel = function(file) {
			var reader = new FileReader();

			sheets = new Array();

			reader.onload = function(e) {
				var data = e.target.result;
				var workbook = XLSX.read(data, {
					type: 'binary'
				});

				workbook.SheetNames.forEach(function(sheetName) {
					// Here is your object
					var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
					var XL = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName]);
					var json_object = JSON.stringify(XL_row_object);
					//console.log(sheetName,JSON.parse(json_object));
					var sheetTable = document.createElement('div');
					$(sheetTable).html(XL);
					//$("body").append(sheetTable);
					sheets.push({name:sheetName, table: sheetTable});

					if (sheetName.toLowerCase() == "chaitra")
					{
						setClassVals();
					}
				})
			};

			reader.onerror = function(ex) {
				console.log(ex);
				$('#message>p').html(ex);
				$('#message').css({backgroundColor:'#f59a9a'});
			};

			reader.readAsBinaryString(file);
		};
	};


	function handleFileSelect(evt) {
		console.log(evt.target.files);
		filesList = evt.target.files;
		totalFiles += filesList.length;

		for(var i=0;i<totalFiles;i++)
		{
			var files = evt.target.files; // FileList object
			var xl2json = new ExcelToJSON();
			xl2json.parseExcel(files[i]);
		}
	}

	document.getElementById('upload').addEventListener('change', handleFileSelect, false);

	function getTableData(sheetName)
	{
		for(var i=0;i<sheets.length;i++)
		{
			if (sheetName == sheets[i].name)
			{
				return sheets[i].table;
				break;
			}
		}
		return null;
	}

	function setClassVals()
	{
		var data = getTableData("Students");
		if (data!=null)
		{
			data = data.getElementsByTagName('table')[0];
			var _grade = rd(data, 0,1);
			grade = _grade;
			$('#class-data>span').html(grade);
			$('#class-data').show();
			populateSheetSelection();	
			generateList(data);	
			generateInvoiceData();
		}
	}

	function populateSheetSelection()
	{
		$('#sheet-select').html(' ');
		for(var i=1;i<sheets.length;i++)
		{
			$('#sheet-select').append('<option value="'+sheets[i].name+'">'+sheets[i].name+'</option>');
		}
		$('#sheet-select').show();
		$('#view-all-btn').show();
	}

	function generateList(data)
	{
		var str = '<tr><th>Sec</th><th>Roll</th><th>Name</th><th>Action</th><th><input type="checkbox" value="main" id="main-checkbox" checked></th></tr>';
		var a = 2;
		var value = "";
		invoiceData = [];
		while(a<100)
		{
			value = rd(data,a,3);
			if (value == null) break;
			name = value;
			var section = rd(data,a,1);
			var roll = rd(data,a,2);
			//console.log(value);
			invoiceData.push({name: name, section: section, roll: roll, data: null});
			str += '<tr><td>'+section+'</td><td>'+roll+'</td><td>'+name+'</td><td><button class="viewbtn" value="'+roll+'_'+section+'">View</button</td><td><input type="checkbox" class="student-checkboxes" id="checkbox_'+section+'_'+roll+'" checked></td></tr>'

			a++;
		}
		$('#students-list').html(str);

		$('.viewbtn').click(function(e){
			var val = this.getAttribute('value').split('_');
			var roll = val[0];
			var section = val[1];
			
			invoiceNumber = $('#invoice-input').val();
			for(var i=0;i<invoiceData.length;i++)
			{
				if (invoiceData[i].roll==roll && invoiceData[i].section==section)
				{
					//console.log("Student Selected: ",roll,invoiceData[i].name);
					$('.mutli-print').remove();
					setData('print-page',i);
					$('#print-page').show();
					break;
				}	
			}
		});

		$('#main-checkbox').click(function(){
			console.log(this.checked);
			var isChecked = this.checked;
			$('.student-checkboxes').each(function(){
				$(this).attr("checked",isChecked);
			})
		});
	}

	$('#sheet-select').change(function(){
		
		generateInvoiceData();		
		
	});

	function generateInvoiceData()
	{
		selectedSheetName = $('#sheet-select').children("option:selected").val();
		console.log("selectedSheetName");
		var data = getTableData(selectedSheetName);
		if (data!=null)
		{
			data = data.getElementsByTagName('table')[0];
			//invoiceData = [];
			feeInfoName = [];

			var a = 5;
			while(a<1000)
			{
				var feeInfo = rd(data, 1,a);
				if (feeInfo!=null)
				{
					feeInfoName.push(feeInfo);
					console.log(feeInfo);
				}
				else break;
				a++;
			}
			a = 2;
			var b = 3;
			while(a<1000)
			{
				var name = rd(data, a,b);
				if (name!=null)
				{
					var section = rd(data, a, 1);
					var roll = rd(data, a, 2);
					var feeInfoData = [];
					for(var i=0;i<feeInfoName.length;i++)
					{
						feeInfoData.push([rd(data,a,b+i+2)]);
					}
					for(var i=0;i<invoiceData.length;i++)
					{

						if (invoiceData[i].name == name && invoiceData[i].section == section && invoiceData[i].roll == roll)
						{
							invoiceData[i].data = feeInfoData;
							break;
						}
					}
				}else break;
				a++;
			}
			console.log(invoiceData);
		}
	}

	$('#view-all-btn').click(function(){
		var counter = 0;
		invoiceNumber = $('#invoice-input').val();
		$('.mutli-print').remove();

		for(var i=0;i<invoiceData.length;i++)
		{
			var isChecked = $('#checkbox_'+invoiceData[i].section+'_'+invoiceData[i].roll)[0].checked;
			if (counter==0)
			{
				setData('print-page',i);
				if (isChecked) $('#print-page').show();
				else $('#print-page').hide();
			}
			else
			{	
				if (isChecked)
				{
					var pg_name = 'page-'+(counter+1);
				
					str = '<div id="'+pg_name+'" class="page mutli-print"></div>';
					$('body').append(str);
					$('#'+pg_name).append(printPageHtml);
					setData(pg_name,i);
				}

			}
			counter++;
		}
	});

	function setData(pageName, index)
	{

		var ivData = invoiceData[index];
		$('#'+pageName + ' .invoice-num')[0].innerHTML = invoiceNumber;
		invoiceNumber++;
		$('#'+pageName + ' .invoice-date')[0].innerHTML = $('#datepicker').val();
		$('#'+pageName + ' .student-name')[0].innerHTML = ivData.name;
		$('#'+pageName + ' .class-name')[0].innerHTML = grade;
		$('#'+pageName + ' .month-name')[0].innerHTML = selectedSheetName;

		$('#'+pageName + ' .fee-description').remove();
		$('#'+pageName + ' .note').before((defaultFeeDescriptionTable));

		var t1 = $('#'+pageName + ' .fee-description')[0];

		var counter = 1;
		for(var i=2;i<ivData.data.length;i++)
		{
			if (ivData.data[i] != "")
			{

				if (counter>8)
				{
					$('#'+pageName + ' .total-invoice-row').before('<tr><td>'+counter+'</td><td></td><td></td></tr>');
					var mytd = $('#'+pageName + ' .fee-description td');
					mytd.height(mytd.height()-2);
				}	
				wrv(t1, counter, 1, feeInfoName[i]);
				wrv(t1, counter, 2, ivData.data[i]);
				counter++;

			}
		}

		// Previous Dues if Any
		if (ivData.data[1] != "" && ivData.data[1] != 0)
		{
			if (counter>8)
			{
				$('#'+pageName + ' .total-invoice-row').before('<tr><td>'+counter+'</td><td></td><td></td></tr>');
				var mytd = $('#'+pageName + ' .fee-description td');
				mytd.height(mytd.height()-2);
			}	
			wrv(t1, counter, 1, feeInfoName[1]);
			wrv(t1, counter, 2, ivData.data[1]);
			counter++;
		}

		//Total
		if (ivData.data[0] != "")
		{
			$('#'+pageName + ' .total-invoice').html(ivData.data[0]);
			$('#amount-words').html(capitalizeTheFirstLetterOfEachWord(numWords(ivData.data[0])) + " Only");
		}

	}

	$('#print-btn').click(function(){
		window.print();
	});

}



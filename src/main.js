window.onload = function () {
  var sheets;

  var printPageHtml = $('#print-page').html();

  var grade = '';

  var invoiceData, feeInfoName;
  var selectedSheetName = "";

  var filesList = new Array();

  var defaultFeeDescriptionTable = $('.fee-description')[0].outerHTML;
  var invoiceNumber;

  $("#datepicker").datepicker().datepicker('setDate', 'today')
    .datepicker("option", "dateFormat", 'dd-mm-yy');

  function rd(t, r, c) {
    var cellData = rdh(t, r, c);
    if (cellData != null) return cellData.innerHTML;
    else return null;
  }

  function rdh(t, r, c) {
    if (t.rows[r] == undefined) return null;
    return t.rows[r].cells[c];
  }

  function wrv(t, r, c, v) {
    t.rows[r].cells[c].innerHTML = v;
  }

  var ExcelToJSON = function () {

    this.parseExcel = function (file) {
      var reader = new FileReader();

      sheets = new Array();

      reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: 'binary'
        });

        workbook.SheetNames.forEach(function (sheetName) {
          // Here is your object
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          var XL = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName]);
          var json_object = JSON.stringify(XL_row_object);

          var sheetTable = document.createElement('div');
          $(sheetTable).html(XL);

          sheets.push({ name: sheetName, table: sheetTable });

          if (sheetName.toLowerCase() == "chaitra") {
            setClassVals();
          }

          $('.print-double-page-container').remove();
          $('#print-page').show();

        })
      };

      reader.onerror = function (ex) {
        console.log(ex);
        $('#message>p').html(ex);
        $('#message').css({ backgroundColor: '#f59a9a' });
      };

      reader.readAsBinaryString(file);
    };
  };


  function handleFileSelect(evt) {
    console.log(evt.target.files);
    filesList = evt.target.files;

    for (var i = 0; i < filesList.length; i++) {
      var files = evt.target.files; // FileList object
      var xl2json = new ExcelToJSON();
      xl2json.parseExcel(files[i]);
    }
  }

  document.getElementById('upload').addEventListener('change', handleFileSelect, false);

  function getTableData(sheetName) {
    for (var i = 0; i < sheets.length; i++) {
      if (sheetName == sheets[i].name) {
        return sheets[i].table;
      }
    }
    return null;
  }

  function setClassVals() {
    var data = getTableData("Students");
    if (data != null) {
      data = data.getElementsByTagName('table')[0];
      var _grade = rd(data, 0, 1);
      grade = _grade;
      $('#class-data>span').html(grade);
      $('#class-data').show();
      populateSectionSelection(data);
      populateSheetSelection();
      generateList(data);
      generateInvoiceData();
    }
  }

  function populateSectionSelection(data) {
    const sections = ["ALL"];

    let a = 2;
    let value;

    do {
      value = rd(data, a, 1);

      if (value && !sections.some((section) => section == value)) sections.push(value);
      a++;
    } while (value);

    $('#section-select').html('');
    sections.forEach((section) => {
      $('#section-select').append('<option value="' + section + '">' + section + '</option>');
    });


    document.getElementById("section-data").style.display = "flex";

    $('#section-select').on("change", () => {
      generateList(data);
      generateInvoiceData();
    });
  }

  function populateSheetSelection() {
    $('#sheet-select').html(' ');
    for (var i = 1; i < sheets.length; i++) {
      $('#sheet-select').append('<option value="' + sheets[i].name + '">' + sheets[i].name + '</option>');
    }
    $('#sheet-select').show();
    $('#view-all-btn').show();
  }

  function generateList(data) {
    var str = '<tr><th>Sec</th><th>Roll</th><th>Name</th><th>Action</th><th><input type="checkbox" value="main" id="main-checkbox" checked></th></tr>';
    var a = 2;
    var value = "";
    invoiceData = [];
    do {
      value = rd(data, a, 3);
      if (value == null) break;

      var section = rd(data, a, 1);
      var roll = rd(data, a, 2);

      const sectionSelect = $("#section-select").val();
      if (sectionSelect == "ALL" || section == sectionSelect) {
        invoiceData.push({ name: value, section: section, roll: roll, data: null });
        str += '<tr><td>' + section + '</td><td>' + roll + '</td><td>' + value + '</td><td><button class="viewbtn" value="' + roll + '_' + section + '">View</button</td><td><input type="checkbox" class="student-checkboxes" id="checkbox_' + section + '_' + roll + '" checked></td></tr>'
      }

      a++;
    } while (value);

    $('#students-list').html(str);

    $('.viewbtn').click(function (e) {
      var val = this.getAttribute('value').split('_');
      var roll = val[0];
      var section = val[1];

      invoiceNumber = $('#invoice-input').val();

      for (var i = 0; i < invoiceData.length; i++) {
        if (invoiceData[i].roll == roll && invoiceData[i].section == section) {

          $('#print-page').hide();
          $('.print-double-page-container').remove();

          createPage(i);
          break;
        }
      }
    });

    $('#main-checkbox').click(function () {
      var isChecked = this.checked;
      $('.student-checkboxes').each(function () {
        $(this).attr("checked", isChecked);
      })
    });
  }

  $('#view-all-btn').click(function () {
    invoiceNumber = $('#invoice-input').val();

    $('#print-page').hide();
    $('.print-double-page-container').remove();

    for (var i = 0; i < invoiceData.length; i++) {
      var isChecked = $('#checkbox_' + invoiceData[i].section + '_' + invoiceData[i].roll)[0].checked;

      if (isChecked) {
        createPage(i);
      }
    }
  });

  function createPage(index) {
    var doublePageContainer = Array.from(document.querySelectorAll('.print-double-page-container'))?.pop();

    if (!doublePageContainer || doublePageContainer.children.length == 2) {
      doublePageContainer = document.createElement("div");
      doublePageContainer.classList.add("print-double-page-container");

      document.body.appendChild(doublePageContainer);
    }

    var pg_name = 'page-' + (index + 1);

    document.body.appendChild(doublePageContainer);

    const page = document.createElement("div");
    page.id = pg_name;
    page.classList.add("page");
    doublePageContainer.appendChild(page);

    $('#' + pg_name).append(printPageHtml);
    setData(pg_name, index);
  }

  $('#sheet-select').change(function () {

    generateInvoiceData();

  });

  function generateInvoiceData() {
    selectedSheetName = $('#sheet-select').children("option:selected").val();

    var data = getTableData(selectedSheetName);
    console.log("data", data);
    if (data != null) {
      data = data.getElementsByTagName('table')[0];
      
      feeInfoName = [];

      var a = 5;
      while (a < 1000) {
        var feeInfo = rd(data, 1, a);
        if (feeInfo != null) {
          feeInfoName.push(feeInfo);
        }
        else break;
        a++;
      }
      a = 2;
      var b = 3;
      while (a < 1000) {
        var name = rd(data, a, b);
        if (name != null) {
          var section = rd(data, a, 1);
          var roll = rd(data, a, 2);
          var feeInfoData = [];
          for (var i = 0; i < feeInfoName.length; i++) {
            feeInfoData.push([rd(data, a, b + i + 2)]);
          }
          for (var i = 0; i < invoiceData.length; i++) {

            if (invoiceData[i].name == name && invoiceData[i].section == section && invoiceData[i].roll == roll) {
              invoiceData[i].data = feeInfoData;
              break;
            }
          }
        } else break;
        a++;
      }
    }
  }

  function setData(pageName, index) {

    var ivData = invoiceData[index];
    if (invoiceNumber != "") {
      $('#' + pageName + ' .invoice-num')[0].innerHTML = invoiceNumber;
      invoiceNumber++;
    } else {
      $('#' + pageName + ' .invoice-num')[0].innerHTML = "";
    }
    $('#' + pageName + ' .invoice-date')[0].innerHTML = $('#datepicker').val();
    $('#' + pageName + ' .student-name')[0].innerHTML = ivData.name;
    $('#' + pageName + ' .class-name')[0].innerHTML = grade;
    $('#' + pageName + ' .month-name')[0].innerHTML = selectedSheetName;

    $('#' + pageName + ' .fee-description').remove();
    $('#' + pageName + ' .qr-signature').before((defaultFeeDescriptionTable));

    var t1 = $('#' + pageName + ' .fee-description')[0];

    var counter = 1;
    for (var i = 2; i < ivData.data.length; i++) {
      var amount = ivData.data[i]
      if (amount != "") {

        if (counter > 8) {
          $('#' + pageName + ' .total-invoice-row').before('<tr><td>' + counter + '</td><td></td><td></td></tr>');
          var mytd = $('#' + pageName + ' .fee-description td');
          mytd.height(mytd.height() - 2);
        }
        wrv(t1, counter, 1, feeInfoName[i]);
        wrv(t1, counter, 2, amount);
        counter++;

      }
    }

    // Previous Dues if Any
    if (ivData.data[1] != "" && ivData.data[1] != 0) {
      if (counter > 8) {
        $('#' + pageName + ' .total-invoice-row').before('<tr><td>' + counter + '</td><td></td><td></td></tr>');
        var mytd = $('#' + pageName + ' .fee-description td');
        mytd.height(mytd.height() - 2);
      }
      wrv(t1, counter, 1, feeInfoName[1]);
      wrv(t1, counter, 2, ivData.data[1]);
      counter++;
    }

    console.log(ivData);

    //Total
    if (ivData.data[0] != "") {
      $('#' + pageName + ' .total-invoice').html(ivData.data[0]);
      $('#' + pageName + ' .amount-words').html(capitalizeTheFirstLetterOfEachWord(numWords(ivData.data[0])) + " Only");
    }

  }

  $('#print-btn').click(function () {
    window.print();
  });

}



let selectedFile;
document.getElementById("input").addEventListener("change", (event) => {
  selectedFile = event.target.files[0];
});

document.getElementById("button").addEventListener("click", () => {
  if (selectedFile) {
    let fileReader = new FileReader();
    fileReader.readAsBinaryString(selectedFile);
    fileReader.onload = (event) => {
      let data = event.target.result;
      let workbook = XLSX.read(data, { type: "binary" });
      workbook.SheetNames.forEach((sheet) => {
        let rowObject = XLSX.utils.sheet_to_row_object_array(
          workbook.Sheets[sheet]
        );
        // sort by 'Supervisor'
        rowObject = rowObject.sort(function (a, b) {
          return a.Supervisor.localeCompare(b.Supervisor);
        });
        // sort by 'Nombre'
        rowObject = rowObject.sort(function (a, b) {
          if(a.Supervisor.localeCompare(b.Supervisor) == 0)
          return a.Nombre.localeCompare(b.Nombre);
        });

        localStorage.setItem("datos", JSON.stringify(rowObject));
      });
    };

    FillTable("datos");
  }
});

function FillTable(storage) {
  DeleteTable();

  var datosJSON = localStorage.getItem(storage);

  objJSON = JSON.parse(datosJSON);

  var table = document.getElementById("data_table");

  // Header
  var tr = document.createElement("tr");

  tr.innerHTML =
    '<th onclick="sortTable(0)" style="width: 16%;">Nombre</th>' +
    '<th onclick="sortTable(1)" style="width: 16%;">Cargo</th>' +
    '<th onclick="sortTable(2)" style="width: 16%;">Supervisor</th>' +
    '<th onclick="sortTable(3)" style="width: 16%;">Clase</th>' +
    '<th onclick="sortTable(4)" style="width: 16%;">Subsidiaria</th>' +
    '<th onclick="sortTable(5)" style="width: 16%;">Departamento</th>';

  table.appendChild(tr);

  objJSON.forEach(function (object) {
    var tr = document.createElement("tr");
    tr.innerHTML =
      '<td style="width: 16%;">' +
      object.Nombre +
      "</td>" +
      '<td style="width: 16%;">' +
      object.Cargo +
      "</td>" +
      '<td style="width: 16%;">' +
      object.Supervisor +
      "</td>" +
      '<td style="width: 16%;">' +
      object.Clase +
      "</td>" +
      '<td style="width: 16%;">' +
      object.Subsidiaria +
      "</td>" +
      '<td style="width: 16%;">' +
      object.Departamento +
      "</td>";
    table.appendChild(tr);
  });
}

function DeleteTable() {
  var Table = document.getElementById("data_table");
  Table.innerHTML = "";
}

function sortTable(index) {
  var datosJSON = localStorage.getItem("datos");

  objJSON = JSON.parse(datosJSON);
/**
 * 0. Nombre
 * 1. Cargo
 * 2. Supervisor
 * 3. Clase
 * 4. Subsidiaria
 * 5. Departamento
*/

  objJSON = objJSON.sort(function (a, b) {
    switch (index) {
      case 0:
        // sort by 'Supervisor'
        objJSON = objJSON.sort(function (a, b) {
          return a.Supervisor.localeCompare(b.Supervisor);
        });
        // sort by 'Nombre'
        objJSON = objJSON.sort(function (a, b) {
          if(a.Supervisor.localeCompare(b.Supervisor) == 0)
          return a.Nombre.localeCompare(b.Nombre);
        });
      case 1:
        return a.Cargo.localeCompare(b.Cargo);
      case 2:
        return a.Supervisor.localeCompare(b.Supervisor);
      case 3:
        return a.Clase.localeCompare(b.Clase);
      case 4:
        return a.Subsidiaria.localeCompare(b.Subsidiaria);
      case 5:
        return a.Departamento.localeCompare(b.Departamento);
    }
  });

  localStorage.setItem("datos", JSON.stringify(objJSON));

  FillTable("datos");
}

function filter() {
  var datosJSON = localStorage.getItem("datos");
  var objJSON = JSON.parse(datosJSON);
  var supervisorFilter = document.getElementById("supervisor_filter").value;
  var claseFilter = document.getElementById("clase_filter").value;
  var departamentoFilter = document.getElementById("departamento_filter").value;
  var subsidiariaFilter = document.getElementById("subsidiaria_filter").value;

  objJSON = objJSON
    .filter(function (obj) {
      return obj.Supervisor.toLowerCase().includes(
        supervisorFilter.toLowerCase()
      );
    })
    .filter(function (obj) {
      return obj.Clase.toLowerCase().includes(claseFilter.toLowerCase());
    })
    .filter(function (obj) {
      return obj.Departamento.toLowerCase().includes(
        departamentoFilter.toLowerCase()
      );
    })
    .filter(function (obj) {
      return obj.Subsidiaria.toLowerCase().includes(
        subsidiariaFilter.toLowerCase()
      );
    });

  localStorage.setItem("datosFiltrados", JSON.stringify(objJSON));

  FillTable("datosFiltrados");
}

function exportToCSV(filename) {
  var csv = [];

  var rows = document.querySelectorAll("table tr");
  
  for (var i = 0; i < rows.length; i++) {
      var row = [], cols = rows[i].querySelectorAll("td, th");

      for (var j = 0; j < cols.length; j++) 
        row.push(cols[j].innerText);

      csv.push(row.join(","));        
  }

  filename += '.csv';
  csv = csv.join("\n");

  var csvFile = new Blob([csv], {type: "text/csv"});
  var downloadLink = document.createElement("a");

  downloadLink.download = filename;
  downloadLink.href = window.URL.createObjectURL(csvFile);
  downloadLink.style.display = "none";
  document.body.appendChild(downloadLink);
  downloadLink.click();
}

function exportToExcel(filename){
  var dataType = 'application/vnd.ms-excel';
  var tableSelect = document.getElementById('data_table');
  var tableHTML = tableSelect.outerHTML.replace(/ /g, '%20');
  
  filename += '.xls';
  
  var downloadLink = document.createElement("a");
  document.body.appendChild(downloadLink);
  
  if(navigator.msSaveOrOpenBlob){
      var blob = new Blob(['ufeff', tableHTML], {
        type: dataType
      });
      navigator.msSaveOrOpenBlob( blob, filename);
  }else{
      downloadLink.href = 'data:' + dataType + ', ' + tableHTML;
      downloadLink.download = filename;
      downloadLink.click();
  }
}

function exportToPDF(filename){
  var doc = new jsPDF('p', 'pt', 'letter');
  var pageHeight = 0;  
  pageHeight = doc.internal.pageSize.height;  
  specialElementHandlers = {
      '#bypassme': function(element, renderer) {
        return true  
      }  
  };  
  margins = {  
    top: 150,  
    bottom: 60,  
    left: 40,  
    right: 40,  
    width: 600  
  };  
  var y = 20;  
  doc.setLineWidth(2);  
  doc.text(200, y = y + 30, "Lista de Empleados");  
  doc.autoTable({  
      html: '#data_table',  
      startY: 70,  
      theme: 'grid',  
      styles: {  
        minCellHeight: 40  
      }  
  })  
  doc.save(filename + '.pdf'); 
}

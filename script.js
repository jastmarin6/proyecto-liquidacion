document.getElementById('fileInput').addEventListener('change', function(event) {
    let file = event.target.files[0];
    let reader = new FileReader();

    if (!file) {
        alert("Por favor, selecciona un archivo Excel.");
        return;
    }

    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, {type: 'array'});
        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(sheet);

        // 🔹 1. Filtrar registros (Excluir permisos, incapacidades y vacaciones)
        let diasLaborados = jsonData.filter(row => 
            !["PERMISO", "INCAPACIDAD", "VACACIONES", "FUNC ADMON", "RETIRO"].includes(row["ACTIVIDAD"])
        );

        // 🔹 2. Contar días laborados por cada inspector
        let conteoDias = {};
        diasLaborados.forEach(row => {
            if (!conteoDias[row["CEDULA INSPECTOR"]]) {
                conteoDias[row["CEDULA INSPECTOR"]] = new Set();
            }
            conteoDias[row["CEDULA INSPECTOR"]].add(row["FECHA"]);
        });

        // Convertir a array estructurado
        let resultado = Object.keys(conteoDias).map(cedula => ({
            "CEDULA INSPECTOR": cedula,
            "Total Dias Laborados": conteoDias[cedula].size
        }));

        // 🔹 3. Calcular total de inspecciones y bonificaciones
        jsonData.forEach(row => {
            let inspector = resultado.find(ins => ins["CEDULA INSPECTOR"] === row["CEDULA INSPECTOR"]);
            if (inspector) {
                inspector["Total Inspecciones"] = (inspector["Total Inspecciones"] || 0) + row["TOTAL REVISIONES"];
                inspector["Total LM"] = (inspector["Total LM"] || 0) + row["LM"];
                inspector["Total Suspensiones"] = (inspector["Total Suspensiones"] || 0) + (row["TOTAL SUSPENSIONES"] * 3000);
                inspector["Auxilio Moto"] = inspector["Total Dias Laborados"] * 22000;
                inspector["Bono Gestión"] = calcularBonoGestion(inspector["Total Inspecciones"]);
                inspector["Bono Adicional"] = calcularBonoAdicional(inspector["Total Inspecciones"]);
                inspector["Categoría"] = categorizarInspector(inspector["Total Inspecciones"]);
            }
        });

        // 🔹 4. Convertir resultados a archivo Excel
        let newSheet = XLSX.utils.json_to_sheet(resultado);
        let newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Liquidacion");

        // 🔹 5. Generar archivo para descargar
        let excelBuffer = XLSX.write(newWorkbook, {bookType: "xlsx", type: "array"});
        let blob = new Blob([excelBuffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
        let link = document.getElementById("downloadLink");
        link.href = URL.createObjectURL(blob);
        link.download = "Liquidacion_Bonificaciones.xlsx";
        link.style.display = "block";
        link.innerText = "Descargar Resultado";
    };

    reader.readAsArrayBuffer(file);
});

// 🔹 6. Funciones auxiliares
function calcularBonoGestion(inspecciones) {
    if (inspecciones > 250) return (inspecciones - 160) * 15000;
    else if (inspecciones > 210) return (inspecciones - 160) * 13000;
    else if (inspecciones > 180) return (inspecciones - 160) * 10000;
    else return 0;
}

function calcularBonoAdicional(inspecciones) {
    if (inspecciones > 250) return 500000;
    else if (inspecciones > 230) return 330000;
    else if (inspecciones > 210) return 180000;
    else return 0;
}

function categorizarInspector(inspecciones) {
    if (inspecciones > 250) return "ORO";
    else if (inspecciones > 230) return "PLATA";
    else if (inspecciones > 210) return "BRONCE";
    else return "SIN CATEGORIA";
}
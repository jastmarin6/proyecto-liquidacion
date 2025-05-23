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
        jsonData.forEach(row => { 
            if (!conteoDias[row["CEDULA INSPECTOR"]]) {
                conteoDias[row["CEDULA INSPECTOR"]] = new Set();
            }
            conteoDias[row["CEDULA INSPECTOR"]].add(row["FECHA"]);
        });

        // 🔹 3. Crear estructura con todas las columnas
        let resultado = {};
        jsonData.forEach(row => {
            let cedula = row["CEDULA INSPECTOR"];
            if (!resultado[cedula]) {
                resultado[cedula] = {
                    "CENTRO_DE_VINCULACION": row["CENTRO DE VINCULACIÓN"] || "",
                    "CEDULA INSPECTOR": cedula,
                    "NOMBRE_INSPECTOR": row["NOMBRE_INSPECTOR"] || "",
                    "Total Dias Laborados": conteoDias[cedula]?.size || 0,
                    "TOTAL_SUSPENSIONES": 0,
                    "TOTAL_INSPECCIONES": 0,
                    "Total_LM": 0,
                    "Bono_Gestion": 0,
                    "Bono_Adicional": 0,
                    "Auxilio_Moto": 0,
                    "Auxilio_Suspensiones": 0,
                    "Bono_Total": 0,
                    "Auxilio_Total": 0,
                    "Categoria": ""
                };
            }

            // 🔹 4. Acumular valores correctamente
            resultado[cedula]["TOTAL_INSPECCIONES"] += parseInt(row["TOTAL_REVISIONES"]) || 0;
            resultado[cedula]["Total_LM"] += parseInt(row["LM"]) || 0;
            resultado[cedula]["TOTAL_SUSPENSIONES"] += parseInt(row["TOTAL_SUSPENSIONES"])  || 0;
        });

        // 🔹 5. Calcular bonificaciones y auxilios
        Object.values(resultado).forEach(inspector => {
            inspector["Bono_Gestion"] = calcularBonoGestion(inspector["TOTAL_INSPECCIONES"]);
            inspector["Bono_Adicional"] = calcularBonoAdicional(inspector["TOTAL_INSPECCIONES"]);
            inspector["Auxilio_Moto"] = inspector["Total Dias Laborados"] * 22000;
            inspector["Auxilio_Suspensiones"] = inspector["TOTAL_SUSPENSIONES"] * 3000;
            inspector["Bono_Total"] = inspector["Bono_Gestion"] + inspector["Bono_Adicional"];
            inspector["Auxilio_Total"] = inspector["Auxilio_Moto"] + inspector["Auxilio_Suspensiones"];
            inspector["Categoria"] = categorizarInspector(inspector["TOTAL_INSPECCIONES"]);
        });

        // 🔹 6. Convertir a formato Excel
        let newSheet = XLSX.utils.json_to_sheet(Object.values(resultado));
        let newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Liquidacion");

        // 🔹 7. Descargar el archivo generado
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
// 🔹 8. Funciones auxiliares
function calcularBonoGestion(inspecciones) {
    if (inspecciones > 210) return (inspecciones - 160) * 15000;
    else if (inspecciones > 180) return (inspecciones - 160) * 13000;
    else if (inspecciones > 160) return (inspecciones - 160) * 10000;
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
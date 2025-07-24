// Archivo: taskpane.js

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = pintarCelda;
    document.getElementById("enviar").onclick = enviarFormulario;
    document.getElementById("abrirForm").onclick = abrirFormularioEmergente;
  }
});

async function pintarCelda() {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange("A1");
      range.values = [["¡Hola, Benito!"]];
      range.format.fill.color = "yellow";
      await context.sync();
    });
  } catch (error) {
    console.error("Error al ejecutar Excel.run:", error);
  }
}

async function enviarFormulario() {
  const nombre = document.getElementById("nombre").value;
  const edad = document.getElementById("edad").value;

  if (!nombre || !edad) {
    console.log("Por favor completa ambos campos");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.getRange("A3:B3").values = [["Nombre", "Edad"]];
      worksheet.getRange("A4:B4").values = [[nombre, parseInt(edad)]];
      await context.sync();
    });
  } catch (error) {
    console.error("Error al enviar formulario:", error);
  }
}

function abrirFormularioEmergente() {
  const url = "https://beno5000.github.io/Exito1/popup.html";

  Office.context.ui.displayDialogAsync(
    url,
    { height: 50, width: 40, displayInIframe: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          const datos = JSON.parse(arg.message);
          insertarDatosDesdePopup(datos);
          dialog.close();
        });
      } else {
        console.error("No se pudo abrir el formulario emergente:", asyncResult.error.message);
      }
    }
  );
}

async function insertarDatosDesdePopup(datos) {
  try {
    await Excel.run(async (context) => {
      const hoja = context.workbook.worksheets.getActiveWorksheet();
      hoja.getRange("A6:B6").values = [["Nombre", "Edad"]];
      hoja.getRange("A7:B7").values = [[datos.nombre, datos.edad]];
      await context.sync();
    });
  } catch (error) {
    console.error("Error al insertar datos desde el popup:", error);
  }
}

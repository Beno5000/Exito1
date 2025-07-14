Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = pintarCelda;
    document.getElementById("enviar").onclick = enviarFormulario;
    document.getElementById("abrirForm").onclick = abrirFormularioEmergente;
  }
});

async function pintarCelda() {
  try {
    console.log("¡Ejecutando código correcto!");
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
      const rangeHeaders = worksheet.getRange("A3:B3");
      rangeHeaders.values = [["Nombre", "Edad"]];

      const rangeData = worksheet.getRange("A4:B4");
      rangeData.values = [[nombre, parseInt(edad)]];

      await context.sync();
      console.log("Formulario enviado");
    });
  } catch (error) {
    console.error("Error al enviar formulario:", error);
  }
}

// NUEVO: Abrir formulario emergente
function abrirFormularioEmergente() {
  const url = window.location.origin + "/popup.html";

  Office.context.ui.displayDialogAsync(
    url,
    { height: 50, width: 40, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = asyncResult.value;

        // Escuchar los datos enviados desde el popup
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          const datos = JSON.parse(arg.message);
          insertarDatosDesdePopup(datos);
          dialog.close();
        });
      } else {
        console.error("No se pudo abrir el formulario emergente");
      }
    }
  );
}

// NUEVO: Insertar datos desde la ventana emergente
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

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

// DICCIONARIO DE RUCS DE EMPRESAS
const diccRUCs = {
  "IMPULSA365 S.A.C.": "20506760721",
  "ECOMDATA PERU S.A.C.": "20538295541",
  "TEC4BIZ PERU S.A.C.": "20601483689",
  "CARE24 S.A.": "20606564946",
  "ASISTENCIA 365 S.A.C.": "20565750730",
  "QTECH EXPERIENCE S.A.C.": "20613921045",
  "ACCION COMERCIAL 365 S.A.C.": "20614009021",
  "RECUPERATE S.A.C.": "20606358882"
};

// FUNCIÓN PARA BUSCAR EL RUC
function obtenerRUC(nombreEmpresa) {
  if (!nombreEmpresa) return '__________';
  const nombreUpper = nombreEmpresa.toString().trim().toUpperCase();
  
  for (const [empresa, ruc] of Object.entries(diccRUCs)) {
    if (nombreUpper.includes(empresa) || empresa.includes(nombreUpper)) {
      return ruc;
    }
  }
  return '__________';
}

// FUNCIÓN AUXILIAR PARA CONVERTIR PARTE ENTERA A LETRAS
function convertirParteEntera(entero) {
  const Unidades = (num) => {
    switch (num) {
      case 1: return 'UN'; case 2: return 'DOS'; case 3: return 'TRES';
      case 4: return 'CUATRO'; case 5: return 'CINCO'; case 6: return 'SEIS';
      case 7: return 'SIETE'; case 8: return 'OCHO'; case 9: return 'NUEVE';
    }
    return '';
  };

  const DecenasAnd = (strSin, numUnidades) => {
    if (numUnidades > 0) return strSin + ' Y ' + Unidades(numUnidades);
    return strSin;
  };

  const Decenas = (num) => {
    let decena = Math.floor(num / 10);
    let unidad = num - (decena * 10);
    switch (decena) {
      case 1:
        switch (unidad) {
          case 0: return 'DIEZ'; case 1: return 'ONCE'; case 2: return 'DOCE';
          case 3: return 'TRECE'; case 4: return 'CATORCE'; case 5: return 'QUINCE';
          default: return 'DIECI' + Unidades(unidad);
        }
      case 2:
        if (unidad === 0) return 'VEINTE';
        return 'VEINTI' + Unidades(unidad);
      case 3: return DecenasAnd('TREINTA', unidad);
      case 4: return DecenasAnd('CUARENTA', unidad);
      case 5: return DecenasAnd('CINCUENTA', unidad);
      case 6: return DecenasAnd('SESENTA', unidad);
      case 7: return DecenasAnd('SETENTA', unidad);
      case 8: return DecenasAnd('OCHENTA', unidad);
      case 9: return DecenasAnd('NOVENTA', unidad);
      case 0: return Unidades(unidad);
    }
  };

  const Centenas = (num) => {
    let centena = Math.floor(num / 100);
    let decena = num - (centena * 100);
    switch (centena) {
      case 1:
        if (decena > 0) return 'CIENTO ' + Decenas(decena);
        return 'CIEN';
      case 2: return 'DOSCIENTOS ' + Decenas(decena);
      case 3: return 'TRESCIENTOS ' + Decenas(decena);
      case 4: return 'CUATROCIENTOS ' + Decenas(decena);
      case 5: return 'QUINIENTOS ' + Decenas(decena);
      case 6: return 'SEISCIENTOS ' + Decenas(decena);
      case 7: return 'SETECIENTOS ' + Decenas(decena);
      case 8: return 'OCHOCIENTOS ' + Decenas(decena);
      case 9: return 'NOVECIENTOS ' + Decenas(decena);
      case 0: return Decenas(decena);
    }
  };

  if (entero === 0) return 'CERO';
  if (entero < 1000) return Centenas(entero).trim();

  let miles = Math.floor(entero / 1000);
  let resto = entero % 1000;
  let strMiles = '';

  if (miles === 1) {
    strMiles = 'MIL';
  } else if (miles > 1) {
    strMiles = Centenas(miles) + ' MIL';
  }

  let strCentenas = resto > 0 ? ' ' + Centenas(resto) : '';
  return (strMiles + strCentenas).trim();
}

// FUNCIÓN PRINCIPAL PARA MONTOS CON DECIMALES
function numeroALetras(num) {
  if (num === undefined || num === null) return '';
  
  const valorFijo = (Math.round(num * 100) / 100).toFixed(2);
  const [strEntero, strDecimal] = valorFijo.split('.');
  
  const entero = parseInt(strEntero, 10);
  const decimal = parseInt(strDecimal, 10);
  
  let resultado = convertirParteEntera(entero);
  
  if (decimal > 0) {
    resultado += ' CON ' + convertirParteEntera(decimal);
  }
  
  return resultado.trim();
}

function obtenerNombreMes(numeroMes) {
  const meses = [
    'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
  ];
  return meses[numeroMes] || '__________';
}

function cuotasALetras(num) {
  const cuotas = parseInt(num, 10);
  const lista = ['CERO', 'UNA', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'DIEZ', 'ONCE', 'DOCE'];
  return lista[cuotas] || num.toString();
}

function App() {
  const [archivoExcel, setArchivoExcel] = useState(null);
  const [cargando, setCargando] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [darkMode, setDarkMode] = useState(false);
  const [totalFilas, setTotalFilas] = useState(0);

  useEffect(() => {
    if (darkMode) document.documentElement.classList.add('dark');
    else document.documentElement.classList.remove('dark');
  }, [darkMode]);

  const manejarSubida = (e) => {
    const file = e.target.files?.[0] || e.dataTransfer?.files?.[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      setArchivoExcel(file);
      
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const hoja = workbook.Sheets[workbook.SheetNames[0]];
        const datos = XLSX.utils.sheet_to_json(hoja);
        setTotalFilas(datos.length);
      };
      reader.readAsArrayBuffer(file);
    } else {
      alert('Por favor, sube un archivo Excel válido (.xlsx o .xls)');
    }
    setDragOver(false);
  };

  const generarDocumentosMasivos = async () => {
    if (!archivoExcel) return;
    setCargando(true);

    try {
      const data = await archivoExcel.arrayBuffer();
      const workbook = XLSX.read(data);
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const datosExcel = XLSX.utils.sheet_to_json(hoja);

      if (datosExcel.length === 0) throw new Error('Excel vacío');

      const respuestaPlantilla = await fetch('/plantilla.docx');
      if (!respuestaPlantilla.ok) throw new Error('Plantilla no encontrada');
      const arrayBufferPlantilla = await respuestaPlantilla.arrayBuffer();

      const hoy = new Date();
      const diaHoy = hoy.getDate() < 10 ? `0${hoy.getDate()}` : `${hoy.getDate()}`;
      const mesHoyNum = (hoy.getMonth() + 1) < 10 ? `0${hoy.getMonth() + 1}` : `${hoy.getMonth() + 1}`;
      const anioHoy = hoy.getFullYear();
      const mesHoyStr = obtenerNombreMes(hoy.getMonth());
      const fechaHoyFormateada = `${diaHoy}/${mesHoyNum}/${anioHoy}`;

      for (const fila of datosExcel) {
        const datosParaWord = { ...fila };
        const dniCliente = datosParaWord['N°_DOCUMENTO'] || 'Sin_DNI';

        // 1. FORZAMOS A QUE TODAS LAS ETIQUETAS DE FECHA SEAN LA DE HOY (Ignoramos el Excel)
        datosParaWord['FECHA_SOLICITUD'] = fechaHoyFormateada; // La de arriba en el Word
        datosParaWord['FECHA_HOY'] = fechaHoyFormateada;       // La de abajo en el Word
        datosParaWord['DIA_SOLICITUD'] = diaHoy;
        datosParaWord['MES_SOLICITUD'] = mesHoyStr;

        // 2. Formatear N° de Cuotas
        const numCuotas = parseInt(datosParaWord['CUOTAS N°'], 10) || 0;
        datosParaWord['CUOTAS_NUM_FORMATO'] = numCuotas < 10 ? `0${numCuotas}` : `${numCuotas}`;
        datosParaWord['CUOTAS_LETRAS'] = cuotasALetras(numCuotas);

        // 3. Mapeo de montos en letras
        const montoTotal = parseFloat(datosParaWord['MONTO APROBADO']) || 0;
        const montoCuota = parseFloat(datosParaWord['VALOR CUOTA']) || 0;

        datosParaWord['MONTO_LETRAS'] = numeroALetras(montoTotal);
        datosParaWord['CUOTA_LETRAS'] = numeroALetras(montoCuota);

        // 4. FIX TASA DE INTERÉS
        for (const key of Object.keys(datosParaWord)) {
          const keyLower = key.toLowerCase();
          if (keyLower.includes('tasa') || keyLower.includes('interes') || keyLower.includes('interés') || keyLower.includes('tea')) {
            let valor = datosParaWord[key];
            if (valor !== undefined && valor !== null) {
              let valorStr = valor.toString().trim();
              let tienePorcentaje = valorStr.includes('%');
              let num = parseFloat(valorStr);
              
              if (!isNaN(num)) {
                let valorFormateado = num.toFixed(2);
                datosParaWord[key] = tienePorcentaje ? `${valorFormateado}%` : valorFormateado;
              }
            }
          }
        }

        // 5. OBTENER RUC DE LA EMPRESA
        datosParaWord['RUC_EMPRESA'] = obtenerRUC(datosParaWord['EMPRESA']);

        // Construcción del documento individual
        const zip = new PizZip(arrayBufferPlantilla.slice(0)); 
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        doc.render(datosParaWord);
        const out = doc.getZip().generate({ 
          type: 'blob', 
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
        });

        // 6. NUEVO NOMBRE DEL ARCHIVO WORD EXPORTADO
        saveAs(out, `contrato de crédito ${dniCliente}.docx`);
      }

    } catch (error) {
      console.error(error);
      alert('Error procesando la exportación masiva.');
    } finally {
      setCargando(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 relative flex flex-col items-center justify-center p-6 font-sans overflow-hidden transition-colors duration-500">
      <button 
        onClick={() => setDarkMode(!darkMode)}
        className="absolute top-6 right-6 p-3 rounded-full bg-white dark:bg-slate-800 text-slate-600 dark:text-yellow-400 shadow-lg border border-slate-200 dark:border-slate-700 hover:scale-110 transition-all duration-300 z-50"
      >
        {darkMode ? (
          <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z" />
          </svg>
        ) : (
          <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z" />
          </svg>
        )}
      </button>

      <div className="absolute top-0 left-1/4 w-96 h-96 bg-indigo-500/10 dark:bg-indigo-500/20 rounded-full blur-[128px] pointer-events-none" />
      <div className="absolute bottom-0 right-1/4 w-96 h-96 bg-emerald-500/10 rounded-full blur-[128px] pointer-events-none" />

      <div className="bg-white dark:bg-slate-900/60 backdrop-blur-xl p-14 rounded-3xl shadow-xl max-w-xl w-full border border-slate-200 dark:border-slate-800 transition-all duration-500 ease-out hover:scale-[1.01] relative group z-10">
        <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500 rounded-t-3xl" />
        
        <div className="text-center mb-12">
          <div className="inline-flex items-center justify-center w-22 h-22 rounded-2xl bg-indigo-50 dark:bg-slate-800 border border-indigo-100 dark:border-slate-700 shadow-md mb-8">
            <svg className="w-11 h-11 text-indigo-600 dark:text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
            </svg>
          </div>
          <h1 className="text-4xl font-extrabold text-slate-800 dark:text-white tracking-tighter leading-tight">
            De Excel a Word en un clic ⛄
          </h1>
          <p className="text-slate-600 dark:text-slate-400 mt-4 text-lg max-w-sm mx-auto font-medium">
            Sube tu Excel. Cada fila exportará un Word independiente leyendo dinámicamente el Monto Aprobado de cada fila.
          </p>
        </div>
        
        <div 
          className="mb-12 relative"
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); manejarSubida(e); }}
        >
          <label className={`flex flex-col items-center justify-center w-full h-48 rounded-2xl cursor-pointer transition-all duration-300 relative overflow-hidden group
            ${archivoExcel 
              ? 'bg-emerald-50 dark:bg-emerald-950/30 border-2 border-emerald-400 dark:border-emerald-500/50 shadow-sm' 
              : dragOver
                ? 'bg-indigo-50 dark:bg-indigo-950/50 border-2 border-indigo-400 border-dashed'
                : 'bg-slate-50 dark:bg-slate-800/40 border border-slate-300 dark:border-slate-700 hover:border-indigo-400 dark:hover:border-indigo-500'}`}>
            
            <div className="flex flex-col items-center justify-center pt-8 pb-10 px-6 text-center z-10">
              {archivoExcel ? (
                <>
                  <div className="w-18 h-18 bg-emerald-100 dark:bg-emerald-500/10 rounded-full flex items-center justify-center mb-5 border border-emerald-200 dark:border-emerald-500/30">
                    <svg className="w-9 h-9 text-emerald-600 dark:text-emerald-400 animate-pulse" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                  </div>
                  <p className="mb-2 text-base text-slate-800 dark:text-white font-bold truncate max-w-xs">{archivoExcel.name}</p>
                  <p className="text-sm text-emerald-700 dark:text-emerald-300 font-bold bg-emerald-100 dark:bg-emerald-950 px-4 py-1.5 rounded-full border border-emerald-200 dark:border-emerald-500/30">
                    {totalFilas} {totalFilas === 1 ? 'cliente detectado' : 'clientes detectados'}
                  </p>
                </>
              ) : (
                <>
                  <div className={`w-18 h-18 rounded-full flex items-center justify-center mb-5 border ${dragOver ? 'bg-indigo-100 dark:bg-indigo-500/20' : 'bg-white dark:bg-slate-700'}`}>
                    <svg className="w-9 h-9 text-slate-400 group-hover:text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                    </svg>
                  </div>
                  <p className="mb-2 text-base text-slate-600 dark:text-slate-300 font-medium">
                    Haz clic para subir o arrastra el archivo
                  </p>
                  <p className="text-sm text-slate-400 font-medium">Soporta .xlsx y .xls</p>
                </>
              )}
            </div>
            <input type="file" accept=".xlsx, .xls" onChange={manejarSubida} className="hidden" />
          </label>
        </div>

        <button
          onClick={generarDocumentosMasivos}
          disabled={!archivoExcel || cargando}
          className={`group/btn relative w-full py-4.5 px-7 rounded-2xl font-extrabold text-base uppercase tracking-widest transition-all duration-300 overflow-hidden flex justify-center items-center gap-3.5
            ${!archivoExcel || cargando 
              ? 'bg-slate-200 dark:bg-slate-800 cursor-not-allowed text-slate-500 dark:text-slate-400 border border-slate-300 dark:border-slate-700' 
              : 'bg-gradient-to-r from-indigo-600 via-indigo-500 to-indigo-600 text-white shadow-lg transform hover:-translate-y-0.5 active:translate-y-0'
            }`}
        >
          {cargando ? (
            <>
              <svg className="animate-spin h-6 w-6 currentColor" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              <span>Generando {totalFilas} documentos...</span>
            </>
          ) : archivoExcel ? (
            `Descargar ${totalFilas} Hojas Resumen`
          ) : (
            'Esperando fuente de datos'
          )}
        </button>
      </div>

      <div className="mt-18 text-center text-base text-slate-500 font-semibold tracking-wide border-t border-slate-200 dark:border-slate-800 pt-8 max-w-sm mx-auto">
        SMART CASH HERRAMIENTA
        <span className="text-slate-400 dark:text-slate-700 block text-sm mt-1.5">Desarrollado por Smart Cash y Nicolas</span>
      </div>
    </div>
  );
}

export default App;
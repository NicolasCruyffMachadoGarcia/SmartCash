import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

function App() {
  const [archivoExcel, setArchivoExcel] = useState(null);
  const [cargando, setCargando] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  
  // ESTADO PARA EL MODO OSCURO
  const [darkMode, setDarkMode] = useState(false);

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [darkMode]);

  const manejarSubida = (e) => {
    const file = e.target.files?.[0] || e.dataTransfer?.files?.[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      setArchivoExcel(file);
    } else {
      alert('Por favor, sube un archivo Excel válido (.xlsx o .xls)');
    }
    setDragOver(false);
  };

  const generarDocumento = async () => {
    if (!archivoExcel) return;
    setCargando(true);

    try {
      const data = await archivoExcel.arrayBuffer();
      const workbook = XLSX.read(data);
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const datosExcel = XLSX.utils.sheet_to_json(hoja);

      if (datosExcel.length === 0) throw new Error('Excel vacío');

      const datosParaWord = datosExcel[0];
      const respuesta = await fetch('/plantilla.docx');
      if (!respuesta.ok) throw new Error('Plantilla no encontrada');
      
      const zip = new PizZip(await (await respuesta.blob()).arrayBuffer());
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

      doc.render(datosParaWord);
      const out = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

      saveAs(out, `Hoja_Resumen_${datosParaWord['N°_DOCUMENTO'] || 'Cliente'}.docx`);

    } catch (error) {
      console.error(error);
      alert('Error procesando archivos.');
    } finally {
      setCargando(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 relative flex flex-col items-center justify-center p-6 font-sans overflow-hidden transition-colors duration-500">
      
      {/* BOTÓN MODO OSCURO / CLARO */}
      <button 
        onClick={() => setDarkMode(!darkMode)}
        className="absolute top-6 right-6 p-3 rounded-full bg-white dark:bg-slate-800 text-slate-600 dark:text-yellow-400 shadow-lg border border-slate-200 dark:border-slate-700 hover:scale-110 transition-all duration-300 z-50"
        title="Alternar Modo Oscuro"
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

      {/* Decoración de fondo */}
      <div className="absolute top-0 left-1/4 w-96 h-96 bg-indigo-500/10 dark:bg-indigo-500/20 rounded-full blur-[128px] pointer-events-none transition-colors duration-500" />
      <div className="absolute bottom-0 right-1/4 w-96 h-96 bg-emerald-500/10 rounded-full blur-[128px] pointer-events-none transition-colors duration-500" />

      {/* CONTENEDOR PRINCIPAL */}
      <div className="bg-white dark:bg-slate-900/60 backdrop-blur-xl p-14 rounded-3xl shadow-xl dark:shadow-[0_35px_100px_-15px_rgba(0,0,0,0.3)] max-w-xl w-full border border-slate-200 dark:border-slate-800 transition-all duration-500 ease-out hover:border-indigo-400/50 dark:hover:border-indigo-500/30 hover:scale-[1.01] hover:-translate-y-1 relative group z-10">
        
        <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500 rounded-t-3xl" />
        
        {/* Cabecera */}
        <div className="text-center mb-12">
          <div className="inline-flex items-center justify-center w-22 h-22 rounded-2xl bg-indigo-50 dark:bg-slate-800 border border-indigo-100 dark:border-slate-700 shadow-md mb-8 relative overflow-hidden group-hover:border-indigo-500 transition-colors">
            <div className="absolute inset-0 bg-gradient-to-br from-indigo-600 to-purple-600 opacity-0 group-hover:opacity-10 transition-opacity" />
            <svg className="w-11 h-11 text-indigo-600 dark:text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
            </svg>
          </div>
          <h1 className="text-4xl font-extrabold text-slate-800 dark:text-white tracking-tighter leading-tight transition-colors">
            Automatización de documentos
          </h1>
          <p className="text-slate-600 dark:text-slate-400 mt-4 text-lg max-w-sm mx-auto font-medium leading-relaxed transition-colors">
            Sube tu fuente de datos Excel para autocompletar la Hoja Resumen al instante.
          </p>
        </div>
        
        {/* ÁREA DE SUBIDA */}
        <div 
          className="mb-12 relative group/drop"
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); manejarSubida(e); }}
        >
          <label className={`flex flex-col items-center justify-center w-full h-48 rounded-2xl cursor-pointer transition-all duration-300 relative overflow-hidden group
            ${archivoExcel 
              ? 'bg-emerald-50 dark:bg-emerald-950/30 border-2 border-emerald-400 dark:border-emerald-500/50 shadow-sm' 
              : dragOver
                ? 'bg-indigo-50 dark:bg-indigo-950/50 border-2 border-indigo-400 border-dashed'
                : 'bg-slate-50 dark:bg-slate-800/40 border border-slate-300 dark:border-slate-700 hover:border-indigo-400 dark:hover:border-indigo-500 hover:bg-slate-100 dark:hover:bg-slate-800/80'}`}>
            
            <div className="flex flex-col items-center justify-center pt-8 pb-10 px-6 text-center z-10 relative">
              {archivoExcel ? (
                <>
                  <div className="w-18 h-18 bg-emerald-100 dark:bg-emerald-500/10 rounded-full flex items-center justify-center mb-5 border border-emerald-200 dark:border-emerald-500/30">
                    <svg className="w-9 h-9 text-emerald-600 dark:text-emerald-400 animate-[pulse_2s_infinite]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                  </div>
                  <p className="mb-2 text-base text-slate-800 dark:text-white font-bold tracking-tight truncate max-w-xs">{archivoExcel.name}</p>
                  <p className="text-sm text-emerald-700 dark:text-emerald-300 font-medium bg-emerald-100 dark:bg-emerald-950 px-4 py-1.5 rounded-full border border-emerald-200 dark:border-emerald-500/30">Listo para procesar</p>
                </>
              ) : (
                <>
                  <div className={`w-18 h-18 rounded-full flex items-center justify-center mb-5 transition-all duration-300 border 
                    ${dragOver ? 'bg-indigo-100 dark:bg-indigo-500/20 border-indigo-300 dark:border-indigo-400 animate-pulse' : 'bg-white dark:bg-slate-700 border-slate-200 dark:border-slate-600 shadow-sm'}`}>
                    <svg className={`w-9 h-9 transition-colors ${dragOver ? 'text-indigo-500 dark:text-indigo-300' : 'text-slate-400 dark:text-slate-400 group-hover:text-indigo-500 dark:group-hover:text-indigo-400'}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M12 11v6m3-3l-3 3-3-3" className="animate-[bounce_1.5s_infinite]"/>
                    </svg>
                  </div>
                  <p className="mb-2 text-base text-slate-600 dark:text-slate-300 font-medium leading-relaxed">
                    {dragOver ? '¡Suéltalo aquí!' : <><span className="font-bold text-indigo-600 dark:text-indigo-400">Haz clic para subir</span> o arrastra el archivo</>}
                  </p>
                  <p className="text-sm text-slate-400 dark:text-slate-500 font-medium">Soporta formatos .xlsx y .xls</p>
                </>
              )}
            </div>
            
            <input type="file" accept=".xlsx, .xls" onChange={manejarSubida} className="hidden" />
          </label>
        </div>

        {/* BOTÓN DE ACCIÓN */}
        <button
          onClick={generarDocumento}
          disabled={!archivoExcel || cargando}
          className={`group/btn relative w-full py-4.5 px-7 rounded-2xl font-extrabold text-base uppercase tracking-widest transition-all duration-300 overflow-hidden flex justify-center items-center gap-3.5
            ${!archivoExcel || cargando 
              ? 'bg-slate-200 dark:bg-slate-800 cursor-not-allowed text-slate-500 dark:text-slate-400 border border-slate-300 dark:border-slate-700' 
              : 'bg-gradient-to-r from-indigo-600 via-indigo-500 to-indigo-600 text-white bg-[length:200%_auto] hover:bg-right shadow-lg hover:shadow-xl transform hover:-translate-y-0.5 active:translate-y-0 animate-[pulse_3s_infinite]'
            }`}
        >
          {cargando ? (
            <>
              <svg className="animate-spin h-6 w-6 currentColor" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              <span>Generando documento...</span>
            </>
          ) : archivoExcel ? (
            'Generar Hoja Resumen'
          ) : (
            'Esperando fuente de datos'
          )}
          
          <div className="absolute top-0 -inset-full h-full w-1/2 block transform -skew-x-12 bg-gradient-to-r from-transparent to-white opacity-20 group-hover/btn:animate-shine" />
        </button>
      </div>

      {/* Footer */}
      <div className="mt-18 text-center text-base text-slate-500 dark:text-slate-600 font-semibold tracking-wide border-t border-slate-200 dark:border-slate-800 pt-8 max-w-sm mx-auto transition-colors">
        SMART CASH TOOL
        <span className="text-slate-400 dark:text-slate-700 block text-sm mt-1.5">v1.2 | Procesamiento Local Seguro</span>
      </div>

      <style>{`
        @keyframes shine {
          100% { left: 125%; }
        }
        .animate-shine {
          animation: shine 1s;
        }
      `}</style>
    </div>
  );
}

export default App;
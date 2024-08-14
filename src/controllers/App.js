import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';

// Obtener la ruta del archivo actual
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Función para procesar el archivo Excel y generar el nuevo archivo
export const procesarExcel = (req, res) => {
    try {
        // Ruta del archivo Excel original
        const rutaArchivo = path.resolve(__dirname, '../data/CORRERIA.xlsx');
        const workbook = xlsx.readFile(rutaArchivo);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        // Agrupar por TPL y calcular ciclo mínimo y máximo
        const agrupadosPorTPL = data.reduce((acc, row) => {
            const { TPL, Ciclo } = row;

            if (!acc[TPL]) {
                acc[TPL] = { minCiclo: Ciclo, maxCiclo: Ciclo };
            } else {
                acc[TPL].minCiclo = Math.min(acc[TPL].minCiclo, Ciclo);
                acc[TPL].maxCiclo = Math.max(acc[TPL].maxCiclo, Ciclo);
            }

            return acc;
        }, {});

        // Procesar cada fila
        const datosProcesados = data.map(row => {
            const { Ciclo, TPL, Tarea } = row;

            const { minCiclo, maxCiclo } = agrupadosPorTPL[TPL];

            let tareaCodigo;
            switch (Tarea) {
                case 'RECONEXIÓN':
                    tareaCodigo = '03';
                    break;
                case 'SUSPENSIÓN':
                    tareaCodigo = '02';
                    break;
                case 'DESVIACIÓN SIGNIFICATIVA':
                    tareaCodigo = '01';
                    break;
                default:
                    tareaCodigo = '00';
            }

            const tplNumero = TPL.match(/\d+/)[0];
            const nombreCorreria = `SCR CL ${minCiclo}/${maxCiclo} ${tplNumero}`;

            return {
                ...row,
                'nombre correria': nombreCorreria,
                'Tarea-codigo': tareaCodigo
            };
        });

        // Crear un nuevo libro de trabajo y una hoja
        const newWorkbook = xlsx.utils.book_new();
        const newWorksheet = xlsx.utils.json_to_sheet(datosProcesados);
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Datos Procesados');

        // Guardar el archivo Excel en la carpeta 'file'
        const nuevaRutaArchivo = path.resolve(__dirname, '../file/Datos_Procesados.xlsx');
        xlsx.writeFile(newWorkbook, nuevaRutaArchivo);

        // Responder al cliente
        res.status(200).json({ message: 'Archivo procesado y guardado como Datos_Procesados.xlsx' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: 'Error al procesar el archivo Excel' });
    }
};

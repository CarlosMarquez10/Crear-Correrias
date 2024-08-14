import express from 'express';
import { procesarExcel } from '../controllers/App.js';

const router = express.Router();

// Ruta para procesar el archivo Excel
router.get('/procesar-excel', procesarExcel);

export default router;

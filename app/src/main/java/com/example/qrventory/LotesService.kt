package com.example.qrventory

import android.util.Log

class LotesService(private val driveHelper: DriveServiceHelper) {

    /**
     * Busca la fila que coincide con el código y la variante.
     */
    private fun buscarFilaPorCodigoYVariante(
        spreadsheetId: String,
        hoja: String,
        codigo: String,
        variante: String
    ): Int? {
        return try {
            val rango = "$hoja!A:D"
            val response = driveHelper.getSheetsService().spreadsheets().values()
                .get(spreadsheetId, rango)
                .execute()

            val valores = response.getValues()

            if (valores != null) {
                valores.forEachIndexed { index, fila ->
                    val celdaCodigo = fila.getOrNull(0)?.toString()?.trim()
                    val celdaVariante = fila.getOrNull(2)?.toString()?.trim()

                    if (celdaCodigo == codigo && celdaVariante == variante) {
                        return index + 1
                    }
                }
            }

            null
        } catch (e: Exception) {
            Log.e("LotesService", "Error al buscar fila por código y variante", e)
            null
        }
    }

    /**
     * Procesa el lote actualizando la cantidad solo si coincide el código y variante.
     */
    fun procesarLote(
        spreadsheetId: String,
        codigo: String,
        variante: String,
        columnaDestino: String,
        cantidad: Int
    ): String {
        val hoja = driveHelper.obtenerPrimerNombreHoja(spreadsheetId)
            ?: return "No se pudo obtener la hoja"

        val fila = buscarFilaPorCodigoYVariante(spreadsheetId, hoja, codigo, variante)

        if (fila == null) {
            return "Lote no encontrado con código y variante especificados."
        }

        val celdaDestino = "$columnaDestino$fila"

        val valorActualStr = driveHelper.obtenerValorCelda(spreadsheetId, hoja, celdaDestino)
        val valorActual = valorActualStr?.toIntOrNull() ?: 0
        val nuevoValor = valorActual + cantidad

        return if (driveHelper.actualizarCelda(spreadsheetId, hoja, celdaDestino, nuevoValor.toString())) {
            "Lote actualizado en $celdaDestino con cantidad $cantidad"
        } else {
            "Error al actualizar lote en $celdaDestino"
        }

    }
}

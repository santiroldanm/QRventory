package com.example.qrventory

class SerialesService(private val driveHelper: DriveServiceHelper) {

    fun procesarSerial(
        spreadsheetId: String,
        codigo: String,
        variante: String,
        columnaDestino: String
    ): String {
        val hoja = driveHelper.obtenerNombreSegundaHoja(spreadsheetId)
            ?: return "No se pudo obtener la hoja 2"

        return try {
            val rango = "$hoja!A:D"
            val response = driveHelper.getSheetsService().spreadsheets().values()
                .get(spreadsheetId, rango)
                .execute()

            val valores = response.getValues() ?: return "Hoja vacía"

            val filasCoincidentes = valores.withIndex().filter {
                val fila = it.value
                fila.getOrNull(0)?.toString()?.trim() == codigo.trim()
            }

            if (filasCoincidentes.isEmpty()) return "Código no encontrado"

            for ((index, fila) in filasCoincidentes) {
                val varianteFila = fila.getOrNull(2)?.toString()?.trim() ?: continue
                if (varianteFila == variante.trim()) {
                    val filaIndex   = index + 1
                    val celdaDestino = "$columnaDestino$filaIndex"
                    // ← aquí pasamos variante en lugar de código:
                    val actualizado = driveHelper.actualizarCelda(spreadsheetId, hoja, celdaDestino, variante)
                    return if (actualizado)
                        "Serial (variante) actualizado en $celdaDestino"
                    else
                        "Error al escribir en la celda $celdaDestino"
                }
            }

            "Variante no encontrada para este código"
        } catch (e: Exception) {
            e.printStackTrace()
            "Error al procesar el serial"
        }
    }

}

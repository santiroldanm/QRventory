package com.example.qrventory

import android.content.Context
import android.util.Log
import com.google.android.gms.auth.api.signin.GoogleSignInAccount
import com.google.api.client.extensions.android.http.AndroidHttp
import com.google.api.client.googleapis.extensions.android.gms.auth.GoogleAccountCredential
import com.google.api.client.json.gson.GsonFactory
import com.google.api.services.drive.Drive
import com.google.api.services.drive.model.FileList
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.model.ValueRange

class DriveServiceHelper(context: Context, account: GoogleSignInAccount) {
    private val driveService: Drive
    private val sheetsService: Sheets

    init {
        val scopes = listOf(
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        )
        val credential = GoogleAccountCredential.usingOAuth2(context, scopes)
        credential.selectedAccount = account.account

        driveService = Drive.Builder(
            AndroidHttp.newCompatibleTransport(),
            GsonFactory.getDefaultInstance(),
            credential
        ).setApplicationName("QRventory").build()

        sheetsService = Sheets.Builder(
            AndroidHttp.newCompatibleTransport(),
            GsonFactory.getDefaultInstance(),
            credential
        ).setApplicationName("QRventory").build()
    }

    fun listarArchivosSheets(): List<Pair<String, String>> {
        return try {
            val result: FileList = driveService.files().list()
                .setQ("mimeType='application/vnd.google-apps.spreadsheet' and trashed = false")
                .setFields("files(id, name)")
                .execute()

            result.files.map { it.id to it.name }
        } catch (e: Exception) {
            e.printStackTrace()
            emptyList()
        }
    }

    fun buscarFilaPorCodigo(
        spreadsheetId: String,
        hoja: String,
        columna: String,
        codigo: String
    ): Int? {
        return try {
            val rango = "$hoja!$columna:$columna"
            Log.d("SheetsHelper", "Buscando código en rango: $rango")
            val response = sheetsService.spreadsheets().values()
                .get(spreadsheetId, rango)
                .execute()

            val valores = response.getValues()
            if (valores == null) {
                Log.d("SheetsHelper", "No se encontraron valores en la columna.")
                return null
            }

            valores.forEachIndexed { index, fila ->
                val celda = fila.getOrNull(0)?.toString()?.trim()
                Log.d("SheetsHelper", "Fila ${index + 1}: '$celda'")
                if (celda == codigo.trim()) {
                    Log.d("SheetsHelper", "Código encontrado en fila ${index + 1}")
                    return index + 1
                }
            }

            Log.d("SheetsHelper", "Código '$codigo' no encontrado.")
            null
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    fun actualizarCelda(
        spreadsheetId: String,
        hoja: String,
        celda: String,
        valor: String
    ): Boolean {
        return try {
            val rango = "$hoja!$celda"
            val valueRange = ValueRange()
                .setRange(rango)
                .setValues(listOf(listOf(valor)))

            Log.d("SheetsHelper", "Actualizando celda $rango con valor '$valor'")

            sheetsService.spreadsheets().values()
                .update(spreadsheetId, rango, valueRange)
                .setValueInputOption("USER_ENTERED")
                .execute()

            true
        } catch (e: Exception) {
            e.printStackTrace()
            false
        }
    }

    fun obtenerValorCelda(
        spreadsheetId: String,
        hoja: String,
        celda: String
    ): String? {
        return try {
            val rango = "$hoja!$celda"
            Log.d("SheetsHelper", "Obteniendo valor de celda: $rango")
            val response = sheetsService.spreadsheets().values()
                .get(spreadsheetId, rango)
                .execute()

            val valor = response.getValues()?.firstOrNull()?.firstOrNull()?.toString()
            Log.d("SheetsHelper", "Valor obtenido: $valor")
            valor
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    fun obtenerPrimerNombreHoja(spreadsheetId: String): String? {
        return try {
            val spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute()
            val hojas = spreadsheet.sheets
            val primera = hojas?.firstOrNull()?.properties?.title
            Log.d("SheetsHelper", "Primera hoja: $primera")
            primera
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    fun obtenerNombreSegundaHoja(spreadsheetId: String): String? {
        return try {
            val spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute()
            spreadsheet.sheets?.getOrNull(1)?.properties?.title
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }


    fun buscarProductosPorNombre(spreadsheetId: String, keyword: String): List<Pair<String, String>> {
        val hoja = obtenerPrimerNombreHoja(spreadsheetId) ?: return emptyList()

        return try {
            val rango = "$hoja!A:B"
            val response = sheetsService.spreadsheets().values()
                .get(spreadsheetId, rango)
                .execute()

            val filas = response.getValues() ?: return emptyList()
            filas.drop(1).mapNotNull { fila ->
                val codigo = fila.getOrNull(0)?.toString()?.trim() ?: return@mapNotNull null
                val nombre = fila.getOrNull(1)?.toString()?.trim() ?: return@mapNotNull null
                val regex = Regex(keyword.replace("%", ".*"), RegexOption.IGNORE_CASE)
                if (regex.containsMatchIn(nombre) || regex.containsMatchIn(codigo)) {
                    codigo to nombre
                } else null
            }.distinctBy { it.first } // ← Esta línea es clave
        } catch (e: Exception) {
            e.printStackTrace()
            emptyList()
        }
    }

    fun getSheetsService(): Sheets {
        return sheetsService
    }

    fun buscarProductosPorNombreEnHoja2(spreadsheetId: String, keyword: String): List<Pair<String, String>> {
        val hoja = obtenerNombreSegundaHoja(spreadsheetId) ?: return emptyList()

        return try {
            val rango = "$hoja!A:C" // A: código, B: nombre (si aplica), C: variante
            val response = sheetsService.spreadsheets().values()
                .get(spreadsheetId, rango)
                .execute()

            val filas = response.getValues() ?: return emptyList()
            filas.drop(1).mapNotNull { fila ->
                val codigo = fila.getOrNull(0)?.toString()?.trim() ?: return@mapNotNull null
                val nombre = fila.getOrNull(1)?.toString()?.trim() ?: "Sin nombre"
                val regex = Regex(keyword.replace("%", ".*"), RegexOption.IGNORE_CASE)
                if (regex.containsMatchIn(nombre) || regex.containsMatchIn(codigo)) {
                    codigo to nombre
                } else null
            } .distinctBy { it.first }
        } catch (e: Exception) {
            e.printStackTrace()
            emptyList()
        }
    }

}

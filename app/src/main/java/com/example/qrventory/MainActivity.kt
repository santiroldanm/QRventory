package com.example.qrventory

import android.content.Intent
import android.database.Cursor
import android.graphics.Color
import android.net.Uri
import android.os.Bundle
import android.provider.OpenableColumns
import android.util.Log
import android.view.View
import android.widget.*
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import com.google.android.gms.auth.api.signin.GoogleSignIn
import com.google.android.gms.auth.api.signin.GoogleSignInAccount
import com.google.android.gms.auth.api.signin.GoogleSignInClient
import com.google.android.gms.auth.api.signin.GoogleSignInOptions
import com.google.android.gms.common.api.Scope
import com.google.android.gms.tasks.Task
import com.google.api.services.drive.DriveScopes
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream


class MainActivity : AppCompatActivity() {

    companion object {
        private const val KEY_URI = "key_excel_uri"
        private const val KEY_NAME = "key_excel_name"
        private const val RC_SIGN_IN = 9001
    }

    private lateinit var btnArchivoLocal: Button
    private lateinit var btnArchivoDrive: Button
    private lateinit var btnConectarDrive: Button
    private lateinit var btnLeerQR: Button
    private lateinit var btnActualizar: Button
    private lateinit var textRutaArchivo: TextView
    private lateinit var editCodigo: EditText
    private lateinit var editCantidad: EditText
    private lateinit var editColumna: EditText
    private lateinit var textEstado: TextView
    private lateinit var btnBuscarNombre: Button
    private lateinit var editNombreProducto: EditText
    private lateinit var editCodigoSerial: EditText
    private lateinit var editVarianteSerial: EditText
    private lateinit var editColumnaSerial: EditText
    private lateinit var btnModoSeriales: Button
    private lateinit var btnActualizarSerial : Button
    private lateinit var editCodigoLote: EditText
    private lateinit var editVarianteLote: EditText
    private lateinit var editColumnaLote: EditText
    private lateinit var editCantidadLote: EditText
    private lateinit var btnActualizarLote: Button
    private lateinit var btnModoLotes: Button
    private lateinit var lotesService: LotesService
    private lateinit var layoutLotes: LinearLayout



    private var modoSerialesActivo = false
    private var modoLotesActivo: Boolean = false

    private var archivoUri: Uri? = null
    private var nombreArchivo: String = ""
    private var archivoIdDrive: String? = null

    private lateinit var googleSignInClient: GoogleSignInClient
    private var cuentaGoogle: GoogleSignInAccount? = null
    private lateinit var driveServiceHelper: DriveServiceHelper

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        initUI()
        restoreFileState(savedInstanceState)
        setupGoogleSignIn()
        setupListeners()
    }

    private fun initUI() {
        btnLeerQR = findViewById(R.id.btnLeerQR)
        btnActualizar = findViewById(R.id.btnActualizar)
        textRutaArchivo = findViewById(R.id.textRutaArchivo)
        editCodigo = findViewById(R.id.editCodigo)
        editCantidad = findViewById(R.id.editCantidad)
        editColumna = findViewById(R.id.editColumna)
        textEstado = findViewById(R.id.textEstado)
        btnArchivoLocal = findViewById(R.id.btnArchivoLocal)
        btnArchivoDrive = findViewById(R.id.btnArchivoDrive)
        btnConectarDrive = findViewById(R.id.btnConectarDrive)
        btnBuscarNombre = findViewById(R.id.btnBuscarNombre)
        btnModoSeriales = findViewById(R.id.btnModoSeriales)
        editNombreProducto = findViewById(R.id.editNombreProducto)
        editCodigoSerial = findViewById(R.id.editCodigoSerial)
        editVarianteSerial = findViewById(R.id.editVarianteSerial)
        editColumnaSerial = findViewById(R.id.editColumnaSerial)
        btnActualizarSerial = findViewById<Button>(R.id.btnActualizarSerial)
        editCodigoLote = findViewById(R.id.editCodigoLote)
        editVarianteLote = findViewById(R.id.editVarianteLote)
        editColumnaLote = findViewById(R.id.editColumnaLote)
        editCantidadLote = findViewById(R.id.editCantidadLote)
        btnActualizarLote = findViewById(R.id.btnActualizarLote)
        btnModoLotes = findViewById(R.id.btnModoLotes)
        layoutLotes = findViewById(R.id.layoutLotes)



    }

    private fun restoreFileState(savedInstanceState: Bundle?) {
        savedInstanceState?.getString(KEY_URI)?.let {
            archivoUri = Uri.parse(it)
            nombreArchivo = savedInstanceState.getString(KEY_NAME, "")
            textRutaArchivo.text = nombreArchivo
        }
        archivoIdDrive = savedInstanceState?.getString("KEY_DRIVE_ID")
        archivoIdDrive?.let {
            textRutaArchivo.text = "Drive: $nombreArchivo"
        }
    }

    private fun setupGoogleSignIn() {
        val opcionesLogin = GoogleSignInOptions.Builder(GoogleSignInOptions.DEFAULT_SIGN_IN)
            .requestEmail()
            .requestScopes(Scope(DriveScopes.DRIVE), Scope("https://www.googleapis.com/auth/spreadsheets"))
            .build()

        googleSignInClient = GoogleSignIn.getClient(this, opcionesLogin)
        cuentaGoogle = GoogleSignIn.getLastSignedInAccount(this)

        if (cuentaGoogle == null) {
            startActivityForResult(googleSignInClient.signInIntent, RC_SIGN_IN)
        } else {
            Toast.makeText(this, "Sesión con Google activa", Toast.LENGTH_SHORT).show()
            setupDriveService()

        }

        findViewById<Button>(R.id.btnConectarDrive).setOnClickListener {
            startActivityForResult(googleSignInClient.signInIntent, RC_SIGN_IN)
        }
    }

    private fun setupListeners() {
        btnConectarDrive.setOnClickListener {
            googleSignInClient.signOut().addOnCompleteListener {
                startActivityForResult(googleSignInClient.signInIntent, RC_SIGN_IN)
            }
        }

        btnArchivoLocal.setOnClickListener {
            archivoPicker.launch(arrayOf("*/*"))
        }

        btnArchivoDrive.setOnClickListener {
            if (cuentaGoogle != null) mostrarSelectorDeArchivosDrive()
            else Toast.makeText(this, "Primero inicia sesión con Google", Toast.LENGTH_SHORT).show()
        }

        btnLeerQR.setOnClickListener {
            if (archivoUri == null && archivoIdDrive == null) {
                textEstado.text = "Primero selecciona un archivo (local o de Drive)."
            } else {
                startActivityForResult(Intent(this, QrScannerActivity::class.java), 200)
            }
        }

        btnBuscarNombre.setOnClickListener {
            val keyword = editNombreProducto.text.toString().trim()

            if (keyword.isBlank()) {
                Toast.makeText(this, "Ingresa una palabra clave", Toast.LENGTH_SHORT).show()
                return@setOnClickListener
            }

            textEstado.text = "Buscando coincidencias..."

            Thread {
                val resultados = when {
                    archivoIdDrive != null -> {
                        if (modoSerialesActivo) {
                            driveServiceHelper.buscarProductosPorNombreEnHoja2(archivoIdDrive!!, keyword)
                        } else {
                            driveServiceHelper.buscarProductosPorNombre(archivoIdDrive!!, keyword)
                        }
                    }
                    archivoUri != null -> {
                        buscarProductosPorNombreLocal(archivoUri!!, keyword) // Esta funcion solo busca en hoja 1
                    }
                    else -> emptyList()
                }

                runOnUiThread {
                    if (resultados.isEmpty()) {
                        textEstado.text = "No se encontraron coincidencias."
                    } else {
                        val nombres = resultados.map { "${it.second} [${it.first}]" }.toTypedArray()
                        AlertDialog.Builder(this)
                            .setTitle("Selecciona un producto")
                            .setItems(nombres) { _, which ->
                                val (codigo, _) = resultados[which]

                                if (modoSerialesActivo) {
                                    editCodigoSerial.setText(codigo)
                                } else if (modoLotesActivo) {
                                    editCodigoLote.setText(codigo)
                                } else {
                                    editCodigo.setText(codigo)
                                }

                                // En todos los casos:
                                editNombreProducto.setText(resultados[which].second)


                                textEstado.text = "Producto seleccionado"
                            }
                            .setNegativeButton("Cancelar", null)
                            .show()
                    }
                }
            }.start()
        }

        btnModoSeriales.setOnClickListener {
            if (modoLotesActivo) {
                Toast.makeText(this, "Sal de Modo Lotes para activar Seriales", Toast.LENGTH_SHORT).show()
                return@setOnClickListener
            }

            modoSerialesActivo = !modoSerialesActivo

            val todosLosBotones = listOf(
                btnModoSeriales,
                btnActualizar,
                btnActualizarSerial,
                btnActualizarLote,
                btnLeerQR,
                btnConectarDrive,
                btnArchivoLocal,
                btnArchivoDrive,
                btnBuscarNombre
            )

            if (modoSerialesActivo) {
                val serialColor = Color.parseColor("#FF6200EE")
                todosLosBotones.forEach { it.setBackgroundColor(serialColor) }

                btnModoSeriales.text = "Modo Normal"
                btnModoLotes.visibility = View.GONE

                // Mostrar solo campos serial
                editCodigo.visibility        = View.GONE
                editCantidad.visibility      = View.GONE
                editColumna.visibility       = View.GONE

                editCodigoSerial.visibility  = View.VISIBLE
                editVarianteSerial.visibility= View.VISIBLE
                editColumnaSerial.visibility = View.VISIBLE

                layoutLotes.visibility       = View.GONE

                btnActualizar.visibility      = View.GONE
                btnActualizarSerial.visibility= View.VISIBLE
                btnActualizarLote.visibility  = View.GONE

                Toast.makeText(this, "Modo Seriales activado", Toast.LENGTH_SHORT).show()
            } else {
                val defaultColor = Color.parseColor("#169976")
                todosLosBotones.forEach { it.setBackgroundColor(defaultColor) }

                btnModoSeriales.text = "Modo Seriales"
                btnModoLotes.visibility = View.VISIBLE

                // Mostrar campos normales
                editCodigo.visibility        = View.VISIBLE
                editCantidad.visibility      = View.VISIBLE
                editColumna.visibility       = View.VISIBLE

                editCodigoSerial.visibility  = View.GONE
                editVarianteSerial.visibility= View.GONE
                editColumnaSerial.visibility = View.GONE

                layoutLotes.visibility       = View.GONE

                btnActualizar.visibility      = View.VISIBLE
                btnActualizarSerial.visibility= View.GONE
                btnActualizarLote.visibility  = View.GONE

                Toast.makeText(this, "Modo Seriales desactivado", Toast.LENGTH_SHORT).show()
            }
        }

        btnModoLotes.setOnClickListener {
            if (modoSerialesActivo) {
                Toast.makeText(this, "Sal de Modo Seriales para activar Lotes", Toast.LENGTH_SHORT).show()
                return@setOnClickListener
            }

            modoLotesActivo = !modoLotesActivo

            val todosLosBotones = listOf(
                btnModoLotes,
                btnActualizar,
                btnActualizarSerial,
                btnActualizarLote,
                btnLeerQR,
                btnConectarDrive,
                btnArchivoLocal,
                btnArchivoDrive,
                btnBuscarNombre
            )

            if (modoLotesActivo) {
                val loteColor = Color.parseColor("#C45F06")
                todosLosBotones.forEach { it.setBackgroundColor(loteColor) }

                btnModoLotes.text = "Modo Normal"
                btnModoSeriales.visibility = View.GONE

                // Mostrar campos lote
                editCodigo.visibility        = View.GONE
                editCantidad.visibility      = View.GONE
                editColumna.visibility       = View.GONE

                editCodigoSerial.visibility  = View.GONE
                editVarianteSerial.visibility= View.GONE
                editColumnaSerial.visibility = View.GONE

                layoutLotes.visibility       = View.VISIBLE
                editColumnaLote.visibility = View.VISIBLE

                btnActualizar.visibility      = View.GONE
                btnActualizarSerial.visibility= View.GONE
                btnActualizarLote.visibility  = View.VISIBLE

                Toast.makeText(this, "Modo Lotes activado", Toast.LENGTH_SHORT).show()
            } else {
                val defaultColor = Color.parseColor("#169976")
                todosLosBotones.forEach { it.setBackgroundColor(defaultColor) }

                btnModoLotes.text = "Modo Lotes"
                btnModoSeriales.visibility = View.VISIBLE

                // Mostrar campos normales
                editCodigo.visibility        = View.VISIBLE
                editCantidad.visibility      = View.VISIBLE
                editColumna.visibility       = View.VISIBLE

                editCodigoSerial.visibility  = View.GONE
                editVarianteSerial.visibility= View.GONE
                editColumnaSerial.visibility = View.GONE

                layoutLotes.visibility       = View.GONE
                editColumnaLote.visibility = View.GONE

                btnActualizar.visibility      = View.VISIBLE
                btnActualizarSerial.visibility= View.GONE
                btnActualizarLote.visibility  = View.GONE

                Toast.makeText(this, "Modo Lotes desactivado", Toast.LENGTH_SHORT).show()
            }
        }





        btnActualizar.setOnClickListener {
            val codigo = editCodigo.text.toString().trim().replace(Regex("[\\r\\n]"), "")
            val cantidad = editCantidad.text.toString().toIntOrNull()
            val columna = editColumna.text.toString().trim()

            if ((archivoUri == null && archivoIdDrive == null) || codigo.isBlank() || cantidad == null || columna.isBlank()) {
                textEstado.text = "Todos los campos deben estar completos y válidos."
                return@setOnClickListener
            }

            btnActualizar.isEnabled = false
            textEstado.text = "Actualizando..."

            Thread {
                val esDrive = archivoIdDrive != null
                val (dummyResult, nombreProducto) = if (esDrive) {
                    actualizarExcelDrive(archivoIdDrive!!, codigo, 0, columna)
                } else {
                    actualizarExcel(archivoUri!!, codigo, 0, columna)
                }

                fun aplicarActualizacion(finalCantidad: Int) {
                    Thread {
                        val (resultado, nombreFinal) = if (esDrive) {
                            actualizarExcelDrive(archivoIdDrive!!, codigo, finalCantidad, columna)
                        } else {
                            actualizarExcel(archivoUri!!, codigo, finalCantidad, columna)
                        }

                        runOnUiThread {
                            btnActualizar.isEnabled = true
                            if (resultado == "Actualización exitosa") {
                                // Limpiar los campos
                                editNombreProducto.text.clear()
                                editCodigo.text.clear()
                                editCodigo.requestFocus()
                                editCantidad.setText("") // Opcional, puedes comentar esta línea si quieres conservar la cantidad anterior

                                textEstado.text = "Producto $nombreFinal actualizado con $finalCantidad unidades."
                            } else {
                                textEstado.text = resultado
                            }
                        }
                    }.start()
                }



                runOnUiThread {
                    val nombreMayus = nombreProducto.uppercase()

                    val requiereMultiplicador = nombreMayus.startsWith("ETIQUETA") ||
                            nombreMayus.startsWith("GONDOLA") ||
                            nombreMayus.startsWith("TUFFMARK") ||
                            nombreMayus.startsWith("POLIESTER PLATA")||
                            nombreMayus.startsWith("CARTON GRAFA")

                    if (requiereMultiplicador) {
                        mostrarDialogoMultiplicador(cantidad!!) { nuevaCantidad ->
                            if (nuevaCantidad != null) {
                                aplicarActualizacion(nuevaCantidad)
                            } else {
                                // Reactivar botón o restaurar estado si se canceló
                                reactivarBotonActualizar() // Esta función debes implementarla si necesitas reactivar algo
                            }
                        }
                    } else {
                        aplicarActualizacion(cantidad!!)
                    }
                }

            }.start()

        }

        editCodigo.setOnKeyListener { _, keyCode, event ->
            if (keyCode == android.view.KeyEvent.KEYCODE_ENTER && event.action == android.view.KeyEvent.ACTION_UP) {
                editCantidad.requestFocus()
                return@setOnKeyListener true
            }
            false
        }

        btnActualizarSerial.setOnClickListener {
            val serial = editCodigoSerial.text.toString().trim()
            val variante = editVarianteSerial.text.toString().trim()
            val columna = editColumnaSerial.text.toString().trim()

            if (serial.isEmpty() || variante.isEmpty() || columna.isEmpty()) {
                textEstado.text = "Todos los campos de serial deben estar completos."
                return@setOnClickListener
            }

            textEstado.text = "Procesando serial..."

            Thread {
                val resultado = when {
                    archivoIdDrive != null -> {
                        val helper = SerialesService(driveServiceHelper)
                        helper.procesarSerial(archivoIdDrive!!, serial, variante, columna)
                    }
                    archivoUri != null -> {
                        procesarSerialLocal(archivoUri!!, serial, variante, columna)
                    }
                    else -> "No se ha seleccionado un archivo"
                }

                runOnUiThread {
                    textEstado.text = resultado
                    if (resultado.startsWith("Serial (variante) actualizado")) {
                        editNombreProducto.text.clear()
                        editCodigoSerial.text.clear()
                        editVarianteSerial.text.clear()
                        editCodigoSerial.requestFocus()
                    }
                }
            }.start()
        }

        editCodigoSerial.setOnKeyListener { _, keyCode, event ->
            if (keyCode == android.view.KeyEvent.KEYCODE_ENTER && event.action == android.view.KeyEvent.ACTION_UP) {
                editVarianteSerial.requestFocus()
                return@setOnKeyListener true
            }
            false
        }

        btnActualizarLote.setOnClickListener {
            val codigo = editCodigoLote.text.toString().trim()
            val variante = editVarianteLote.text.toString().trim()
            val columna = editColumnaLote.text.toString().trim()
            val cantidad = editCantidadLote.text.toString().toIntOrNull()

            if (codigo.isEmpty() || variante.isEmpty() || columna.isEmpty() || cantidad == null) {
                textEstado.text = "Todos los campos deben estar completos."
                return@setOnClickListener
            }

            textEstado.text = "Procesando lote..."

            Thread {
                val resultado = when {
                    archivoIdDrive != null -> {
                        lotesService.procesarLote(archivoIdDrive!!, codigo, variante, columna, cantidad)
                    }
                    archivoUri != null -> {
                        procesarLoteLocal(archivoUri!!, codigo, variante, columna, cantidad)
                    }
                    else -> "No se ha seleccionado un archivo"
                }

                runOnUiThread {
                    textEstado.text = resultado
                    if (resultado.startsWith("Lote actualizado")) {
                        editNombreProducto.text.clear()
                        editCodigoLote.text.clear()
                        editVarianteLote.text.clear()
                        editCantidadLote.setText("")
                        editNombreProducto.requestFocus()
                    }

                }
            }.start()
        }

        editCodigoLote.setOnKeyListener { _, keyCode, event ->
            if (keyCode == android.view.KeyEvent.KEYCODE_ENTER && event.action == android.view.KeyEvent.ACTION_UP) {
                editCantidadLote.requestFocus()
                return@setOnKeyListener true
            }
            false
        }

    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        archivoUri?.let {
            outState.putString(KEY_URI, it.toString())
            outState.putString(KEY_NAME, nombreArchivo)
        }
        archivoIdDrive?.let {
            outState.putString("KEY_DRIVE_ID", it)
        }
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode == 200 && resultCode == RESULT_OK) {
            val qr = data?.getStringExtra("qr_result")

            // Si estamos en modo seriales, ponemos el QR en editCodigoSerial
            if (modoSerialesActivo) {
                editCodigoSerial.setText(qr)
            } else {
                // En modo normal, lo ponemos en editCodigo
                editCodigo.setText(qr)
            }

            if (modoLotesActivo) {
                editCodigoLote.setText(qr)
            }
        }

        if (requestCode == RC_SIGN_IN) {
            val task: Task<GoogleSignInAccount> = GoogleSignIn.getSignedInAccountFromIntent(data)
            if (task.isSuccessful) {
                cuentaGoogle = task.result
                Toast.makeText(this, "Login exitoso: ${cuentaGoogle?.email}", Toast.LENGTH_SHORT).show()
                setupDriveService()
            } else {
                Toast.makeText(this, "Fallo en el login con Google", Toast.LENGTH_SHORT).show()
            }
        }
    }


    private fun setupDriveService() {
        cuentaGoogle?.let {
            driveServiceHelper = DriveServiceHelper(this, it)
        }
        lotesService = LotesService(driveServiceHelper)
    }

    private val archivoPicker = registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri: Uri? ->
        uri?.let {
            try {
                contentResolver.takePersistableUriPermission(
                    uri,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                )
                archivoUri = uri
                nombreArchivo = obtenerNombreArchivo(uri)
                textRutaArchivo.text = nombreArchivo
                archivoIdDrive = null
            } catch (e: Exception) {
                Log.e("MainActivity", "Error al tomar permiso", e)
                textRutaArchivo.text = "Error al seleccionar archivo"
            }
        } ?: run {
            textRutaArchivo.text = "No se seleccionó ningún archivo"
        }
    }

    private fun obtenerNombreArchivo(uri: Uri): String {
        var nombre = "Archivo seleccionado"
        val cursor: Cursor? = contentResolver.query(uri, null, null, null, null)
        cursor?.use {
            if (it.moveToFirst()) {
                val index = it.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                if (index >= 0) {
                    nombre = it.getString(index)
                }
            }
        }
        return nombre
    }

    private fun buscarProductosPorNombreLocal(uri: Uri, keyword: String): List<Pair<String, String>> {
        val resultados = mutableListOf<Pair<String, String>>()
        try {
            val inputStream = contentResolver.openInputStream(uri)
            val workbook = XSSFWorkbook(inputStream)

            // Seleccionamos la hoja correcta según el modo
            val sheet = if (modoSerialesActivo) {
                workbook.getSheetAt(1) // Hoja 2
            } else {
                workbook.getSheetAt(0) // Hoja 1
            }

            for (i in 1..sheet.lastRowNum) {
                val row = sheet.getRow(i) ?: continue
                val codigo = row.getCell(0)?.toString()?.trim() ?: continue
                val nombre = row.getCell(1)?.toString()?.trim() ?: continue
                val regex = Regex(keyword.replace("%", ".*"), RegexOption.IGNORE_CASE)
                if (regex.containsMatchIn(nombre) || regex.containsMatchIn(codigo)) {
                    resultados.add(codigo to nombre)
                }
            }

            workbook.close()
            inputStream?.close()
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return resultados
    }

    // --- Lógica de actualización Excel con manejo de tipos de celda (archivo local) ---
    private fun actualizarExcel(
        uri: Uri,
        codigo: String,
        cantidad: Int,
        columnaLetra: String
    ): Pair<String, String> = try {
        val inputStream: InputStream? =
            contentResolver.openInputStream(uri)
        val workbook = XSSFWorkbook(inputStream)
        val sheet = workbook.getSheetAt(0)
        val colIndex = columnaLetra.uppercase()[0] - 'A'

        var codigoEncontrado = false
        var nombreProducto = ""

        // Recorremos filas (empezando en 1 para saltar encabezado)
        for (i in 1..sheet.lastRowNum) {
            val row = sheet.getRow(i) ?: continue
            val cellCodigo = row.getCell(0) ?: continue

            val valorCelda = when (cellCodigo.cellType) {
                CellType.STRING -> cellCodigo.stringCellValue.trim()
                CellType.NUMERIC -> cellCodigo.numericCellValue
                    .toInt().toString().trim()
                else -> cellCodigo.toString().trim()
            }

            if (valorCelda == codigo) {
                val cellCantidad = row.getCell(colIndex)
                    ?: row.createCell(colIndex)
                val actual = try {
                    cellCantidad.numericCellValue.toInt()
                } catch (_: Exception) {
                    0
                }
                cellCantidad.setCellValue((actual + cantidad).toDouble())

                nombreProducto = row.getCell(1)?.toString() ?: ""
                codigoEncontrado = true
                break
            }
        }

        inputStream?.close()

        // Guardar cambios
        contentResolver.openOutputStream(uri)?.use { out ->
            workbook.write(out)
        }
        workbook.close()

        if (codigoEncontrado) "Actualización exitosa" to nombreProducto
        else "Código no encontrado" to ""
    } catch (e: Exception) {
        Log.e("MainActivity", "Error al actualizar Excel", e)
        "Error: No se pudo actualizar el archivo." to ""
    }

    private fun procesarSerialLocal(uri: Uri, serial: String, variante: String, columnaDestino: String): String {
        return try {
            val inputStream = contentResolver.openInputStream(uri)
            val workbook = XSSFWorkbook(inputStream)
            val sheet = workbook.getSheetAt(1) // Hoja 2

            val colIndexDestino = columnaDestino.uppercase()[0] - 'A'
            var celdaActualizada = ""

            for (i in 1..sheet.lastRowNum) {
                val row = sheet.getRow(i) ?: continue
                val celdaCodigo = row.getCell(0)?.toString()?.trim()
                val celdaVariante = row.getCell(2)?.toString()?.trim()

                if (celdaCodigo == serial && celdaVariante.equals(variante, ignoreCase = true)) {
                    val cell = row.getCell(colIndexDestino) ?: row.createCell(colIndexDestino)
                    cell.setCellValue(variante)
                    celdaActualizada = "${columnaDestino.uppercase()}${i + 1}"
                    break
                }
            }

            inputStream?.close()
            if (celdaActualizada.isNotEmpty()) {
                contentResolver.openOutputStream(uri)?.use { workbook.write(it) }
                workbook.close()
                "Serial (variante) actualizado en $celdaActualizada"
            } else {
                workbook.close()
                "Variante no encontrada para este código"
            }
        } catch (e: Exception) {
            e.printStackTrace()
            "Error al procesar serial local"
        }
    }

    private fun procesarLoteLocal(uri: Uri, codigo: String, variante: String, columnaDestino: String, cantidad: Int): String {
        return try {
            val inputStream = contentResolver.openInputStream(uri)
            val workbook = XSSFWorkbook(inputStream)
            val sheet = workbook.getSheetAt(1) // Hoja 2 (para lotes)

            val colIndexDestino = columnaDestino.uppercase()[0] - 'A'
            var celdaActualizada = ""

            for (i in 1..sheet.lastRowNum) {
                val row = sheet.getRow(i) ?: continue
                val celdaCodigo = row.getCell(0)?.toString()?.trim()
                val celdaVariante = row.getCell(2)?.toString()?.trim()

                if (celdaCodigo == codigo && celdaVariante.equals(variante, ignoreCase = true)) {
                    val cell = row.getCell(colIndexDestino) ?: row.createCell(colIndexDestino)
                    val valorActual = try {
                        cell.numericCellValue.toInt()
                    } catch (_: Exception) {
                        0
                    }
                    cell.setCellValue((valorActual + cantidad).toDouble())
                    celdaActualizada = "${columnaDestino.uppercase()}${i + 1}"
                    break
                }
            }

            inputStream?.close()
            if (celdaActualizada.isNotEmpty()) {
                contentResolver.openOutputStream(uri)?.use { workbook.write(it) }
                workbook.close()
                "Lote actualizado en $celdaActualizada con cantidad $cantidad"
            } else {
                workbook.close()
                "Lote no encontrado con código y variante especificados."
            }
        } catch (e: Exception) {
            e.printStackTrace()
            "Error al procesar lote local"
        }
    }


    private fun mostrarSelectorDeArchivosDrive() {
        Thread {
            val archivos = driveServiceHelper.listarArchivosSheets()

            if (archivos.isEmpty()) {
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "No se encontraron archivos Excel en Drive", Toast.LENGTH_LONG).show()
                }
                return@Thread
            }

            val nombres = archivos.map { it.second }.toTypedArray()
            val ids = archivos.map { it.first }

            runOnUiThread {
                AlertDialog.Builder(this@MainActivity)
                    .setTitle("Selecciona un archivo de Drive")
                    .setItems(nombres) { _, which ->
                        // Guardar el ID del archivo seleccionado
                        archivoIdDrive = ids[which]
                        nombreArchivo = nombres[which]
                        textRutaArchivo.text = "Drive: $nombreArchivo"
                        Toast.makeText(this@MainActivity, "Archivo seleccionado", Toast.LENGTH_SHORT).show()
                    }
                    .setNegativeButton("Cancelar", null)
                    .show()
            }
        }.start()
    }

    // --- Lógica de actualización Excel en Google Drive ---
    private fun actualizarExcelDrive(
        spreadsheetId: String,
        codigo: String,
        cantidad: Int,
        columnaLetra: String
    ): Pair<String, String> {
        val hoja = driveServiceHelper.obtenerPrimerNombreHoja(spreadsheetId)
            ?: return "Error: No se pudo obtener el nombre de la hoja" to ""

        return try {
            // Buscar la fila que contiene el código
            val fila = driveServiceHelper.buscarFilaPorCodigo(spreadsheetId, hoja, "A", codigo)
                ?: return "Código no encontrado" to ""

            val celda = "$columnaLetra$fila"

            // Obtener valor actual de la celda
            val valorActualStr = driveServiceHelper.obtenerValorCelda(spreadsheetId, hoja, celda)
            val valorActual = valorActualStr?.toIntOrNull() ?: 0

            // Sumar la nueva cantidad
            val nuevoValor = valorActual + cantidad

            // Actualizar la celda con el nuevo valor
            val exito = driveServiceHelper.actualizarCelda(spreadsheetId, hoja, celda, nuevoValor.toString())

            if (exito) {
                val nombreCelda = "B$fila"
                val nombreProducto = driveServiceHelper.obtenerValorCelda(spreadsheetId, hoja, nombreCelda) ?: ""
                "Actualización exitosa" to nombreProducto
            } else {
                "Error al actualizar celda" to ""
            }

        } catch (e: Exception) {
            Log.e("MainActivity", "Error al usar Sheets API", e)
            "Error inesperado" to ""
        }
    }

    private fun mostrarDialogoMultiplicador(
        cantidadOriginal: Int,
        callback: (multiplicada: Int?) -> Unit
    ) {
        val input = EditText(this).apply {
            hint = "Ej. 1000"
            inputType = android.text.InputType.TYPE_CLASS_NUMBER
        }

        AlertDialog.Builder(this)
            .setTitle("Producto Multiplicable")
            .setMessage("¿Por cuánto vas a multiplicarlo?")
            .setView(input)
            .setPositiveButton("Aceptar") { _, _ ->
                val multiplicador = input.text.toString().toIntOrNull()
                if (multiplicador != null && multiplicador > 0) {
                    callback(cantidadOriginal * multiplicador)
                } else {
                    Toast.makeText(this, "Multiplicador inválido, se usará cantidad original", Toast.LENGTH_SHORT).show()
                    callback(cantidadOriginal)
                }
            }
            .setNegativeButton("Cancelar") { _, _ ->
                callback(null) // ← indicamos que se canceló
            }
            .show()
    }

    private fun reactivarBotonActualizar() {
        btnActualizar.isEnabled = true
        btnActualizar.text = "Actualizar"
        textEstado.text = "Actualización cancelada"
    }
}